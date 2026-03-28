"""Route handlers for loose language tie-out."""
import os
from fastapi import APIRouter, Request, UploadFile, File, Form, Depends
from fastapi.responses import RedirectResponse, FileResponse, JSONResponse
from fastapi.templating import Jinja2Templates
from sqlalchemy.ext.asyncio import AsyncSession
from sqlalchemy import select

from app.config import TEMPLATES_DIR, UPLOAD_DIR, OUTPUT_DIR
from app.database import get_db
from app.models import Deal, LooseComparison, LooseComparisonItem, SentenceTemplate
from app.services.sentence_matcher import extract_document_text, find_matching_sentences
from app.services.excel_reader import get_sheet_names, read_sheet_data
from app.services.loose_comparator import extract_and_compare, build_loose_output_data
from app.services.output_builder import build_output_workbook

router = APIRouter(prefix="/loose")
templates = Jinja2Templates(directory=TEMPLATES_DIR)


# ── Step 1: Enter candidate sentences ─────────────────────────────────────────

@router.get("/sentences/{deal_id}")
async def sentences_form(
    request: Request,
    deal_id: int,
    db: AsyncSession = Depends(get_db),
):
    deal = await db.get(Deal, deal_id)
    if not deal:
        return RedirectResponse(url="/", status_code=303)

    # Load saved templates
    result = await db.execute(
        select(SentenceTemplate).order_by(SentenceTemplate.name)
    )
    saved_templates = result.scalars().all()

    # Find the Word filename from any existing comparison for this deal
    from app.models import Comparison
    comp_result = await db.execute(
        select(Comparison).where(Comparison.deal_id == deal_id).limit(1)
    )
    comparison = comp_result.scalar_one_or_none()
    word_filename = comparison.word_filename if comparison else None

    return templates.TemplateResponse(request, "loose/sentences.html", {
        "deal_id": deal_id,
        "deal_name": deal.name,
        "word_filename": word_filename,
        "saved_templates": saved_templates,
    })


@router.post("/sentences/{deal_id}")
async def submit_sentences(
    request: Request,
    deal_id: int,
    db: AsyncSession = Depends(get_db),
):
    deal = await db.get(Deal, deal_id)
    if not deal:
        return RedirectResponse(url="/", status_code=303)

    form = await request.form()
    sentences_text = form.get("sentences", "").strip()

    if not sentences_text:
        return templates.TemplateResponse(request, "loose/sentences.html", {
            "deal_id": deal_id,
            "deal_name": deal.name,
            "saved_templates": [],
            "error": "Please enter at least one candidate sentence.",
        })

    # Parse sentences (one per line, skip blanks)
    candidate_sentences = [s.strip() for s in sentences_text.split("\n") if s.strip()]

    if not candidate_sentences:
        return templates.TemplateResponse(request, "loose/sentences.html", {
            "deal_id": deal_id,
            "deal_name": deal.name,
            "saved_templates": [],
            "error": "Please enter at least one candidate sentence.",
        })

    # Find the Word file for this deal
    from app.models import Comparison
    comp_result = await db.execute(
        select(Comparison).where(Comparison.deal_id == deal_id).limit(1)
    )
    comparison = comp_result.scalar_one_or_none()
    if not comparison:
        return templates.TemplateResponse(request, "loose/sentences.html", {
            "deal_id": deal_id,
            "deal_name": deal.name,
            "saved_templates": [],
            "error": "No Word document found for this deal. Please upload one first.",
        })

    word_path = os.path.join(UPLOAD_DIR, str(deal_id), comparison.word_filename)

    # Extract document text and find matches
    document_text = extract_document_text(word_path)
    matches = find_matching_sentences(document_text, candidate_sentences, threshold=0.75)

    # Create LooseComparison
    loose_comp = LooseComparison(
        deal_id=deal_id,
        word_filename=comparison.word_filename,
        candidate_sentences=candidate_sentences,
        document_text=document_text,
    )
    db.add(loose_comp)
    await db.flush()

    # Create items for each candidate
    for match in matches:
        item = LooseComparisonItem(
            loose_comparison_id=loose_comp.id,
            candidate_sentence=match["candidate"],
            matched_sentence=match["matched_sentence"],
            similarity_score=match["similarity"],
            status="no_match" if not match["matched_sentence"] else None,
        )
        db.add(item)

    await db.commit()

    return RedirectResponse(url=f"/loose/matches/{loose_comp.id}", status_code=303)


# ── Save / Load templates ─────────────────────────────────────────────────────

@router.post("/save-template")
async def save_template(
    request: Request,
    db: AsyncSession = Depends(get_db),
):
    form = await request.form()
    template_name = form.get("template_name", "").strip()
    sentences_text = form.get("sentences", "").strip()
    deal_id = form.get("deal_id")

    if not template_name or not sentences_text:
        return RedirectResponse(url=f"/loose/sentences/{deal_id}", status_code=303)

    sentences = [s.strip() for s in sentences_text.split("\n") if s.strip()]

    template = SentenceTemplate(
        name=template_name,
        sentences=sentences,
    )
    db.add(template)
    await db.commit()

    return RedirectResponse(url=f"/loose/sentences/{deal_id}", status_code=303)


@router.get("/load-template/{template_id}")
async def load_template(
    template_id: int,
    db: AsyncSession = Depends(get_db),
):
    template = await db.get(SentenceTemplate, template_id)
    if not template:
        return JSONResponse({"sentences": []})
    return JSONResponse({"sentences": template.sentences, "name": template.name})


# ── Step 2: Review matches ────────────────────────────────────────────────────

@router.get("/matches/{loose_id}")
async def review_matches(
    request: Request,
    loose_id: int,
    db: AsyncSession = Depends(get_db),
):
    loose_comp = await db.get(LooseComparison, loose_id)
    if not loose_comp:
        return RedirectResponse(url="/", status_code=303)

    deal = await db.get(Deal, loose_comp.deal_id)

    result = await db.execute(
        select(LooseComparisonItem)
        .where(LooseComparisonItem.loose_comparison_id == loose_id)
        .order_by(LooseComparisonItem.id)
    )
    items = result.scalars().all()

    return templates.TemplateResponse(request, "loose/matches.html", {
        "loose_id": loose_id,
        "deal_id": loose_comp.deal_id,
        "deal_name": deal.name,
        "items": items,
        "matched_count": sum(1 for i in items if i.matched_sentence),
        "total_count": len(items),
    })


@router.post("/matches/{loose_id}")
async def confirm_matches(
    request: Request,
    loose_id: int,
    db: AsyncSession = Depends(get_db),
):
    """User confirms matches and proceeds to Excel upload."""
    loose_comp = await db.get(LooseComparison, loose_id)
    if not loose_comp:
        return RedirectResponse(url="/", status_code=303)

    form = await request.form()

    # Allow user to override matched sentences or remove items
    result = await db.execute(
        select(LooseComparisonItem)
        .where(LooseComparisonItem.loose_comparison_id == loose_id)
        .order_by(LooseComparisonItem.id)
    )
    items = result.scalars().all()

    for item in items:
        include = form.get(f"include_{item.id}")
        if not include:
            item.status = "skipped"
        else:
            override = form.get(f"override_{item.id}", "").strip()
            if override:
                item.matched_sentence = override
                item.status = None  # reset — will be processed

    await db.commit()
    return RedirectResponse(url=f"/loose/upload-excel/{loose_id}", status_code=303)


# ── Step 3: Upload Excel for loose comparison ─────────────────────────────────

@router.get("/upload-excel/{loose_id}")
async def upload_excel_form(
    request: Request,
    loose_id: int,
    db: AsyncSession = Depends(get_db),
):
    loose_comp = await db.get(LooseComparison, loose_id)
    if not loose_comp:
        return RedirectResponse(url="/", status_code=303)

    deal = await db.get(Deal, loose_comp.deal_id)

    # Count active items
    result = await db.execute(
        select(LooseComparisonItem)
        .where(LooseComparisonItem.loose_comparison_id == loose_id)
        .where(LooseComparisonItem.status != "skipped")
        .where(LooseComparisonItem.status != "no_match")
    )
    active_items = result.scalars().all()

    return templates.TemplateResponse(request, "loose/upload_excel.html", {
        "loose_id": loose_id,
        "deal_name": deal.name,
        "active_count": len(active_items),
        "excel_filename": loose_comp.excel_filename,
    })


@router.post("/upload-excel/{loose_id}")
async def upload_excel(
    request: Request,
    loose_id: int,
    excel_file: UploadFile = File(...),
    db: AsyncSession = Depends(get_db),
):
    loose_comp = await db.get(LooseComparison, loose_id)
    if not loose_comp:
        return RedirectResponse(url="/", status_code=303)

    deal = await db.get(Deal, loose_comp.deal_id)

    if not excel_file.filename.endswith(".xlsx"):
        return templates.TemplateResponse(request, "loose/upload_excel.html", {
            "loose_id": loose_id,
            "deal_name": deal.name,
            "active_count": 0,
            "error": "Please upload an .xlsx file.",
        })

    deal_dir = os.path.join(UPLOAD_DIR, str(loose_comp.deal_id))
    os.makedirs(deal_dir, exist_ok=True)
    filepath = os.path.join(deal_dir, excel_file.filename)

    with open(filepath, "wb") as f:
        content = await excel_file.read()
        f.write(content)

    # Validate
    try:
        sheet_names = get_sheet_names(filepath)
    except Exception:
        return templates.TemplateResponse(request, "loose/upload_excel.html", {
            "loose_id": loose_id,
            "deal_name": deal.name,
            "active_count": 0,
            "error": "Unable to read file. Please ensure it is a valid .xlsx file.",
        })

    loose_comp.excel_filename = excel_file.filename
    await db.commit()

    return RedirectResponse(url=f"/loose/sheet-map/{loose_id}", status_code=303)


# ── Step 3b: Map items to sheets ──────────────────────────────────────────────

@router.get("/sheet-map/{loose_id}")
async def sheet_map_form(
    request: Request,
    loose_id: int,
    db: AsyncSession = Depends(get_db),
):
    loose_comp = await db.get(LooseComparison, loose_id)
    if not loose_comp:
        return RedirectResponse(url="/", status_code=303)

    deal = await db.get(Deal, loose_comp.deal_id)

    filepath = os.path.join(UPLOAD_DIR, str(loose_comp.deal_id), loose_comp.excel_filename)
    sheet_names = get_sheet_names(filepath)

    result = await db.execute(
        select(LooseComparisonItem)
        .where(LooseComparisonItem.loose_comparison_id == loose_id)
        .where(LooseComparisonItem.status != "skipped")
        .where(LooseComparisonItem.status != "no_match")
        .order_by(LooseComparisonItem.id)
    )
    items = result.scalars().all()

    return templates.TemplateResponse(request, "loose/sheet_map.html", {
        "loose_id": loose_id,
        "deal_name": deal.name,
        "items": items,
        "sheet_names": sheet_names,
    })


@router.post("/sheet-map/{loose_id}")
async def submit_sheet_map(
    request: Request,
    loose_id: int,
    db: AsyncSession = Depends(get_db),
):
    loose_comp = await db.get(LooseComparison, loose_id)
    if not loose_comp:
        return RedirectResponse(url="/", status_code=303)

    form = await request.form()

    result = await db.execute(
        select(LooseComparisonItem)
        .where(LooseComparisonItem.loose_comparison_id == loose_id)
        .where(LooseComparisonItem.status != "skipped")
        .where(LooseComparisonItem.status != "no_match")
        .order_by(LooseComparisonItem.id)
    )
    items = result.scalars().all()

    filepath = os.path.join(UPLOAD_DIR, str(loose_comp.deal_id), loose_comp.excel_filename)

    # Run Claude extraction + comparison for each item
    for item in items:
        sheet_name = form.get(f"sheet_{item.id}", "")
        if not sheet_name:
            item.status = "error"
            item.error_message = "No sheet selected"
            continue

        try:
            excel_data = read_sheet_data(filepath, sheet_name)
        except Exception as e:
            item.status = "error"
            item.error_message = f"Error reading sheet: {e}"
            continue

        try:
            result_data = extract_and_compare(
                matched_sentence=item.matched_sentence,
                excel_data=excel_data,
                sheet_name=sheet_name,
            )

            item.extracted_values = result_data.get("excel_values", [])
            item.document_values = result_data.get("document_values", [])
            item.comparison_result = result_data
            item.status = result_data.get("status", "error")
            item.error_message = result_data.get("error")

        except Exception as e:
            item.status = "error"
            item.error_message = str(e)

    # Determine overall status
    statuses = [item.status for item in items]
    if all(s == "pass" for s in statuses if s not in ("skipped", "no_match")):
        loose_comp.status = "pass"
    elif any(s in ("fail", "error") for s in statuses):
        loose_comp.status = "fail"
    else:
        loose_comp.status = "pass"

    # Build output workbook
    output_items = []
    deal = await db.get(Deal, loose_comp.deal_id)
    for item in items:
        if item.comparison_result:
            output_items.append({
                "comparison_result": item.comparison_result,
                "summary": item.comparison_result.get("summary", ""),
                "word_filename": loose_comp.word_filename,
                "excel_filename": loose_comp.excel_filename,
                "sheet_name": form.get(f"sheet_{item.id}", ""),
            })

    output_data = build_loose_output_data(output_items)
    if output_data:
        output_dir = os.path.join(OUTPUT_DIR, f"loose_{loose_comp.id}")
        os.makedirs(output_dir, exist_ok=True)
        output_filename = f"{deal.name.replace(' ', '_')}_loose_comparison.xlsx"
        output_path = os.path.join(output_dir, output_filename)
        build_output_workbook(output_data, output_path)
        loose_comp.output_filename = output_filename

    await db.commit()

    return RedirectResponse(url=f"/loose/results/{loose_id}", status_code=303)


# ── Step 4: Results ───────────────────────────────────────────────────────────

@router.get("/results/{loose_id}")
async def loose_results(
    request: Request,
    loose_id: int,
    db: AsyncSession = Depends(get_db),
):
    loose_comp = await db.get(LooseComparison, loose_id)
    if not loose_comp:
        return RedirectResponse(url="/", status_code=303)

    deal = await db.get(Deal, loose_comp.deal_id)

    result = await db.execute(
        select(LooseComparisonItem)
        .where(LooseComparisonItem.loose_comparison_id == loose_id)
        .order_by(LooseComparisonItem.id)
    )
    items = result.scalars().all()

    total_comparisons = 0
    total_matches = 0
    total_mismatches = 0

    for item in items:
        if item.comparison_result and "comparisons" in item.comparison_result:
            for comp in item.comparison_result["comparisons"]:
                total_comparisons += 1
                if comp.get("status") == "match":
                    total_matches += 1
                else:
                    total_mismatches += 1

    return templates.TemplateResponse(request, "loose/results.html", {
        "loose_id": loose_id,
        "deal_id": loose_comp.deal_id,
        "deal_name": deal.name,
        "overall_status": loose_comp.status,
        "items": items,
        "total_comparisons": total_comparisons,
        "total_matches": total_matches,
        "total_mismatches": total_mismatches,
        "has_output": bool(loose_comp.output_filename),
    })


@router.get("/download/{loose_id}")
async def loose_download(
    loose_id: int,
    db: AsyncSession = Depends(get_db),
):
    loose_comp = await db.get(LooseComparison, loose_id)
    if not loose_comp or not loose_comp.output_filename:
        return RedirectResponse(url="/", status_code=303)

    output_path = os.path.join(OUTPUT_DIR, f"loose_{loose_id}", loose_comp.output_filename)
    if not os.path.exists(output_path):
        return RedirectResponse(url=f"/loose/results/{loose_id}", status_code=303)

    return FileResponse(
        output_path,
        filename=loose_comp.output_filename,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

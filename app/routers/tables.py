"""Route handlers for table selection, multi-file upload, and per-table mapping."""
import os
from fastapi import APIRouter, Request, UploadFile, File, Form, Depends
from fastapi.responses import RedirectResponse
from fastapi.templating import Jinja2Templates
from sqlalchemy.ext.asyncio import AsyncSession
from sqlalchemy import select

from app.config import TEMPLATES_DIR, UPLOAD_DIR
from app.database import get_db
from app.models import Comparison, ComparisonTable, UploadedFile
from app.services.word_parser import get_table_label
from app.services.excel_reader import get_file_sheets, read_file_data
from app.services.range_detector import detect_ranges

router = APIRouter()
templates = Jinja2Templates(directory=TEMPLATES_DIR)

ALLOWED_EXTENSIONS = {".xlsx", ".xml"}


# ── Step 2: Table selection (unchanged logic) ─────────────────────────────────

@router.get("/tables/{comparison_id}")
async def table_selection(
    request: Request,
    comparison_id: int,
    db: AsyncSession = Depends(get_db),
):
    comparison = await db.get(Comparison, comparison_id)
    if not comparison or not comparison.parsed_tables:
        return RedirectResponse(url="/", status_code=303)

    tables = comparison.parsed_tables
    table_labels = [get_table_label(t, i) for i, t in enumerate(tables)]

    return templates.TemplateResponse(request, "table_selection.html", {
        "comparison_id": comparison_id,
        "tables": tables,
        "table_labels": table_labels,
        "total_tables": len(tables),
        "selected_indices": comparison.selected_table_indices or [],
    })


@router.post("/select-tables/{comparison_id}")
async def select_tables(
    request: Request,
    comparison_id: int,
    db: AsyncSession = Depends(get_db),
):
    comparison = await db.get(Comparison, comparison_id)
    if not comparison:
        return RedirectResponse(url="/", status_code=303)

    form = await request.form()
    selected = form.getlist("selected_tables")
    selected_indices = [int(i) for i in selected]

    if not selected_indices:
        tables = comparison.parsed_tables
        table_labels = [get_table_label(t, i) for i, t in enumerate(tables)]
        return templates.TemplateResponse(request, "table_selection.html", {
            "comparison_id": comparison_id,
            "tables": tables,
            "table_labels": table_labels,
            "total_tables": len(tables),
            "selected_indices": [],
            "error": "Please select at least one table.",
        })

    comparison.selected_table_indices = sorted(selected_indices)
    await db.commit()

    return RedirectResponse(url=f"/upload-files/{comparison_id}", status_code=303)


# ── Step 3: Multi-file upload ─────────────────────────────────────────────────

@router.get("/upload-files/{comparison_id}")
async def file_upload_form(
    request: Request,
    comparison_id: int,
    db: AsyncSession = Depends(get_db),
):
    comparison = await db.get(Comparison, comparison_id)
    if not comparison:
        return RedirectResponse(url="/", status_code=303)

    selected_count = len(comparison.selected_table_indices or [])

    # Load already-uploaded files for this comparison
    result = await db.execute(
        select(UploadedFile)
        .where(UploadedFile.comparison_id == comparison_id)
        .order_by(UploadedFile.id)
    )
    uploaded_files = result.scalars().all()

    return templates.TemplateResponse(request, "file_upload.html", {
        "comparison_id": comparison_id,
        "selected_count": selected_count,
        "uploaded_files": uploaded_files,
    })


@router.post("/upload-files/{comparison_id}")
async def upload_files(
    request: Request,
    comparison_id: int,
    db: AsyncSession = Depends(get_db),
):
    comparison = await db.get(Comparison, comparison_id)
    if not comparison:
        return RedirectResponse(url="/", status_code=303)

    form = await request.form()
    files = form.getlist("source_files")

    if not files or all(getattr(f, "filename", "") == "" for f in files):
        return templates.TemplateResponse(request, "file_upload.html", {
            "comparison_id": comparison_id,
            "selected_count": len(comparison.selected_table_indices or []),
            "uploaded_files": [],
            "error": "Please upload at least one file.",
        })

    deal_dir = os.path.join(UPLOAD_DIR, str(comparison.deal_id))
    os.makedirs(deal_dir, exist_ok=True)

    errors = []
    for f in files:
        if not hasattr(f, "filename") or not f.filename:
            continue

        ext = os.path.splitext(f.filename)[1].lower()
        if ext not in ALLOWED_EXTENSIONS:
            errors.append(f"{f.filename}: unsupported file type. Use .xlsx or .xml")
            continue

        filepath = os.path.join(deal_dir, f.filename)
        content = await f.read()
        with open(filepath, "wb") as out:
            out.write(content)

        # Validate we can read the file
        try:
            sheets = get_file_sheets(filepath)
        except Exception as e:
            os.remove(filepath)
            errors.append(f"{f.filename}: unable to read file ({e})")
            continue

        file_type = ext.lstrip(".")
        uploaded = UploadedFile(
            comparison_id=comparison_id,
            filename=f.filename,
            file_type=file_type,
        )
        db.add(uploaded)

    if errors:
        await db.rollback()
        return templates.TemplateResponse(request, "file_upload.html", {
            "comparison_id": comparison_id,
            "selected_count": len(comparison.selected_table_indices or []),
            "uploaded_files": [],
            "error": " | ".join(errors),
        })

    await db.commit()
    return RedirectResponse(url=f"/map-tables/{comparison_id}", status_code=303)


# ── Step 4: Per-table mapping ─────────────────────────────────────────────────

@router.get("/map-tables/{comparison_id}")
async def map_tables_form(
    request: Request,
    comparison_id: int,
    db: AsyncSession = Depends(get_db),
):
    comparison = await db.get(Comparison, comparison_id)
    if not comparison:
        return RedirectResponse(url="/", status_code=303)

    from app.models import Deal
    deal = await db.get(Deal, comparison.deal_id)

    tables = comparison.parsed_tables
    selected_indices = comparison.selected_table_indices or []

    # Load uploaded files
    result = await db.execute(
        select(UploadedFile)
        .where(UploadedFile.comparison_id == comparison_id)
        .order_by(UploadedFile.id)
    )
    uploaded_files = result.scalars().all()

    # Build sheet lists for each file
    file_sheets = {}
    for uf in uploaded_files:
        filepath = os.path.join(UPLOAD_DIR, str(comparison.deal_id), uf.filename)
        try:
            file_sheets[uf.id] = get_file_sheets(filepath)
        except Exception:
            file_sheets[uf.id] = []

    # Load existing mappings if any
    result = await db.execute(
        select(ComparisonTable)
        .where(ComparisonTable.comparison_id == comparison_id)
        .order_by(ComparisonTable.table_index)
    )
    existing_mappings = {ct.table_index: ct for ct in result.scalars().all()}

    # Build data for each selected table
    mapping_data = []
    for table_idx in selected_indices:
        table = tables[table_idx] if table_idx < len(tables) else []
        label = get_table_label(table, table_idx)
        existing = existing_mappings.get(table_idx)

        mapping_data.append({
            "table_index": table_idx,
            "table_label": label,
            "word_table": table,
            "word_rows": len(table),
            "word_cols": max((len(r) for r in table), default=0),
            "current_file_id": existing.uploaded_file_id if existing else None,
            "current_sheet": existing.excel_tab_name if existing else None,
        })

    return templates.TemplateResponse(request, "table_mapping.html", {
        "comparison_id": comparison_id,
        "deal_name": deal.name if deal else "",
        "mapping_data": mapping_data,
        "uploaded_files": uploaded_files,
        "file_sheets": file_sheets,
    })


@router.post("/map-tables/{comparison_id}")
async def submit_mappings(
    request: Request,
    comparison_id: int,
    db: AsyncSession = Depends(get_db),
):
    comparison = await db.get(Comparison, comparison_id)
    if not comparison:
        return RedirectResponse(url="/", status_code=303)

    form = await request.form()
    tables = comparison.parsed_tables
    selected_indices = comparison.selected_table_indices or []

    # Delete old ComparisonTable entries for this comparison
    result = await db.execute(
        select(ComparisonTable)
        .where(ComparisonTable.comparison_id == comparison_id)
    )
    for old_ct in result.scalars().all():
        await db.delete(old_ct)

    # Load uploaded files for path resolution
    result = await db.execute(
        select(UploadedFile)
        .where(UploadedFile.comparison_id == comparison_id)
    )
    file_map = {uf.id: uf for uf in result.scalars().all()}

    errors = []
    for table_idx in selected_indices:
        file_id_str = form.get(f"file_{table_idx}")
        sheet_name = form.get(f"sheet_{table_idx}", "")

        if not file_id_str:
            errors.append(f"Table {table_idx + 1}: no file selected")
            continue

        file_id = int(file_id_str)
        uf = file_map.get(file_id)
        if not uf:
            errors.append(f"Table {table_idx + 1}: invalid file")
            continue

        # Run multi-range detection
        filepath = os.path.join(UPLOAD_DIR, str(comparison.deal_id), uf.filename)
        detected_range_list = []

        if uf.file_type == "xlsx" and sheet_name:
            try:
                regions = detect_ranges(filepath, sheet_name)
                detected_range_list = [r["range"] for r in regions if r.get("range")]
            except Exception:
                detected_range_list = []

        ct = ComparisonTable(
            comparison_id=comparison_id,
            table_index=table_idx,
            table_label=get_table_label(tables[table_idx], table_idx),
            uploaded_file_id=file_id,
            excel_tab_name=sheet_name,
            detected_ranges=detected_range_list,
            selected_range=detected_range_list[0] if detected_range_list else None,
            detected_range=detected_range_list[0] if detected_range_list else None,
        )
        db.add(ct)

    if errors:
        await db.rollback()
        # Re-render with error — simplified: redirect back
        return RedirectResponse(url=f"/map-tables/{comparison_id}", status_code=303)

    # Set excel_filename for backward compat (use first uploaded file)
    first_file = list(file_map.values())[0] if file_map else None
    if first_file:
        comparison.excel_filename = first_file.filename

    await db.commit()
    return RedirectResponse(url=f"/range-review/{comparison_id}", status_code=303)


# ── HTMX: Get sheets for a file ──────────────────────────────────────────────

@router.get("/api/file-sheets/{comparison_id}/{file_id}")
async def get_sheets_for_file(
    comparison_id: int,
    file_id: int,
    db: AsyncSession = Depends(get_db),
):
    """Return sheet names as JSON for a given uploaded file."""
    from fastapi.responses import JSONResponse

    comparison = await db.get(Comparison, comparison_id)
    uf = await db.get(UploadedFile, file_id)
    if not comparison or not uf:
        return JSONResponse({"sheets": []})

    filepath = os.path.join(UPLOAD_DIR, str(comparison.deal_id), uf.filename)
    try:
        sheets = get_file_sheets(filepath)
    except Exception:
        sheets = []

    return JSONResponse({"sheets": sheets, "file_type": uf.file_type})

"""Route handlers for table selection and Excel upload."""
import os
from fastapi import APIRouter, Request, UploadFile, File, Form, Depends
from fastapi.responses import RedirectResponse
from fastapi.templating import Jinja2Templates
from sqlalchemy.ext.asyncio import AsyncSession
from sqlalchemy import select

from app.config import TEMPLATES_DIR, UPLOAD_DIR
from app.database import get_db
from app.models import Comparison, ComparisonTable
from app.services.word_parser import get_table_label
from app.services.excel_reader import get_sheet_names, read_sheet_data
from app.services.range_detector import detect_range

router = APIRouter()
templates = Jinja2Templates(directory=TEMPLATES_DIR)


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

    return RedirectResponse(url=f"/upload-excel/{comparison_id}", status_code=303)


@router.get("/upload-excel/{comparison_id}")
async def excel_upload_form(
    request: Request,
    comparison_id: int,
    db: AsyncSession = Depends(get_db),
):
    comparison = await db.get(Comparison, comparison_id)
    if not comparison:
        return RedirectResponse(url="/", status_code=303)

    selected_count = len(comparison.selected_table_indices or [])

    return templates.TemplateResponse(request, "excel_upload.html", {
        "comparison_id": comparison_id,
        "selected_count": selected_count,
    })


@router.post("/upload-excel/{comparison_id}")
async def upload_excel(
    request: Request,
    comparison_id: int,
    excel_file: UploadFile = File(...),
    db: AsyncSession = Depends(get_db),
):
    comparison = await db.get(Comparison, comparison_id)
    if not comparison:
        return RedirectResponse(url="/", status_code=303)

    # Validate file type
    if not excel_file.filename.endswith(".xlsx"):
        return templates.TemplateResponse(request, "excel_upload.html", {
            "comparison_id": comparison_id,
            "selected_count": len(comparison.selected_table_indices or []),
            "error": "Please upload an .xlsx file",
        })

    # Save file
    deal_dir = os.path.join(UPLOAD_DIR, str(comparison.deal_id))
    os.makedirs(deal_dir, exist_ok=True)
    filepath = os.path.join(deal_dir, excel_file.filename)

    with open(filepath, "wb") as f:
        content = await excel_file.read()
        f.write(content)

    comparison.excel_filename = excel_file.filename

    # Read sheet names
    try:
        sheet_names = get_sheet_names(filepath)
    except Exception:
        return templates.TemplateResponse(request, "excel_upload.html", {
            "comparison_id": comparison_id,
            "selected_count": len(comparison.selected_table_indices or []),
            "error": "Unable to read file. Please ensure it is a valid, unprotected .xlsx file.",
        })

    selected_indices = comparison.selected_table_indices or []
    if len(sheet_names) < len(selected_indices):
        return templates.TemplateResponse(request, "excel_upload.html", {
            "comparison_id": comparison_id,
            "selected_count": len(selected_indices),
            "error": f"Excel has {len(sheet_names)} tabs but {len(selected_indices)} tables were selected. Please upload a file with at least {len(selected_indices)} tabs.",
        })

    # Run range detection for each mapped tab
    detected_ranges = {}
    tables = comparison.parsed_tables

    for idx, table_idx in enumerate(selected_indices):
        sheet_name = sheet_names[idx]
        try:
            result = detect_range(filepath, sheet_name)
            detected_ranges[str(idx)] = result.get("range", "")
        except Exception:
            detected_ranges[str(idx)] = ""

    comparison.detected_ranges = detected_ranges

    # Create ComparisonTable entries
    for idx, table_idx in enumerate(selected_indices):
        ct = ComparisonTable(
            comparison_id=comparison.id,
            table_index=table_idx,
            table_label=get_table_label(tables[table_idx], table_idx),
            excel_tab_name=sheet_names[idx],
            detected_range=detected_ranges.get(str(idx), ""),
        )
        db.add(ct)

    await db.commit()

    return RedirectResponse(url=f"/range-review/{comparison_id}", status_code=303)

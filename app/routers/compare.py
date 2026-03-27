"""Route handlers for range review, precision review, comparison, and results."""
import os
from fastapi import APIRouter, Request, Form, Depends
from fastapi.responses import RedirectResponse, FileResponse
from fastapi.templating import Jinja2Templates
from sqlalchemy.ext.asyncio import AsyncSession
from sqlalchemy import select

from app.database import get_db
from app.models import Comparison, ComparisonTable, Deal
from app.services.word_parser import get_table_label
from app.services.excel_reader import get_sheet_names, read_sheet_data
from app.services.range_detector import detect_range
from app.services.comparator import (
    detect_table_precision,
    compare_tables,
    is_numeric_string,
)
from app.services.output_builder import build_output_workbook

router = APIRouter()
templates = Jinja2Templates(directory="app/templates")

UPLOAD_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), "uploads")
OUTPUT_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), "outputs")


@router.get("/range-review/{comparison_id}")
async def range_review(
    request: Request,
    comparison_id: int,
    db: AsyncSession = Depends(get_db),
):
    comparison = await db.get(Comparison, comparison_id)
    if not comparison:
        return RedirectResponse(url="/", status_code=303)

    deal = await db.get(Deal, comparison.deal_id)

    # Load comparison tables
    result = await db.execute(
        select(ComparisonTable)
        .where(ComparisonTable.comparison_id == comparison_id)
        .order_by(ComparisonTable.table_index)
    )
    comp_tables = result.scalars().all()

    # Load Excel data for each tab
    excel_path = os.path.join(UPLOAD_DIR, str(comparison.deal_id), comparison.excel_filename)
    tables = comparison.parsed_tables
    selected_indices = comparison.selected_table_indices or []

    mappings = []
    for idx, ct in enumerate(comp_tables):
        table_idx = ct.table_index
        word_table = tables[table_idx] if table_idx < len(tables) else []

        # Get detected range data
        current_range = ct.user_range_override or ct.detected_range
        excel_data = []
        error = None

        if current_range:
            try:
                excel_data = read_sheet_data(excel_path, ct.excel_tab_name, current_range)
                excel_data = _serialize_data(excel_data)
            except Exception as e:
                error = str(e)
        else:
            error = f"No data detected on sheet {ct.excel_tab_name}. Please specify a range manually."

        mappings.append({
            "idx": idx,
            "table_index": table_idx,
            "table_label": ct.table_label,
            "excel_tab_name": ct.excel_tab_name,
            "detected_range": ct.detected_range or "",
            "current_range": current_range or "",
            "word_table": word_table,
            "excel_data": excel_data,
            "word_rows": len(word_table),
            "word_cols": max((len(r) for r in word_table), default=0),
            "excel_rows": len(excel_data),
            "excel_cols": max((len(r) for r in excel_data), default=0),
            "error": error,
            "ct_id": ct.id,
        })

    return templates.TemplateResponse("range_review.html", {
        "request": request,
        "comparison_id": comparison_id,
        "deal_name": deal.name,
        "mappings": mappings,
    })


@router.post("/update-range/{comparison_id}/{ct_id}")
async def update_range(
    request: Request,
    comparison_id: int,
    ct_id: int,
    range_override: str = Form(""),
    db: AsyncSession = Depends(get_db),
):
    """HTMX endpoint to update range and return refreshed preview."""
    comparison = await db.get(Comparison, comparison_id)
    ct = await db.get(ComparisonTable, ct_id)
    if not comparison or not ct:
        return templates.TemplateResponse("partials/range_preview.html", {
            "request": request,
            "mapping": {"error": "Comparison not found"},
        })

    range_override = range_override.strip().upper()
    ct.user_range_override = range_override if range_override else None
    await db.commit()

    # Re-read data with new range
    excel_path = os.path.join(UPLOAD_DIR, str(comparison.deal_id), comparison.excel_filename)
    current_range = range_override or ct.detected_range
    excel_data = []
    error = None

    if current_range:
        try:
            excel_data = read_sheet_data(excel_path, ct.excel_tab_name, current_range)
            excel_data = _serialize_data(excel_data)
        except ValueError as e:
            error = str(e)
        except Exception as e:
            error = f"Error reading range: {e}"

    tables = comparison.parsed_tables
    word_table = tables[ct.table_index] if ct.table_index < len(tables) else []

    mapping = {
        "idx": 0,
        "table_index": ct.table_index,
        "table_label": ct.table_label,
        "excel_tab_name": ct.excel_tab_name,
        "detected_range": ct.detected_range or "",
        "current_range": current_range or "",
        "word_table": word_table,
        "excel_data": excel_data,
        "word_rows": len(word_table),
        "word_cols": max((len(r) for r in word_table), default=0),
        "excel_rows": len(excel_data),
        "excel_cols": max((len(r) for r in excel_data), default=0),
        "error": error,
        "ct_id": ct.id,
    }

    return templates.TemplateResponse("partials/range_preview.html", {
        "request": request,
        "comparison_id": comparison_id,
        "mapping": mapping,
    })


@router.post("/confirm-ranges/{comparison_id}")
async def confirm_ranges(
    request: Request,
    comparison_id: int,
    db: AsyncSession = Depends(get_db),
):
    return RedirectResponse(url=f"/precision/{comparison_id}", status_code=303)


@router.get("/precision/{comparison_id}")
async def precision_review(
    request: Request,
    comparison_id: int,
    db: AsyncSession = Depends(get_db),
):
    comparison = await db.get(Comparison, comparison_id)
    if not comparison:
        return RedirectResponse(url="/", status_code=303)

    deal = await db.get(Deal, comparison.deal_id)
    tables = comparison.parsed_tables

    result = await db.execute(
        select(ComparisonTable)
        .where(ComparisonTable.comparison_id == comparison_id)
        .order_by(ComparisonTable.table_index)
    )
    comp_tables = result.scalars().all()

    precision_data = []
    for ct in comp_tables:
        word_table = tables[ct.table_index] if ct.table_index < len(tables) else []
        auto_precisions = detect_table_precision(word_table)
        overrides = ct.precision_overrides or {}

        rows = []
        for r_idx, row in enumerate(word_table):
            auto_prec = auto_precisions[r_idx] if r_idx < len(auto_precisions) else 2
            override = overrides.get(str(r_idx))
            rows.append({
                "row_index": r_idx,
                "cells": row,
                "auto_precision": auto_prec,
                "override": override,
                "effective": override if override is not None else auto_prec,
            })

        precision_data.append({
            "ct_id": ct.id,
            "table_label": ct.table_label,
            "rows": rows,
        })

    return templates.TemplateResponse("precision.html", {
        "request": request,
        "comparison_id": comparison_id,
        "deal_name": deal.name,
        "precision_data": precision_data,
    })


@router.post("/precision/{comparison_id}")
async def submit_precision(
    request: Request,
    comparison_id: int,
    db: AsyncSession = Depends(get_db),
):
    comparison = await db.get(Comparison, comparison_id)
    if not comparison:
        return RedirectResponse(url="/", status_code=303)

    form = await request.form()

    result = await db.execute(
        select(ComparisonTable)
        .where(ComparisonTable.comparison_id == comparison_id)
        .order_by(ComparisonTable.table_index)
    )
    comp_tables = result.scalars().all()

    # Parse precision overrides from form
    for ct in comp_tables:
        overrides = {}
        word_table = comparison.parsed_tables[ct.table_index]
        auto_precisions = detect_table_precision(word_table)

        for r_idx in range(len(word_table)):
            field_name = f"precision_{ct.id}_{r_idx}"
            val = form.get(field_name)
            if val is not None and val != "":
                val_int = int(val)
                auto = auto_precisions[r_idx] if r_idx < len(auto_precisions) else 2
                if val_int != auto:
                    overrides[str(r_idx)] = val_int

        ct.precision_overrides = overrides if overrides else None

    await db.commit()

    # Now run the comparison
    deal = await db.get(Deal, comparison.deal_id)
    excel_path = os.path.join(UPLOAD_DIR, str(comparison.deal_id), comparison.excel_filename)
    tables = comparison.parsed_tables

    output_comparisons = []
    overall_pass = True

    for ct in comp_tables:
        word_table = tables[ct.table_index] if ct.table_index < len(tables) else []
        current_range = ct.user_range_override or ct.detected_range

        try:
            excel_data = read_sheet_data(excel_path, ct.excel_tab_name, current_range)
        except Exception:
            excel_data = []

        # Build effective precisions
        auto_precisions = detect_table_precision(word_table)
        overrides = ct.precision_overrides or {}
        row_precisions = []
        for r_idx in range(len(word_table)):
            auto = auto_precisions[r_idx] if r_idx < len(auto_precisions) else 2
            override = overrides.get(str(r_idx))
            row_precisions.append(override if override is not None else auto)

        # Run comparison
        comp_result = compare_tables(word_table, excel_data, row_precisions)

        ct.match_count = comp_result["match_count"]
        ct.mismatch_count = comp_result["mismatch_count"]
        ct.total_cells = comp_result["total_cells"]
        ct.comparison_data = {
            "diff_grid": _serialize_data(comp_result["diff_grid"]),
            "status_grid": comp_result["status_grid"],
        }

        if comp_result["mismatch_count"] > 0:
            overall_pass = False

        output_comparisons.append({
            "table_label": ct.table_label,
            "word_filename": comparison.word_filename,
            "excel_filename": comparison.excel_filename,
            "excel_tab_name": ct.excel_tab_name,
            "excel_range": current_range or "",
            "word_table": word_table,
            "excel_data": _serialize_output_data(excel_data),
            "diff_grid": comp_result["diff_grid"],
            "status_grid": comp_result["status_grid"],
            "row_precisions": row_precisions,
        })

    # Build output workbook
    output_dir = os.path.join(OUTPUT_DIR, str(comparison_id))
    os.makedirs(output_dir, exist_ok=True)
    output_filename = f"{deal.name.replace(' ', '_')}_comparison.xlsx"
    output_path = os.path.join(output_dir, output_filename)

    build_output_workbook(output_comparisons, output_path)

    comparison.status = "pass" if overall_pass else "fail"
    comparison.output_filename = output_filename
    await db.commit()

    return RedirectResponse(url=f"/results/{comparison_id}", status_code=303)


@router.get("/results/{comparison_id}")
async def results(
    request: Request,
    comparison_id: int,
    db: AsyncSession = Depends(get_db),
):
    comparison = await db.get(Comparison, comparison_id)
    if not comparison:
        return RedirectResponse(url="/", status_code=303)

    deal = await db.get(Deal, comparison.deal_id)

    result = await db.execute(
        select(ComparisonTable)
        .where(ComparisonTable.comparison_id == comparison_id)
        .order_by(ComparisonTable.table_index)
    )
    comp_tables = result.scalars().all()

    total_match = sum(ct.match_count or 0 for ct in comp_tables)
    total_mismatch = sum(ct.mismatch_count or 0 for ct in comp_tables)
    total_cells = sum(ct.total_cells or 0 for ct in comp_tables)

    table_results = []
    tables = comparison.parsed_tables
    for ct in comp_tables:
        word_table = tables[ct.table_index] if ct.table_index < len(tables) else []
        table_results.append({
            "table_label": ct.table_label,
            "excel_tab_name": ct.excel_tab_name,
            "match_count": ct.match_count or 0,
            "mismatch_count": ct.mismatch_count or 0,
            "total_cells": ct.total_cells or 0,
            "status": "pass" if (ct.mismatch_count or 0) == 0 else "fail",
            "diff_grid": ct.comparison_data.get("diff_grid", []) if ct.comparison_data else [],
            "status_grid": ct.comparison_data.get("status_grid", []) if ct.comparison_data else [],
            "word_table": word_table,
        })

    return templates.TemplateResponse("results.html", {
        "request": request,
        "comparison_id": comparison_id,
        "deal_name": deal.name,
        "overall_status": comparison.status,
        "total_match": total_match,
        "total_mismatch": total_mismatch,
        "total_cells": total_cells,
        "table_results": table_results,
    })


@router.get("/download/{comparison_id}")
async def download(
    comparison_id: int,
    db: AsyncSession = Depends(get_db),
):
    comparison = await db.get(Comparison, comparison_id)
    if not comparison or not comparison.output_filename:
        return RedirectResponse(url="/", status_code=303)

    output_path = os.path.join(OUTPUT_DIR, str(comparison_id), comparison.output_filename)
    if not os.path.exists(output_path):
        return RedirectResponse(url=f"/results/{comparison_id}", status_code=303)

    return FileResponse(
        output_path,
        filename=comparison.output_filename,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


def _serialize_data(data: list[list]) -> list[list]:
    """Convert data to JSON-serializable format."""
    result = []
    for row in data:
        serialized_row = []
        for val in row:
            if val is None:
                serialized_row.append(None)
            elif isinstance(val, float):
                serialized_row.append(val)
            elif isinstance(val, int):
                serialized_row.append(val)
            else:
                serialized_row.append(str(val))
        result.append(serialized_row)
    return result


def _serialize_output_data(data: list[list]) -> list[list]:
    """Preserve types for output builder."""
    result = []
    for row in data:
        out_row = []
        for val in row:
            out_row.append(val)
        result.append(out_row)
    return result

"""Route handlers for file uploads and new comparison creation."""
import os
import shutil
from fastapi import APIRouter, Request, UploadFile, File, Form, Depends
from fastapi.responses import RedirectResponse
from fastapi.templating import Jinja2Templates
from sqlalchemy.ext.asyncio import AsyncSession

from app.config import TEMPLATES_DIR, UPLOAD_DIR
from app.database import get_db
from app.models import Deal, Comparison
from app.services.word_parser import extract_tables, get_table_label

router = APIRouter()
templates = Jinja2Templates(directory=TEMPLATES_DIR)


@router.get("/")
async def index(request: Request):
    return templates.TemplateResponse(request, "index.html")


@router.post("/upload-word")
async def upload_word(
    request: Request,
    deal_name: str = Form(...),
    word_file: UploadFile = File(...),
    db: AsyncSession = Depends(get_db),
):
    # Validate file type
    if not word_file.filename.endswith(".docx"):
        return templates.TemplateResponse(request, "index.html", {
            "error": "Please upload a .docx file",
            "deal_name": deal_name,
        })

    # Create deal
    deal = Deal(name=deal_name)
    db.add(deal)
    await db.flush()

    # Create comparison
    comparison = Comparison(
        deal_id=deal.id,
        word_filename=word_file.filename,
    )
    db.add(comparison)
    await db.flush()

    # Save file
    deal_dir = os.path.join(UPLOAD_DIR, str(deal.id))
    os.makedirs(deal_dir, exist_ok=True)
    filepath = os.path.join(deal_dir, word_file.filename)

    with open(filepath, "wb") as f:
        content = await word_file.read()
        f.write(content)

    # Parse tables
    try:
        parsed_tables = extract_tables(filepath)
    except Exception as e:
        return templates.TemplateResponse(request, "index.html", {
            "error": f"Unable to read file. Please ensure it is a valid, unprotected .docx file. Error: {e}",
            "deal_name": deal_name,
        })

    # Store parsed tables (may be empty — that's OK for loose language track)
    comparison.parsed_tables = parsed_tables or []
    await db.commit()

    # Go to the track selection page
    return RedirectResponse(url=f"/choose-track/{deal.id}/{comparison.id}", status_code=303)


@router.get("/choose-track/{deal_id}/{comparison_id}")
async def choose_track(
    request: Request,
    deal_id: int,
    comparison_id: int,
    db: AsyncSession = Depends(get_db),
):
    deal = await db.get(Deal, deal_id)
    comparison = await db.get(Comparison, comparison_id)
    if not deal or not comparison:
        return RedirectResponse(url="/", status_code=303)

    table_count = len(comparison.parsed_tables) if comparison.parsed_tables else 0

    return templates.TemplateResponse(request, "choose_track.html", {
        "deal_id": deal_id,
        "comparison_id": comparison_id,
        "deal_name": deal.name,
        "word_filename": comparison.word_filename,
        "table_count": table_count,
    })

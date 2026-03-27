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
    return templates.TemplateResponse("index.html", {"request": request})


@router.post("/upload-word")
async def upload_word(
    request: Request,
    deal_name: str = Form(...),
    word_file: UploadFile = File(...),
    db: AsyncSession = Depends(get_db),
):
    # Validate file type
    if not word_file.filename.endswith(".docx"):
        return templates.TemplateResponse("index.html", {
            "request": request,
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
        tables = extract_tables(filepath)
    except Exception as e:
        return templates.TemplateResponse("index.html", {
            "request": request,
            "error": f"Unable to read file. Please ensure it is a valid, unprotected .docx file. Error: {e}",
            "deal_name": deal_name,
        })

    if not tables:
        return templates.TemplateResponse("index.html", {
            "request": request,
            "error": "No tables found in this document",
            "deal_name": deal_name,
        })

    # Store parsed tables as JSON-serializable data
    comparison.parsed_tables = tables
    await db.commit()

    return RedirectResponse(url=f"/tables/{comparison.id}", status_code=303)

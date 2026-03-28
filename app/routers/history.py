"""Route handlers for comparison history."""
from fastapi import APIRouter, Request, Depends
from fastapi.responses import RedirectResponse
from fastapi.templating import Jinja2Templates
from sqlalchemy.ext.asyncio import AsyncSession
from sqlalchemy import select
from sqlalchemy.orm import selectinload

from app.config import TEMPLATES_DIR
from app.database import get_db
from app.models import Deal, Comparison, ComparisonTable, LooseComparison

router = APIRouter()
templates = Jinja2Templates(directory=TEMPLATES_DIR)


@router.get("/history")
async def history_list(
    request: Request,
    db: AsyncSession = Depends(get_db),
):
    result = await db.execute(
        select(Deal).order_by(Deal.created_at.desc())
    )
    deals = result.scalars().all()

    deals_data = []
    for deal in deals:
        comp_result = await db.execute(
            select(Comparison)
            .where(Comparison.deal_id == deal.id)
            .order_by(Comparison.created_at.desc())
        )
        comparisons = comp_result.scalars().all()

        loose_result = await db.execute(
            select(LooseComparison)
            .where(LooseComparison.deal_id == deal.id)
            .order_by(LooseComparison.created_at.desc())
        )
        loose_comparisons = loose_result.scalars().all()

        deals_data.append({
            "deal": deal,
            "comparisons": comparisons,
            "loose_comparisons": loose_comparisons,
        })

    return templates.TemplateResponse(request, "history.html", {
        "deals_data": deals_data,
    })


@router.get("/history/{comparison_id}")
async def history_detail(
    request: Request,
    comparison_id: int,
    db: AsyncSession = Depends(get_db),
):
    return RedirectResponse(url=f"/results/{comparison_id}", status_code=303)

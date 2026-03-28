"""FastAPI application entry point."""
import logging
from contextlib import asynccontextmanager
from fastapi import FastAPI, Request
from fastapi.responses import PlainTextResponse
from fastapi.staticfiles import StaticFiles

from app.config import STATIC_DIR
from app.database import init_db
from app.routers import upload, tables, compare, history

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


@asynccontextmanager
async def lifespan(app: FastAPI):
    await init_db()
    logger.info("Database initialized, app ready")
    yield


app = FastAPI(
    title="Offering Document Reconciliation Tool",
    version="1.1",
    lifespan=lifespan,
)


@app.exception_handler(Exception)
async def global_exception_handler(request: Request, exc: Exception):
    logger.exception(f"Unhandled error on {request.url.path}")
    return PlainTextResponse(f"Internal Server Error: {exc}", status_code=500)


# Mount static files
app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")

# Register routers
app.include_router(upload.router)
app.include_router(tables.router)
app.include_router(compare.router)
app.include_router(history.router)

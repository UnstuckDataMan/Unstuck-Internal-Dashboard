import logging
import os
import sys
import importlib.util
from contextlib import asynccontextmanager
from pathlib import Path

from fastapi import FastAPI, Request
from fastapi.staticfiles import StaticFiles
from starlette.middleware.wsgi import WSGIMiddleware

from app.deps import templates
from app.routers import gender, city, dnc, reply_bank, mail_merge, copy_bank, copy_bank_export, campaigns

logger = logging.getLogger(__name__)

# ── Scheduled auto-sync ───────────────────────────────────────────────────────
# Runs twice a day for every campaign that has a linked Google Sheet.
# Override the hours (UTC) via AUTO_SYNC_HOURS env var, e.g. "8,20"
_AUTO_SYNC_HOURS = os.environ.get("AUTO_SYNC_HOURS", "9,21")

@asynccontextmanager
async def _lifespan(app: FastAPI):
    try:
        from apscheduler.schedulers.asyncio import AsyncIOScheduler
        from apscheduler.triggers.cron import CronTrigger
        from app.utils.auto_sync import run_auto_sync

        hours = _AUTO_SYNC_HOURS.strip()
        scheduler = AsyncIOScheduler()
        scheduler.add_job(
            run_auto_sync,
            CronTrigger(hour=hours, minute=0),
            id="auto_sync",
            name="Twice-daily campaign sync",
            replace_existing=True,
        )
        scheduler.start()
        logger.info("Auto-sync scheduler started (UTC hours: %s).", hours)
    except Exception as exc:
        logger.warning("Auto-sync scheduler could not start: %s", exc)
        scheduler = None

    yield  # app runs here

    if scheduler and scheduler.running:
        scheduler.shutdown(wait=False)
        logger.info("Auto-sync scheduler stopped.")


app = FastAPI(title="Data Enrichment Dashboard", lifespan=_lifespan)

_BASE = Path(__file__).resolve().parent
app.mount("/static", StaticFiles(directory=str(_BASE / "static")), name="static")

app.include_router(gender.router)
app.include_router(city.router)
app.include_router(dnc.router)
app.include_router(reply_bank.router)
app.include_router(mail_merge.router)
app.include_router(copy_bank.router)
app.include_router(copy_bank_export.router)
app.include_router(campaigns.router)

# ── Mail Merge Tool (Flask WSGI sub-application) ──────────────────────────────
# Load mail_merge/app.py explicitly by file path to avoid collision with the
# 'app' package name (this directory).  sys.path is extended first so that
# the Flask app's own `from utils.xxx import ...` statements resolve correctly.
_MM_DIR = _BASE.parent / "mail_merge"
if str(_MM_DIR) not in sys.path:
    sys.path.insert(0, str(_MM_DIR))

_spec = importlib.util.spec_from_file_location("mail_merge_flask", str(_MM_DIR / "app.py"))
_mm_module = importlib.util.module_from_spec(_spec)
sys.modules["mail_merge_flask"] = _mm_module   # register before exec so relative imports work
_spec.loader.exec_module(_mm_module)

app.mount("/mail-merge", WSGIMiddleware(_mm_module.app))
# ─────────────────────────────────────────────────────────────────────────────


@app.get("/")
async def index(request: Request):
    return templates.TemplateResponse("index.html", {"request": request, "active": "home"})

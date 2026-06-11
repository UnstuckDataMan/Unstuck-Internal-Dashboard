import logging
import os
from contextlib import asynccontextmanager
from pathlib import Path

from fastapi import FastAPI, Request
from fastapi.staticfiles import StaticFiles

from app.deps import templates
from app.routers import gender, city, dnc, reply_bank, mail_merge, copy_bank, copy_bank_export, campaigns, launch_checker, targeting_checker, bd_targeting

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
app.include_router(launch_checker.router)
app.include_router(targeting_checker.router)
app.include_router(bd_targeting.router)

@app.get("/")
async def index(request: Request):
    return templates.TemplateResponse("index.html", {"request": request, "active": "home"})

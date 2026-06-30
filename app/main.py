import logging
import os
import secrets
from contextlib import asynccontextmanager
from pathlib import Path

from fastapi import FastAPI, Request
from fastapi.staticfiles import StaticFiles
from starlette.middleware.sessions import SessionMiddleware

from app import auth
from app.deps import templates
from app.routers import (
    gender, city, dnc, reply_bank, mail_merge, copy_bank, copy_bank_export,
    campaigns, launch_checker, targeting_checker, bd_targeting, profiles,
    auth as auth_router, admin,
)

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

# ── Auth gate + session ────────────────────────────────────────────────────────
# Default-deny: every request without a valid session is bounced to /login,
# except the public allowlist; then RBAC checks the request's tool against the
# user's tools. Disabled when AUTH_DISABLED is set (offline tests / rollback).
@app.middleware("http")
async def auth_gate(request: Request, call_next):
    blocked = auth.access_response(request)
    if blocked is not None:
        return blocked
    return await call_next(request)

# SessionMiddleware is added LAST so it wraps auth_gate (Starlette runs the most
# recently added middleware first) — the session must be decoded before the gate
# reads it. A missing SESSION_SECRET falls back to a random per-boot key (safe,
# but logs everyone out on restart) rather than a guessable default.
_SESSION_SECRET = os.environ.get("SESSION_SECRET")
if not _SESSION_SECRET:
    _SESSION_SECRET = secrets.token_urlsafe(32)
    logger.warning("SESSION_SECRET not set — using a random per-boot key; "
                   "sessions will not survive restarts. Set SESSION_SECRET in production.")
app.add_middleware(
    SessionMiddleware,
    secret_key=_SESSION_SECRET,
    session_cookie="dashboard_session",
    same_site="lax",
    https_only=os.environ.get("APP_BASE_URL", "").startswith("https"),
    max_age=60 * 60 * 8,   # 8 hours
)

_BASE = Path(__file__).resolve().parent
app.mount("/static", StaticFiles(directory=str(_BASE / "static")), name="static")

app.include_router(auth_router.router)
app.include_router(admin.router)
app.include_router(gender.router)
app.include_router(city.router)
app.include_router(dnc.router)
app.include_router(reply_bank.router)
app.include_router(mail_merge.router)
app.include_router(copy_bank.router)
app.include_router(copy_bank_export.router)
app.include_router(campaigns.router)
app.include_router(profiles.router)
app.include_router(launch_checker.router)
app.include_router(targeting_checker.router)
app.include_router(bd_targeting.router)


@app.get("/healthz")
async def healthz():
    """Public, unauthenticated health check for Render."""
    return {"ok": True}


@app.get("/")
async def index(request: Request):
    user = auth.get_session_user(request)
    # Auth on → the gate guarantees a user here. Auth off → show every tool.
    tools = (user or {}).get("tools") if user else list(auth.TOOLS)
    return templates.TemplateResponse("index.html", {
        "request": request,
        "active":  "home",
        "user":    user,
        "tools":   tools,
    })

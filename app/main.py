import logging
import os
import secrets
import hashlib
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
    auth as auth_router, admin, sops, outbound_pulse, pulse_portal, credit_control,
    client_benchmarks, domain_cluster, gmail_setup, pulse_growth, pulse_meetings,
)

logger = logging.getLogger(__name__)

# ── Scheduled auto-sync ───────────────────────────────────────────────────────
# Runs once an hour for every campaign that has a linked Google Sheet, so the
# per-day send stats stay close to real time.  Campaigns are synced sequentially
# inside run_auto_sync (not in parallel) to stay under the Sheets read quota.
#
# Implemented as a plain daemon thread (see app/utils/auto_sync.py) rather than
# APScheduler — the event-loop-based scheduler proved unreliable under uvicorn
# (first run fired, hourly ticks then stopped). The thread runs once ~20s after
# boot, then at the top of every hour, for as long as the process is alive.

@asynccontextmanager
async def _lifespan(app: FastAPI):
    from app.utils.auto_sync import start_scheduler, stop_scheduler
    from app.utils.pulse.sync import (
        start_scheduler as start_pulse_scheduler,
        stop_scheduler as stop_pulse_scheduler,
    )
    try:
        start_scheduler()
        logger.info("Auto-sync scheduler started (hourly daemon thread; first run ~20s after boot).")
    except Exception as exc:
        logger.warning("Auto-sync scheduler could not start: %s", exc)

    # Outbound Pulse connector sync — same daemon-thread pattern, offset start.
    # No-ops when neither connector has an API key, so an unconfigured
    # deployment doesn't log an hourly failure.
    try:
        start_pulse_scheduler()
    except Exception as exc:
        logger.warning("Pulse sync scheduler could not start: %s", exc)

    from app.utils.gmail_setup.diagnostics import start_check
    start_check()

    yield  # app runs here

    try:
        stop_scheduler()
        logger.info("Auto-sync scheduler stopped.")
    except Exception:
        pass

    try:
        stop_pulse_scheduler()
    except Exception:
        pass


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
    # Bound domain-tool JSON before model parsing, including chunked requests.
    domain_tool = request.url.path.startswith(('/api/domain-cluster', '/api/gmail-setup'))
    if domain_tool and request.method == 'POST':
        chunks, total = [], 0
        async for chunk in request.stream():
            total += len(chunk)
            if total > 8_000_000:
                from fastapi.responses import JSONResponse
                return JSONResponse({'detail': 'Request exceeds 8 MB. Import in smaller batches.'}, status_code=413)
            chunks.append(chunk)
        request._body = b''.join(chunks)
    response = await call_next(request)
    if domain_tool or request.url.path in {'/', '/gmail-set', '/domain-cluster-builder',
            '/static/gmail-setup.js', '/static/gmail-setup.css', '/static/gmail-setup-home.js'}:
        response.headers['Cache-Control'] = 'no-store'
        response.headers['X-Content-Type-Options'] = 'nosniff'
    return response

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
GMAIL_HOME_ASSET_VERSION = hashlib.sha256((_BASE / 'static/gmail-setup-home.js').read_bytes()).hexdigest()[:12]
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
app.include_router(sops.router)
app.include_router(outbound_pulse.router)
app.include_router(pulse_portal.router)
app.include_router(pulse_growth.router)
app.include_router(pulse_meetings.router)
app.include_router(credit_control.router)
app.include_router(client_benchmarks.router)
app.include_router(gmail_setup.router)
app.include_router(domain_cluster.router)


@app.exception_handler(404)
async def not_found(request: Request, exc):
    """Styled 404 for page navigations; API callers still get JSON."""
    if request.url.path.startswith("/api") or "text/html" not in request.headers.get("accept", ""):
        from fastapi.responses import JSONResponse
        return JSONResponse({"detail": "Not Found"}, status_code=404)
    return templates.TemplateResponse(
        "404.html", {"request": request}, status_code=404,
    )


@app.get("/roi-calculator")
async def roi_calculator(request: Request):
    return templates.TemplateResponse("roi_calculator.html", {"request": request, "active": "roi_calculator"})


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
        "sop_key": "home",
        "user":    user,
        "tools":   tools,
        "gmail_home_asset_version": GMAIL_HOME_ASSET_VERSION,
    })


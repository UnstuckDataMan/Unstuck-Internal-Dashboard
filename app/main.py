import sys
from pathlib import Path

from fastapi import FastAPI, Request
from fastapi.staticfiles import StaticFiles
from starlette.middleware.wsgi import WSGIMiddleware

from app.deps import templates
from app.routers import gender, city

app = FastAPI(title="Data Enrichment Dashboard")

_BASE = Path(__file__).resolve().parent
app.mount("/static", StaticFiles(directory=str(_BASE / "static")), name="static")

app.include_router(gender.router)
app.include_router(city.router)

# ── Mail Merge Tool (Flask WSGI sub-application) ──────────────────────────────
# Add mail_merge/ to sys.path so its relative imports (utils.*) resolve correctly
_MM_PATH = str(_BASE.parent / "mail_merge")
if _MM_PATH not in sys.path:
    sys.path.insert(0, _MM_PATH)
from app import app as _flask_mail_merge          # noqa: E402  (mail_merge/app.py)
app.mount("/mail-merge", WSGIMiddleware(_flask_mail_merge))
# ─────────────────────────────────────────────────────────────────────────────


@app.get("/")
async def index(request: Request):
    return templates.TemplateResponse("index.html", {"request": request, "active": "home"})

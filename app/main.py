import sys
import importlib.util
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

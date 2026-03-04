from pathlib import Path

from fastapi import FastAPI, Request
from fastapi.staticfiles import StaticFiles

from app.deps import templates
from app.routers import gender, city

app = FastAPI(title="Data Enrichment Dashboard")

_BASE = Path(__file__).resolve().parent
app.mount("/static", StaticFiles(directory=str(_BASE / "static")), name="static")

app.include_router(gender.router)
app.include_router(city.router)


@app.get("/")
async def index(request: Request):
    return templates.TemplateResponse("index.html", {"request": request, "active": "home"})

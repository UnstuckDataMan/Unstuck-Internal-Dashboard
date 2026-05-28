from fastapi import APIRouter, Request
from app.deps import templates

router = APIRouter()


@router.get("/launch-checker")
async def launch_checker(request: Request):
    return templates.TemplateResponse("launch_checker.html", {"request": request})

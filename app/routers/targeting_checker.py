from fastapi import APIRouter, Request

from app.deps import templates

router = APIRouter()


@router.get("/targeting-checker")
async def targeting_checker(request: Request):
    return templates.TemplateResponse("targeting_checker.html", {"request": request, "active": "targeting_checker", "sop_key": "targeting_checker"})

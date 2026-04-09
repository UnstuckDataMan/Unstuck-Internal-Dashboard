import os

from fastapi import APIRouter, Request

from app.deps import templates

router = APIRouter()


@router.get("/reply-bank")
async def reply_bank(request: Request):
    return templates.TemplateResponse("reply_bank.html", {
        "request":          request,
        "supabase_url":     os.environ["SUPABASE_URL"],
        "supabase_anon_key": os.environ["SUPABASE_ANON_KEY"],
    })

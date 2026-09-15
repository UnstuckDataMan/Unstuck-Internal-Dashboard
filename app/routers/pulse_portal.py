"""
Client-facing portal — read-only funnel for one client, reached by magic link.

This is the only part of the dashboard a non-staff user ever sees, so it is the
only part with its own access rules. Three properties matter:

  1. **Read-only by construction.** The router exposes GET routes only. There is
     no mutation a client could reach even with a valid token.

  2. **Token-scoped, not user-scoped.** A token resolves to exactly one
     client_id, and every query is filtered by that id (and the agency id) taken
     from the token record — never from the URL or a form field. A client cannot
     widen their own scope by editing a parameter, because no parameter feeds
     the scope.

  3. **Not indexable, not cacheable.** A report link that ends up in a browser's
     shared cache or a search index is a data leak, so responses carry
     `noindex` and `no-store`.

Tokens are compared by SHA-256 hash — the plaintext never touches the database.
"""
from __future__ import annotations

import hashlib
from datetime import datetime, timedelta, timezone

from fastapi import APIRouter, Query, Request
from fastapi.responses import HTMLResponse, Response

from app.deps import templates
from app.utils.dates import today_utc
from app.utils.pulse import store
from app.utils.pulse import template_filters
from app.utils.pulse.normalize import EVENT_SENT, funnel_with_rates
from app.utils.pulse.store import PulseNotReady

router = APIRouter()
template_filters.register(templates.env)

# Ranges a client can pick. Deliberately a fixed allowlist rather than free
# dates: it keeps the portal simple, and it means no client-supplied string
# reaches a query builder.
PORTAL_RANGES: dict[str, str] = {
    "30d": "Last 30 days",
    "90d": "Last 90 days",
    "all": "All time",
}


def _hash_token(token: str) -> str:
    return hashlib.sha256(token.encode("utf-8")).hexdigest()


def _resolve_access(token: str) -> tuple[dict | None, str]:
    """Validate a portal token. Returns (access_row, error_message).

    Every rejection returns the same generic message — a revoked token and a
    token that never existed must be indistinguishable from the outside.
    """
    token = (token or "").strip()
    if not token or len(token) > 200:
        return None, "This link is not valid."

    access = store.access_by_token_hash(_hash_token(token))
    if not access:
        return None, "This link is not valid."
    if access.get("revoked_at"):
        return None, "This link is no longer active."

    expires = access.get("expires_at")
    if expires:
        try:
            stamp = datetime.fromisoformat(str(expires).replace("Z", "+00:00"))
            if stamp.tzinfo is None:
                stamp = stamp.replace(tzinfo=timezone.utc)
            if stamp < datetime.now(timezone.utc):
                return None, "This link has expired. Ask your account manager for a new one."
        except ValueError:
            pass
    return access, ""


def _range_bounds(key: str):
    today = today_utc()
    if key == "all":
        return None, None
    if key == "90d":
        return today - timedelta(days=89), today
    return today - timedelta(days=29), today


def _private(response: Response) -> Response:
    response.headers["X-Robots-Tag"] = "noindex, nofollow, noarchive"
    response.headers["Cache-Control"] = "no-store, private"
    return response


def _denied(request: Request, message: str) -> HTMLResponse:
    response = templates.TemplateResponse(
        "portal_denied.html", {"request": request, "message": message}, status_code=404,
    )
    return _private(response)


@router.get("/portal/{token}")
async def portal(request: Request, token: str, range: str = Query("30d")):
    try:
        access, error = _resolve_access(token)
    except PulseNotReady:
        # Never leak infrastructure detail to a client-facing page.
        return _denied(request, "This report is temporarily unavailable. "
                                "Please try again shortly.")
    if access is None:
        return _denied(request, error)

    client_id = str(access["client_id"])
    range_key = range if range in PORTAL_RANGES else "30d"
    date_from, date_to = _range_bounds(range_key)

    try:
        client = next(
            (c for c in store.list_clients() if str(c["id"]) == client_id), None,
        )
        if client is None:
            return _denied(request, "This link is not valid.")

        counts     = store.funnel(client_id=client_id, date_from=date_from, date_to=date_to)
        by_channel = store.funnel_by_channel(client_id=client_id,
                                             date_from=date_from, date_to=date_to)
        timeseries = store.funnel_timeseries(client_id=client_id,
                                             date_from=date_from, date_to=date_to)
    except PulseNotReady:
        return _denied(request, "This report is temporarily unavailable. "
                                "Please try again shortly.")

    # Engagement logging — best-effort inside the store helper, so a failure
    # here never costs the client their report.
    store.record_visit(access, client_id, request.headers.get("user-agent", ""))

    response = templates.TemplateResponse("portal.html", {
        "request":     request,
        "client":      client,
        "token":       token,
        "counts":      counts,
        "funnel":      funnel_with_rates(counts),
        "by_channel":  by_channel,
        "timeseries":  timeseries,
        "total_sent":  counts.get(EVENT_SENT, 0),
        "ranges":      PORTAL_RANGES,
        "range_key":   range_key,
        "range_label": PORTAL_RANGES[range_key],
        "generated":   datetime.now(timezone.utc),
    })
    return _private(response)

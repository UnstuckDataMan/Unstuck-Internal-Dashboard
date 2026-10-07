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
from datetime import datetime, timezone

from fastapi import APIRouter, Query, Request
from fastapi.responses import HTMLResponse, Response

from app.deps import templates
from app.utils.dates import today_utc
from app.utils.pulse import store
from app.utils.pulse import template_filters
from app.utils.pulse.reports import report_heading
from app.utils.pulse.normalize import (
    CHANNEL_LINKEDIN,
    EVENT_SENT,
    funnel_with_rates,
    without,
)
from app.utils.pulse.store import PulseNotReady

router = APIRouter()
template_filters.register(templates.env)

# Ranges a client could pick are gone. A report covers the period its account
# manager chose, and nothing else: letting a client re-slice the data was how
# they ended up reading numbers nobody had looked at before sending.


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


def _private(response: Response) -> Response:
    response.headers["X-Robots-Tag"] = "noindex, nofollow, noarchive"
    response.headers["Cache-Control"] = "no-store, private"
    return response


def _denied(request: Request, message: str) -> HTMLResponse:
    response = templates.TemplateResponse(
        "portal_denied.html", {"request": request, "message": message}, status_code=404,
    )
    return _private(response)


def _snapshot_view(report: dict) -> dict:
    """Unpack a published report's frozen figures for the template.

    A published report renders from this and never re-queries. The numbers a
    client was sent are the numbers they see when they open it again — a later
    sync can still add events inside a closed period, and "that figure has
    moved since you sent it" is the conversation this exists to avoid.

    Tolerant of a snapshot that predates a field: a report published by an
    older version renders with what it has rather than failing.
    """
    snapshot = report.get("snapshot") or {}
    counts = snapshot.get("counts") or {}

    # LINKEDIN IS NOT PART OF THE EMAIL FUNNEL. A connection request is not an
    # email, and an acceptance is not an open, so adding them produced a "sent"
    # a client could not reconcile with anything and a reply rate that was the
    # average of two unlike platforms. LinkedIn gets its own card instead.
    #
    # By subtraction, so a row carrying no channel still reaches the client.
    # Every snapshot has carried by_channel from the start, so reports
    # published before this split render correctly too.
    linkedin = (snapshot.get("by_channel") or {}).get(CHANNEL_LINKEDIN) or {}
    email = without(counts, linkedin)
    return {
        "counts":     email,
        "funnel":     funnel_with_rates(email),
        "linkedin":   linkedin,
        # A client who has never been run on LinkedIn gets no card at all,
        # rather than a card of zeros implying we tried and nothing happened.
        "has_linkedin": any((linkedin.get(k) or 0) > 0 for k in linkedin),
        "trend":      snapshot.get("trend") or {"unit": "day", "buckets": []},
        # Absent from any report published before A/B was captured, which is
        # why every field here is read defensively rather than indexed.
        "ab":         snapshot.get("ab") or [],
        # Absent from every report published before lead quality existed, so
        # read for what it is rather than indexed.
        "icp":        snapshot.get("icp") or {},
        # Gate on leads rather than on the key: a period the sheet covers but
        # has nothing in gets no card, instead of a card of zeros that reads
        # as a campaign that failed.
        "has_icp":    bool((snapshot.get("icp") or {}).get("total")),
        # Both platforms: this is the "did anything happen at all" gate, and a
        # client run only on LinkedIn must not be told there was no activity.
        "total_sent": counts.get(EVENT_SENT, 0) or 0,
        # Email alone, for whether the email half of the page has anything to
        # say. A funnel of zeros above a live LinkedIn card reads as a failure.
        "email_sent": email.get(EVENT_SENT, 0) or 0,
    }


# /r/ is the short form a client is actually sent. /portal/ stays because
# links already handed out use it, and a report link that stops working is a
# client emailing their account manager about a broken report.
@router.get("/r/{token}")
@router.get("/portal/{token}")
async def portal(request: Request, token: str, report: str = Query("")):
    """One link per client, for good.

    It opens the newest published report; `report` selects an earlier one, and
    the arrows page through the same ordering. A report id is only honoured if
    it belongs to this token's client, so the parameter cannot widen scope.
    """
    try:
        access, error = _resolve_access(token)
    except PulseNotReady:
        # Never leak infrastructure detail to a client-facing page.
        return _denied(request, "This report is temporarily unavailable. "
                                "Please try again shortly.")
    if access is None:
        return _denied(request, error)

    client_id = str(access["client_id"])
    try:
        client = next(
            (c for c in store.list_clients() if str(c["id"]) == client_id), None,
        )
        if client is None:
            return _denied(request, "This link is not valid.")
    except PulseNotReady:
        return _denied(request, "This report is temporarily unavailable. "
                                "Please try again shortly.")

    # Newest first, which is the order the arrows page in. Scoped to this
    # token's client, so `report` can only ever select from their own history.
    reports = store.published_reports(client_id)

    index = 0
    if report:
        for position, row in enumerate(reports):
            if str(row.get("id")) == report:
                index = position
                break

    # Engagement logging — best-effort inside the store helper, so a failure
    # here never costs the client their report.
    store.record_visit(access, client_id, request.headers.get("user-agent", ""))

    path = "/r" if request.url.path.startswith("/r/") else "/portal"
    current = reports[index] if reports else None
    context = {
        "request":     request,
        "client":      client,
        "token":       token,
        # Keep the reader on the path they arrived by, so the arrows do not
        # silently move a /r/ link onto /portal/.
        "portal_path": path,
        "report":      current,
        "heading":     report_heading(current) if current else "",
        # Newer is earlier in the list, so "previous" is the higher index.
        "newer":       reports[index - 1] if index > 0 else None,
        "older":       reports[index + 1] if index + 1 < len(reports) else None,
        "count":       len(reports),
        # Counted so the NEWEST is the highest number: a client landing on
        # their latest report should read "4 of 4", not "1 of 4", which looked
        # like they were at the beginning of their history rather than the end.
        # The arrows agree with it — back is a lower number, forward a higher.
        "position":    len(reports) - index if reports else 0,
        "generated":   datetime.now(timezone.utc),
    }
    context.update(_snapshot_view(current) if current else {
        "counts": {}, "funnel": [], "linkedin": {}, "has_linkedin": False,
        "trend": {"unit": "day", "buckets": []}, "ab": [], "icp": {},
        "has_icp": False, "total_sent": 0, "email_sent": 0,
    })
    return _private(templates.TemplateResponse("portal.html", context))

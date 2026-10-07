"""Reading the Copy Bank from outside the Copy Bank.

Two tools need the same two things and used to have neither in common: how a
variant label maps onto the stored arrays, and how to find the stored arrays
in the first place. Copy Bank itself needs them to badge a winning card;
Client Performance & Reports needs them to show what the winning variation
actually said.

THE LABEL IS AN INDEX, NOT A NAME. A manual campaign's "A/B Variant" column
holds a combination like "S2/B1", meaning subject 2 with body 1 of whatever
Copy Bank entry the merge pulled. The label carries no text at all, which is
why a winning variation can be known for weeks without anyone being able to
read it. Resolving it needs the entry as well as the label.

WHICH ENTRY, THOUGH. A campaign records the entry it pulled in
campaigns.copy_territory / copy_industry, so the pair identifies the copy. It
is optional — a merge run without choosing a copy source stores neither — so
callers must cope with a label that cannot be resolved rather than assuming
every campaign has one.
"""
from __future__ import annotations

import logging
import re

import requests as http_req

from app.utils.supabase import SUPABASE_URL, sb_headers

logger = logging.getLogger(__name__)

# Clients are namespaced; Biz Dev content predates that and is stored bare.
_CLIENT_PREFIX = "__c__"

_VARIANT_RE = re.compile(r"^\s*S(\d+)\s*/\s*B(\d+)\s*$")


def parse_variant(label: str) -> tuple[int | None, int | None]:
    """'S2/B1' → (subject_idx=1, body_idx=0), 0-based to match the arrays.

    (None, None) for anything else — Smartlead's variant labels are letters,
    and a sheet can hold whatever somebody typed.
    """
    m = _VARIANT_RE.match(str(label or ""))
    if not m:
        return (None, None)
    return (int(m.group(1)) - 1, int(m.group(2)) - 1)


def template_key(client_id: str, territory: str, industry: str) -> str:
    if client_id == "bizdev":
        return f"{territory}_{industry}"
    return f"{_CLIENT_PREFIX}{client_id}__{territory}_{industry}"


def fetch_content(client_id: str, territory: str, industry: str) -> tuple[dict, bool]:
    """The stored Copy Bank entry, as (content, found).

    Falls back to the bare Biz Dev key when a client has none of their own,
    which is how profiles migrated from Biz Dev still resolve.

    `found` is not the same as a non-empty content: a row that exists with
    nothing in it is a deliberate blank, and the caller's fallbacks differ
    between the two. Never raises — this is always decoration on top of
    something that already works.
    """
    if not SUPABASE_URL:
        return {}, False
    for key in _candidate_keys(client_id, territory, industry):
        try:
            r = http_req.get(
                f"{SUPABASE_URL}/rest/v1/copy_bank_templates",
                params={"key": f"eq.{key}", "select": "content"},
                headers=sb_headers(),
                timeout=10,
            )
            if r.status_code >= 400:
                continue
            rows = r.json()
        except Exception as exc:
            logger.warning("Copy Bank: could not read %s: %s", key, exc)
            return {}, False
        if rows:
            return (rows[0].get("content") or {}), True
    return {}, False


def _candidate_keys(client_id: str, territory: str, industry: str) -> list[str]:
    key = template_key(client_id, territory, industry)
    if client_id == "bizdev":
        return [key]
    return [key, f"{territory}_{industry}"]


def extract(content: dict, channel: str = "email") -> dict:
    """The subject lines and bodies of one channel, blanks dropped.

    LinkedIn and chaser copy is an ordered body-only sequence; email and
    flyout have subject lines against body variations. Dropping blanks here
    rather than at the call site matters more than it looks: the index in a
    variant label counts the copy a person can see in Copy Bank, and Copy Bank
    does not show empty cards.
    """
    content = content or {}
    if channel in ("linkedin", "chaser"):
        steps = (content.get(channel) or {}).get("steps") or []
        return {
            "subjects": [],
            "bodies":   [s["body"] for s in steps if (s.get("body") or "").strip()],
        }
    ch = content.get("flyout" if channel == "flyout" else "email") or {}
    return {
        "subjects": [s for s in (ch.get("subjects") or []) if s and s.strip()],
        "bodies":   [v["body"] for v in (ch.get("variations") or [])
                     if (v.get("body") or "").strip()],
    }


def variant_in(copy: dict, label: str) -> dict | None:
    """What "S2/B1" said, given arrays already fetched by `extract`.

    Split out from `variant_copy` so a caller resolving a whole table of
    variants fetches the entry once instead of once per row.

    None covers every reason equally — an unparseable label, an index past the
    end of the arrays — because the caller does the same thing in all of them:
    offer the copy to be written down by hand instead. The index overrunning
    is the interesting one, and it happens: copy edited after a campaign went
    out renumbers everything below it.
    """
    subject_idx, body_idx = parse_variant(label)
    if subject_idx is None or subject_idx < 0 or body_idx is None or body_idx < 0:
        return None
    subjects, bodies = copy.get("subjects") or [], copy.get("bodies") or []
    # BOTH halves have to land. Half a label is not a near miss: showing body 1
    # under a blank subject would put the right message against a subject line
    # nobody was sent, which reads as fact on a client's report.
    if subject_idx >= len(subjects) or body_idx >= len(bodies):
        return None
    subject, body = subjects[subject_idx], bodies[body_idx]
    if not (subject or "").strip() and not (body or "").strip():
        return None
    return {"subject": (subject or "").strip(), "body": body or ""}


def variant_copy(client_id: str, territory: str, industry: str,
                 label: str, channel: str = "email") -> dict | None:
    """One-shot resolve, for a caller with a single label to look up."""
    content, found = fetch_content(client_id, territory, industry)
    if not found:
        return None
    return variant_in(extract(content, channel), label)

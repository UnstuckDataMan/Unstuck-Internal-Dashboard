"""
Read AI webhook handling with no I/O: signature check, payload parsing,
client matching and onboarding detection. Testable without a network or a
database; the router and meeting_insights do the reading and writing.

Read AI's meeting_end payload carries session_id, request_id, trigger, title,
start_time, end_time, owner, participants [{name, email}], summary,
action_items, key_questions, topics, chapter_summaries, report_url and
transcript (speaker blocks). The exact shapes are read defensively: a field
that arrives as a string instead of an object, or not at all, is not a reason
to drop a client conversation on the floor.
"""
from __future__ import annotations

import base64
import binascii
import hashlib
import hmac
import os
import re
from datetime import datetime, timezone

# Domains that never identify a client: personal mail, and our own.
FREE_MAIL_DOMAINS = frozenset({
    "gmail.com", "googlemail.com", "outlook.com", "hotmail.com", "hotmail.co.uk",
    "live.com", "live.co.uk", "msn.com", "yahoo.com", "yahoo.co.uk", "icloud.com",
    "me.com", "mac.com", "aol.com", "proton.me", "protonmail.com", "gmx.com",
    "zoho.com", "mail.com", "btinternet.com", "sky.com", "virginmedia.com",
})

_ONBOARDING_RE = re.compile(r"\b(on[\s-]?boarding|kick[\s-]?off)\b", re.IGNORECASE)


def agency_domains() -> set[str]:
    raw = os.environ.get("AGENCY_EMAIL_DOMAINS", "") or os.environ.get(
        "ALLOWED_EMAIL_DOMAIN", "unstuck-agency.com")
    return {d for d in (normalise_domain(p) for p in raw.split(",")) if d}


# ── Signature ─────────────────────────────────────────────────────────────────

def verify_signature(raw: bytes, header: str, secret: str) -> bool:
    """HMAC-SHA256 of the raw body, keyed with the webhook's signing key.

    Read AI's sample code base64-decodes the key before signing and compares
    hex digests. The raw key is tried as well, in case the key shown in Read
    AI is ever not base64; both still require knowing the secret.
    """
    if not raw or not header or not secret:
        return False
    given = header.strip()
    if given.lower().startswith("sha256="):
        given = given[7:]
    given = given.strip().lower()
    keys = [secret.encode()]
    try:
        keys.insert(0, base64.b64decode(secret, validate=True))
    except (binascii.Error, ValueError):
        pass
    for key in keys:
        digest = hmac.new(key, raw, hashlib.sha256).hexdigest()
        if hmac.compare_digest(digest, given):
            return True
    return False


# ── Payload ───────────────────────────────────────────────────────────────────

def _text(value) -> str:
    if value is None:
        return ""
    if isinstance(value, str):
        return value.strip()
    if isinstance(value, dict):
        for key in ("text", "summary", "content", "words", "name"):
            if isinstance(value.get(key), str):
                return value[key].strip()
        return ""
    return str(value).strip()


def _when(value) -> str | None:
    """An ISO timestamp from either an ISO string or Unix milliseconds."""
    if value in (None, ""):
        return None
    try:
        if isinstance(value, (int, float)) or str(value).isdigit():
            number = float(value)
            if number > 1e11:  # milliseconds
                number /= 1000.0
            return datetime.fromtimestamp(number, tz=timezone.utc).isoformat()
        text = str(value).strip().replace("Z", "+00:00")
        parsed = datetime.fromisoformat(text)
        if parsed.tzinfo is None:
            parsed = parsed.replace(tzinfo=timezone.utc)
        return parsed.isoformat()
    except (ValueError, OverflowError, OSError):
        return None


def _ms(value) -> int | None:
    try:
        number = int(float(value))
    except (TypeError, ValueError):
        return None
    return number


def _person(value) -> dict:
    if isinstance(value, dict):
        return {"name": _text(value.get("name")),
                "email": _text(value.get("email")).lower()}
    if isinstance(value, str) and "@" in value:
        return {"name": "", "email": value.strip().lower()}
    return {"name": _text(value), "email": ""}


def parse_transcript(value) -> list[dict]:
    """Speaker blocks as [{"i", "speaker", "start_ms", "end_ms", "text"}]."""
    blocks = value
    if isinstance(value, dict):
        blocks = value.get("speaker_blocks") or value.get("blocks") or []
    if not isinstance(blocks, list):
        return []
    out: list[dict] = []
    for block in blocks:
        if not isinstance(block, dict):
            continue
        text = _text(block.get("words") if "words" in block else block.get("text"))
        if not text:
            continue
        speaker = block.get("speaker")
        out.append({
            "i": len(out),
            "speaker": _text(speaker if not isinstance(speaker, dict) else speaker.get("name"))
                       or "Unknown speaker",
            "start_ms": _ms(block.get("start_time")),
            "end_ms": _ms(block.get("end_time")),
            "text": text,
        })
    return out


def parse_payload(payload: dict) -> dict:
    """The fields of a pulse_meetings row that come from Read AI."""
    if not isinstance(payload, dict):
        raise ValueError("The webhook body is not a JSON object.")
    session_id = _text(payload.get("session_id"))
    if not session_id:
        raise ValueError("The webhook body has no session_id.")
    participants = [p for p in (_person(v) for v in (payload.get("participants") or [])
                                if v) if p["name"] or p["email"]]
    actions = payload.get("action_items") or []
    return {
        "session_id": session_id,
        "request_id": _text(payload.get("request_id")),
        "trigger": _text(payload.get("trigger")) or "meeting_end",
        "title": _text(payload.get("title"))[:300] or "Untitled meeting",
        "start_time": _when(payload.get("start_time")),
        "end_time": _when(payload.get("end_time")),
        "owner_email": _person(payload.get("owner")).get("email", ""),
        "participants": participants,
        "report_url": _text(payload.get("report_url")),
        "read_summary": _text(payload.get("summary")),
        "action_items": [t for t in (_text(a) for a in actions) if t]
                        if isinstance(actions, list) else [],
        "transcript": parse_transcript(payload.get("transcript")),
    }


# ── Matching ──────────────────────────────────────────────────────────────────

def normalise_domain(value: str) -> str:
    """acme.com from "https://www.Acme.com/about", "jo@acme.com" or "acme.com"."""
    text = (value or "").strip().lower()
    if "@" in text:
        text = text.rsplit("@", 1)[1]
    text = re.sub(r"^[a-z]+://", "", text)
    text = text.split("/", 1)[0].split("?", 1)[0].split(":", 1)[0].strip(".")
    if text.startswith("www."):
        text = text[4:]
    return text if "." in text and " " not in text else ""


def external_domains(participants: list[dict]) -> set[str]:
    ours = agency_domains()
    out = set()
    for p in participants or []:
        domain = normalise_domain(p.get("email", ""))
        if domain and domain not in FREE_MAIL_DOMAINS and domain not in ours:
            out.add(domain)
    return out


def _covers(client_domain: str, domain: str) -> bool:
    return domain == client_domain or domain.endswith("." + client_domain)


def match_client(participants: list[dict], clients: list[dict]) -> tuple[str, list[str]]:
    """("matched"|"ambiguous"|"unmatched", candidate client ids)."""
    domains = external_domains(participants)
    hits: list[str] = []
    for client in clients:
        owned = {normalise_domain(d) for d in (client.get("domains") or [])} - {""}
        if any(_covers(c, d) for c in owned for d in domains):
            hits.append(str(client["id"]))
    if len(hits) == 1:
        return "matched", hits
    if hits:
        return "ambiguous", hits
    return "unmatched", []


def is_internal(participants: list[dict]) -> bool:
    """True only when every attendee has an email at one of our own domains.

    An attendee with no email is unknown, not internal: Read AI often has
    only a name for people who joined without the calendar invite, and a
    client call made of those must reach the inbox rather than vanish.
    """
    ours = agency_domains()
    people = [p for p in participants or [] if p.get("name") or p.get("email")]
    return bool(people) and all(normalise_domain(p.get("email", "")) in ours for p in people)


def title_candidates(title: str, clients: list[dict]) -> list[str]:
    """Clients named in the meeting title, as suggestions for the inbox."""
    text = f" {re.sub(r'[^a-z0-9]+', ' ', (title or '').lower())} "
    hits = []
    for client in clients:
        name = re.sub(r"[^a-z0-9]+", " ", str(client.get("name") or "").lower()).strip()
        if len(name) >= 3 and f" {name} " in text:
            hits.append(str(client["id"]))
    return hits


def looks_like_onboarding(title: str) -> bool:
    return bool(_ONBOARDING_RE.search(title or ""))

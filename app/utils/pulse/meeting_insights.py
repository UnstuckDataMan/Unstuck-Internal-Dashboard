"""
What the AI does with a Read AI meeting, and when.

Three jobs, all run in daemon threads so no request ever waits on the model:

1. analyse_meeting   one meeting -> what it says went well and badly with the
                     campaigns, each point tied to a transcript block and a
                     verbatim quote.
2. rebuild_overview  every analysed meeting for a client -> 3-5 successes and
                     3-5 pain points. The model cites the per-meeting points
                     by id rather than writing new quotes, so every link it
                     returns lands on a passage that was already checked.
3. draft_onboarding_notes  the onboarding meeting -> notes under fixed
                     headings. Fills untouched notes; once a person has
                     edited them, the draft waits for them to accept it.

Grounding rule: a point whose quote cannot be found in the stored transcript
is dropped. A summary that sends an account manager to a passage that does
not say what the summary claims is worse than a shorter summary.
"""
from __future__ import annotations

import logging
import re
import threading
from html import escape

from app.utils import llm
from app.utils.pulse import meetings_store as ms
from app.utils.pulse import readai, richtext

logger = logging.getLogger(__name__)

# Roughly 30k tokens of transcript: an hour-long call fits, and one that does
# not is cut at a block boundary with the cut stated in the prompt.
TRANSCRIPT_CHAR_BUDGET = 120_000
MAX_POINTS_PER_SIDE = 8

ONBOARDING_HEADINGS = [
    "Offer and services",
    "ICP and targeting",
    "Goals and KPIs",
    "Exclusions and DNC",
    "Messaging and tone",
    "Contacts and approvals",
    "Key dates",
    "Other",
]

_CONTEXT = (
    "Unstuck is an outbound lead-generation agency. It runs cold email and LinkedIn "
    "campaigns for its clients and books meetings and leads for them. The transcript is "
    "a recorded meeting between Unstuck's team and one client. Write in plain UK "
    "English, short and specific, naming figures, segments and examples when the "
    "speakers give them."
)


# ── Transcript helpers ────────────────────────────────────────────────────────

def _squash(text: str) -> str:
    return re.sub(r"\s+", " ", (text or "")).strip().lower()


def transcript_for_prompt(blocks: list[dict]) -> tuple[str, bool]:
    """Numbered blocks within the budget, and whether any were cut."""
    lines, used = [], 0
    for block in blocks:
        line = f"[{block['i']}] {block.get('speaker') or 'Unknown speaker'}: {block.get('text', '')}"
        if used + len(line) > TRANSCRIPT_CHAR_BUDGET and lines:
            return "\n".join(lines), True
        lines.append(line)
        used += len(line) + 1
    return "\n".join(lines), False


def locate_quote(blocks: list[dict], block: int, quote: str) -> int | None:
    """The index of the block that really contains `quote`, or None.

    The model's block number is tried first, then its neighbours (speaker
    turns are often split mid-sentence), then the whole transcript.
    """
    needle = _squash(quote).strip(" .,\"'…")
    if len(needle) < 8:
        return None
    by_index = {b["i"]: b for b in blocks}
    for candidate in (block, block - 1, block + 1):
        if candidate in by_index and needle in _squash(by_index[candidate]["text"]):
            return candidate
    for b in blocks:
        if needle in _squash(b["text"]):
            return b["i"]
    return None


# ── 1. One meeting ────────────────────────────────────────────────────────────

_POINT = {
    "type": "object",
    "additionalProperties": False,
    "required": ["point", "block", "quote"],
    "properties": {
        "point": {"type": "string"},
        "block": {"type": "integer"},
        "quote": {"type": "string"},
    },
}

MEETING_SCHEMA = {
    "type": "object",
    "additionalProperties": False,
    "required": ["about_campaigns", "summary", "positives", "negatives"],
    "properties": {
        "about_campaigns": {"type": "boolean"},
        "summary": {"type": "string"},
        "positives": {"type": "array", "items": _POINT},
        "negatives": {"type": "array", "items": _POINT},
    },
}

_MEETING_PROMPT = _CONTEXT + """

Read the transcript and pull out what it says about how the client's outbound
campaigns are performing.

- positives: what is working. Lead quality, meetings booked, replies, targeting
  that landed, copy that resonated, the client's satisfaction.
- negatives: what is not. Poor lead quality, wrong targets, low volume, slow
  replies, copy problems, deliverability, frustrations, risks to the account.
- Each point is one sentence of up to 25 words, stated as a finding.
- block is the number in square brackets of the passage that supports it.
- quote is copied word for word from that passage, 6 to 30 words. Never
  paraphrase a quote.
- Only include points the speakers actually make. No advice, no speculation.
- summary is two sentences on the meeting as a whole.
- about_campaigns is false when the meeting says nothing about campaign
  performance (for example an internal call); then leave both lists empty.
"""


def _grounded(points, blocks: list[dict]) -> list[dict]:
    out = []
    for p in points or []:
        try:
            point = str(p["point"]).strip()
            quote = str(p["quote"]).strip()
            block = int(p["block"])
        except (KeyError, TypeError, ValueError):
            continue
        found = locate_quote(blocks, block, quote)
        if point and found is not None:
            out.append({"point": point[:400], "block": found, "quote": quote[:400]})
        if len(out) >= MAX_POINTS_PER_SIDE:
            break
    return out


def analyse_meeting(meeting_id: str) -> dict | None:
    """Analyse one meeting and store the result. Returns the analysis."""
    meeting = ms.get_meeting(meeting_id)
    if not meeting:
        return None
    blocks = meeting.get("transcript") or []
    if not blocks:
        ms.update_meeting(meeting_id, {
            "analysis_status": "skipped",
            "analysis_error": "Read AI sent no transcript for this meeting.",
            "analysis": {}, "analysed_at": ms.now_iso()})
        return None
    if not llm.is_configured():
        ms.update_meeting(meeting_id, {
            "analysis_status": "failed",
            "analysis_error": "OpenAI isn't connected, so this meeting couldn't be summarised."})
        return None

    ms.update_meeting(meeting_id, {"analysis_status": "running", "analysis_error": ""})
    text, cut = transcript_for_prompt(blocks)
    user = (f"Meeting: {meeting.get('title') or 'Untitled'}\n"
            f"Read AI's own summary: {meeting.get('read_summary') or '(none)'}\n\n"
            f"Transcript{' (cut short: the end of the meeting is missing)' if cut else ''}:\n{text}")
    try:
        raw = llm.complete_json(_MEETING_PROMPT, user, schema=MEETING_SCHEMA,
                                schema_name="meeting_analysis")
    except llm.LLMError as exc:
        ms.update_meeting(meeting_id, {"analysis_status": "failed", "analysis_error": str(exc)})
        return None

    analysis = {
        "about_campaigns": bool(raw.get("about_campaigns")),
        "summary": str(raw.get("summary") or "").strip()[:800],
        "positives": _grounded(raw.get("positives"), blocks),
        "negatives": _grounded(raw.get("negatives"), blocks),
        "truncated": cut,
    }
    ms.update_meeting(meeting_id, {"analysis": analysis, "analysis_status": "done",
                                   "analysis_error": "", "analysed_at": ms.now_iso()})
    return analysis


# ── 2. The client overview ────────────────────────────────────────────────────

_THEME = {
    "type": "object",
    "additionalProperties": False,
    "required": ["point", "sources"],
    "properties": {
        "point": {"type": "string"},
        "sources": {"type": "array", "items": {"type": "string"}},
    },
}

OVERVIEW_SCHEMA = {
    "type": "object",
    "additionalProperties": False,
    "required": ["summary", "successes", "pain_points"],
    "properties": {
        "summary": {"type": "string"},
        "successes": {"type": "array", "items": _THEME},
        "pain_points": {"type": "array", "items": _THEME},
    },
}

_OVERVIEW_PROMPT = _CONTEXT + """

You are given the findings from every recorded meeting with this client, each
with an id such as M2.N1 (meeting 2, negative 1) or M1.P3 (meeting 1, positive 3).

Write an overview for the account manager:
- summary: at most three sentences on how the account is going and where it is heading.
- successes: the 3 to 5 most important things going well.
- pain_points: the 3 to 5 most important problems, the most serious first.
- Merge findings that say the same thing. Prefer recent meetings when they
  conflict, and say so if a problem has been resolved.
- Each item is one sentence of up to 25 words.
- sources lists the ids of every finding the item is based on. Use only ids
  from the input. Every item needs at least one.
- Fewer items is fine when the meetings support fewer.
"""


def _findings_for_prompt(meetings: list[dict]) -> tuple[str, dict]:
    """The prompt text, and a map from finding id to its transcript ref."""
    lines, refs = [], {}
    for n, m in enumerate(meetings, start=1):
        analysis = m.get("analysis") or {}
        lines.append(f"Meeting M{n}: {m.get('title') or 'Untitled'} "
                     f"({(m.get('start_time') or '')[:10] or 'date unknown'})")
        if analysis.get("summary"):
            lines.append(f"  Summary: {analysis['summary']}")
        for side, letter in (("positives", "P"), ("negatives", "N")):
            for k, p in enumerate(analysis.get(side) or [], start=1):
                fid = f"M{n}.{letter}{k}"
                refs[fid] = {"meeting_id": str(m["id"]), "block": p.get("block"),
                             "quote": p.get("quote", "")}
                lines.append(f"  {fid}: {p.get('point')}")
    return "\n".join(lines), refs


def _themes(items, refs: dict) -> list[dict]:
    out = []
    for item in items or []:
        try:
            point = str(item["point"]).strip()
            sources = [str(s).strip() for s in item.get("sources") or []]
        except (KeyError, TypeError, AttributeError):
            continue
        cited, seen = [], set()
        for s in sources:
            if s in refs and s not in seen:
                seen.add(s)
                cited.append(refs[s])
        if point and cited:
            out.append({"point": point[:400], "refs": cited[:6]})
        if len(out) >= 6:
            break
    return out


def rebuild_overview(client_id: str) -> dict | None:
    meetings = [m for m in ms.list_client_meetings(client_id)
                if m.get("analysis_status") == "done"
                and ((m.get("analysis") or {}).get("positives")
                     or (m.get("analysis") or {}).get("negatives"))]
    # Oldest first, so "M1" is the first meeting and recency reads naturally.
    meetings.sort(key=lambda m: m.get("start_time") or m.get("received_at") or "")
    ids = [str(m["id"]) for m in meetings]
    if not meetings:
        ms.save_insights(client_id, {"overview": {}, "status": "done", "error": "",
                                     "meeting_ids": ids, "generated_at": ms.now_iso()})
        return {}
    if not llm.is_configured():
        ms.save_insights(client_id, {"status": "failed",
                                     "error": "OpenAI isn't connected, so the overview couldn't be written."})
        return None

    ms.save_insights(client_id, {"status": "running", "error": ""})
    text, refs = _findings_for_prompt(meetings)
    try:
        raw = llm.complete_json(_OVERVIEW_PROMPT, text, schema=OVERVIEW_SCHEMA,
                                schema_name="client_overview", max_tokens=2000)
    except llm.LLMError as exc:
        ms.save_insights(client_id, {"status": "failed", "error": str(exc)})
        return None
    overview = {
        "summary": str(raw.get("summary") or "").strip()[:1200],
        "successes": _themes(raw.get("successes"), refs),
        "pain_points": _themes(raw.get("pain_points"), refs),
    }
    ms.save_insights(client_id, {"overview": overview, "status": "done", "error": "",
                                 "meeting_ids": ids, "generated_at": ms.now_iso()})
    return overview


# ── 3. Onboarding notes ───────────────────────────────────────────────────────

NOTES_SCHEMA = {
    "type": "object",
    "additionalProperties": False,
    "required": ["sections"],
    "properties": {
        "sections": {
            "type": "array",
            "items": {
                "type": "object",
                "additionalProperties": False,
                "required": ["heading", "bullets"],
                "properties": {
                    "heading": {"type": "string", "enum": ONBOARDING_HEADINGS},
                    "bullets": {"type": "array", "items": {"type": "string"}},
                },
            },
        },
    },
}

_NOTES_PROMPT = _CONTEXT + """

This is the client's onboarding meeting. Write the onboarding notes the team
will work from, under these headings only:
""" + "\n".join(f"- {h}" for h in ONBOARDING_HEADINGS) + """

- Bullets are short and factual: who they sell to (titles, company sizes,
  industries, regions), what they sell, targets and volumes, companies or
  contacts to exclude, tone and messaging dos and don'ts, who approves copy,
  dates and deadlines.
- Only what was said in the meeting. Leave a heading out if it wasn't covered.
- Keep names, numbers and spellings exactly as spoken.
"""


def notes_html(sections) -> str:
    parts = []
    by_heading = {}
    for s in sections or []:
        if not isinstance(s, dict):
            continue
        heading = s.get("heading")
        bullets = [str(b).strip() for b in s.get("bullets") or [] if str(b).strip()]
        if heading in ONBOARDING_HEADINGS and bullets:
            by_heading.setdefault(heading, []).extend(bullets)
    for heading in ONBOARDING_HEADINGS:
        if heading in by_heading:
            items = "".join(f"<li>{escape(b[:500])}</li>" for b in by_heading[heading][:15])
            parts.append(f"<h3>{escape(heading)}</h3><ul>{items}</ul>")
    return richtext.sanitize("".join(parts))


def draft_onboarding_notes(meeting_id: str) -> str:
    """Draft notes from the meeting and file them. Returns "filled", "pending" or "".

    Never overwrites a person's edits: if anyone has saved the notes, the
    draft goes to pending_draft for them to accept or dismiss.
    """
    meeting = ms.get_meeting(meeting_id)
    if not meeting or not meeting.get("client_id") or not llm.is_configured():
        return ""
    blocks = meeting.get("transcript") or []
    if not blocks and not meeting.get("read_summary"):
        return ""
    text, cut = transcript_for_prompt(blocks)
    actions = "\n".join(f"- {a}" for a in meeting.get("action_items") or [])
    user = (f"Meeting: {meeting.get('title')}\n"
            f"Read AI's summary: {meeting.get('read_summary') or '(none)'}\n"
            f"Action items:\n{actions or '(none)'}\n\n"
            f"Transcript{' (cut short)' if cut else ''}:\n{text or '(none)'}")
    try:
        raw = llm.complete_json(_NOTES_PROMPT, user, schema=NOTES_SCHEMA,
                                schema_name="onboarding_notes", max_tokens=2500)
    except llm.LLMError as exc:
        logger.warning("Onboarding notes draft failed for meeting %s: %s", meeting_id, exc)
        return ""
    html = notes_html(raw.get("sections"))
    if not richtext.to_text(html).strip():
        return ""
    return file_draft(str(meeting["client_id"]), meeting_id, html)


def file_draft(client_id: str, meeting_id: str, html: str) -> str:
    """Write a draft into the notes, or park it if a person has edited them."""
    for _ in range(2):
        previous = ms.get_notes(client_id)
        try:
            if previous and previous.get("edited_by"):
                ms.write_notes(client_id, previous, {
                    "pending_draft": html, "pending_meeting_id": meeting_id})
                return "pending"
            ms.write_notes(client_id, previous, {
                "body": html, "source_meeting_id": meeting_id,
                "ai_generated_at": ms.now_iso(),
                "pending_draft": "", "pending_meeting_id": None})
            return "filled"
        except ms.Conflict:
            continue  # someone saved in between: read again and re-decide
    return ""


# ── Orchestration ─────────────────────────────────────────────────────────────

_locks: dict[str, threading.Lock] = {}
_locks_guard = threading.Lock()


def _client_lock(client_id: str) -> threading.Lock:
    with _locks_guard:
        return _locks.setdefault(client_id, threading.Lock())


# Clients whose notes are being drafted right now, so the panel can say so
# and poll instead of showing an empty state. In-process only: the service
# runs one instance, and a lost entry after a restart only means the panel
# shows the empty state until the draft lands.
_drafting: set[str] = set()


def is_drafting(client_id: str) -> bool:
    return str(client_id) in _drafting


def _draft_tracked(client_id: str, meeting_id: str) -> str:
    _drafting.add(str(client_id))
    try:
        return draft_onboarding_notes(meeting_id)
    finally:
        _drafting.discard(str(client_id))


def process_meeting(meeting_id: str, *, draft_notes: bool = False) -> None:
    """Analyse a meeting (unless already done), draft notes if asked, refresh the overview."""
    try:
        meeting = ms.get_meeting(meeting_id, with_transcript=False) or {}
        if meeting.get("analysis_status") != "done":
            analyse_meeting(meeting_id)
        client_id = meeting.get("client_id")
        if not client_id:
            return
        if draft_notes and meeting.get("is_onboarding"):
            _draft_tracked(str(client_id), meeting_id)
        with _client_lock(str(client_id)):
            rebuild_overview(str(client_id))
    except Exception:
        logger.exception("Read AI processing failed for meeting %s", meeting_id)


def refresh_client(client_id: str) -> None:
    """Retry anything unfinished for a client, then rebuild its overview."""
    try:
        for m in ms.list_client_meetings(client_id):
            if m.get("analysis_status") in ("pending", "failed"):
                analyse_meeting(str(m["id"]))
        with _client_lock(client_id):
            rebuild_overview(client_id)
    except Exception:
        logger.exception("Read AI refresh failed for client %s", client_id)


def redraft_notes(client_id: str, meeting_id: str = "") -> None:
    try:
        if not meeting_id:
            meeting = ms.onboarding_meeting(client_id)
            meeting_id = str(meeting["id"]) if meeting else ""
        if meeting_id:
            _draft_tracked(client_id, meeting_id)
    except Exception:
        logger.exception("Onboarding notes redraft failed for client %s", client_id)
    finally:
        _drafting.discard(str(client_id))


# ── Arrival ───────────────────────────────────────────────────────────────────

_CONTENT_FIELDS = ("request_id", "title", "start_time", "end_time", "owner_email",
                   "participants", "report_url", "read_summary", "action_items",
                   "transcript")


def ingest(parsed: dict) -> dict:
    """File one parsed Read AI report. Returns what happened, for the webhook's reply.

    A repeat delivery (same request_id) is a no-op. A re-push of a known
    meeting refreshes its content but keeps whoever it was assigned to.
    """
    fields = {k: parsed.get(k) for k in _CONTENT_FIELDS}
    existing = ms.find_by_session(parsed["session_id"])
    if existing:
        if parsed.get("request_id") and existing.get("request_id") == parsed["request_id"]:
            return {"result": "duplicate"}
        changed = (existing.get("transcript") or []) != (parsed.get("transcript") or [])
        patch = dict(fields)
        if changed:
            patch.update({"analysis_status": "pending", "analysis_error": ""})
        ms.update_meeting(str(existing["id"]), patch)
        if changed and existing.get("client_id"):
            run_async(process_meeting, str(existing["id"]),
                      draft_notes=bool(existing.get("is_onboarding")))
        return {"result": "updated"}

    try:
        clients = ms.client_domains()
    except Exception as exc:
        # Most likely clients.domains has not been migrated. The meeting is
        # still worth keeping; it waits in the inbox for someone to assign it.
        logger.warning("Read AI: could not read client domains: %s", exc)
        clients = []
    status, candidates = readai.match_client(parsed.get("participants") or [], clients)
    if status == "unmatched" and not readai.external_domains(parsed.get("participants") or []):
        # Nobody from outside the agency: an internal call, not a client meeting.
        status = "ignored"
    client_id = candidates[0] if status == "matched" else None
    row = {**fields, "session_id": parsed["session_id"], "client_id": client_id,
           "match_status": status, "candidate_client_ids": candidates,
           "is_onboarding": bool(client_id) and should_mark_onboarding(parsed, client_id)}
    created = ms.insert_meeting(row)
    if not created:
        return {"result": "duplicate"}
    if client_id:
        run_async(process_meeting, str(created["id"]), draft_notes=row["is_onboarding"])
    return {"result": status}


def run_async(target, *args, **kwargs) -> None:
    threading.Thread(target=target, args=args, kwargs=kwargs, daemon=True,
                     name="pulse-readai").start()


def should_mark_onboarding(meeting: dict, client_id: str) -> bool:
    """A new meeting is the onboarding one if its title says so and the client has none."""
    return (readai.looks_like_onboarding(meeting.get("title", ""))
            and ms.onboarding_meeting(client_id) is None)

"""Is a mail-merge campaign still running?

The rule lived inside auto_sync, which is where it was first needed, and the
reporting tool then had to answer the same question — so it is here, once,
rather than in two places that could disagree about whether the campaign an
account manager is looking at is live.

IT CANNOT BE ASKED OF POSTGREST. "Finished" is partly `sent_count >=
total_prospects`, a comparison between two columns, which PostgREST has no
syntax for. So the rows come back unfiltered and are classified here. That is
fine: the list is scoped to one client and is tens of rows, not thousands.

NOR CAN THE COLUMNS BE RELIED ON. `completed` and `paused` are not in
`migrations/campaigns_schema.sql`; they exist on some deployments and not
others, which is why every caller walks a ladder of selects from richest to
poorest. With only `completed_at` to hand the answer degrades to "has it been
stamped finished", which is the right conservative reading — it errs towards
calling a campaign live, and a live campaign showing up where it should not is
a far smaller problem than the one this module exists to fix.
"""
from __future__ import annotations

# Richest first. A caller tries each in turn until PostgREST stops rejecting
# the select, exactly as auto_sync does.
SELECT_TIERS: tuple[str, ...] = (
    "total_prospects,sent_count,completed,paused,completed_at",
    "total_prospects,sent_count,completed_at",
    "completed_at",
)


def status_of(row: dict) -> str:
    """"paused" | "past" | "active", best effort from whatever columns are set."""
    if row.get("paused"):
        return "paused"
    total = row.get("total_prospects") or 0
    sent = row.get("sent_count") or 0
    if row.get("completed") or row.get("completed_at") or (total > 0 and sent >= total):
        return "past"
    return "active"


def is_finished(row: dict) -> bool:
    """Whether this campaign has stopped sending.

    Paused counts as live, not past. A paused campaign has not finished and
    its result is still the current one; filing it under history would hide
    the copy a client is about to resume sending.
    """
    return status_of(row) == "past"

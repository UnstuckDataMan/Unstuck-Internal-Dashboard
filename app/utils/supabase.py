"""
Shared Supabase REST helpers.

Every router used to define its own copy of these three helpers; the copies
drifted independently (which is how the sync-from-sheet endpoint missed the
on_conflict fix).  All Supabase-facing modules now import from here.
"""
from __future__ import annotations

import os

SUPABASE_URL      = os.environ.get("SUPABASE_URL", "").rstrip("/")
SUPABASE_ANON_KEY = os.environ.get("SUPABASE_ANON_KEY", "")


def sb_headers(prefer: str = "") -> dict:
    h = {
        "apikey":        SUPABASE_ANON_KEY,
        "Authorization": f"Bearer {SUPABASE_ANON_KEY}",
        "Content-Type":  "application/json",
    }
    if prefer:
        h["Prefer"] = prefer
    return h


def sb_configured() -> bool:
    return bool(SUPABASE_URL and SUPABASE_ANON_KEY)


def parse_total(headers) -> int:
    """Extract the total row count from a PostgREST Content-Range header."""
    cr = headers.get("Content-Range", "")
    if "/" in cr:
        try:
            return int(cr.split("/")[1])
        except ValueError:
            pass
    return 0


def pg_in_list(values: list[str]) -> str:
    """Build a PostgREST in.(...) literal with each value double-quoted.

    Unquoted values break on commas/parens: PostgREST URL-decodes the query
    string before parsing the in.() operator, so a value containing a literal
    comma (legal inside a quoted CSV cell) would be split into two separate
    list items.  Double-quoting with backslash escapes makes any punctuation
    filter correctly.
    """
    quoted = ",".join(
        '"' + str(v).replace("\\", "\\\\").replace('"', '\\"') + '"' for v in values
    )
    return f"({quoted})"

"""
Jinja filters for Outbound Pulse templates.

Lead counts and rates are derived from the stored counts rather than stored,
and every view — overview, client detail, client portal — needs them. Filters
keep that arithmetic in one place (normalize.py) instead of re-deriving it in
three templates, where the copies would drift.

Prefixed `pulse_` because the Jinja environment is shared with every other
tool in the dashboard.
"""
from __future__ import annotations

from app.utils.pulse.normalize import (
    lead_count,
    lead_rate,
    outcome_breakdown,
    reply_rate,
)


def _pct(value) -> str:
    """A rate for display: "1.2%", or an em dash when there is no base."""
    return "—" if value is None else f"{value}%"


def register(env) -> None:
    """Install the filters. Idempotent, so both routers can call it."""
    env.filters["pulse_leads"] = lead_count
    env.filters["pulse_lead_rate"] = lead_rate
    env.filters["pulse_reply_rate"] = reply_rate
    env.filters["pulse_outcomes"] = outcome_breakdown
    env.filters["pulse_pct"] = _pct

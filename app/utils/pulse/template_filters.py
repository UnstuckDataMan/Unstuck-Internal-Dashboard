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

import math

from app.utils.pulse.normalize import (
    headline_rates,
    interest_rate,
    lead_count,
    lead_rate,
    outcome_breakdown,
    reply_rate,
    unsubscribe_count,
)


def _pct(value) -> str:
    """A rate for display: "1.2%", "0.045%", or an em dash when there is no base.

    Formatted explicitly rather than by interpolating the float. A small enough
    value renders as "1e-05" through str(), which is not a percentage anybody
    wants to read on a client report.

    Trailing zeros are stripped below 0.1% so the number shows the precision it
    actually has: 0.04% rather than 0.040%.
    """
    if value is None:
        return "—"
    try:
        number = float(value)
    except (TypeError, ValueError):
        return "—"
    if number and abs(number) < 0.1:
        places = min(6, max(2, 1 - math.floor(math.log10(abs(number)))))
        return f"{number:.{places}f}".rstrip("0").rstrip(".") + "%"
    return f"{number:.1f}%"


def register(env) -> None:
    """Install the filters. Idempotent, so both routers can call it."""
    env.filters["pulse_leads"] = lead_count
    env.filters["pulse_lead_rate"] = lead_rate
    env.filters["pulse_reply_rate"] = reply_rate
    env.filters["pulse_outcomes"] = outcome_breakdown
    env.filters["pulse_rates"] = headline_rates
    env.filters["pulse_interest_rate"] = interest_rate
    env.filters["pulse_unsubs"] = unsubscribe_count
    env.filters["pulse_pct"] = _pct

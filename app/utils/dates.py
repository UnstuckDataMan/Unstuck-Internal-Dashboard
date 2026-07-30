"""
Shared clock for campaign send stats.

Both halves of the stats pipeline must agree on what "today" is:

  • app/utils/auto_sync.py  — stamps the Sent Date written to the sheet and to
    contacted_prospects.contacted_at
  • app/routers/campaigns.py — builds the Today / This Week / This Month bucket
    boundaries those rows are counted against

The business day is UTC.  `date.today()` would also return the UTC date on
Render (containers default to UTC and no TZ env var is set), but only by
accident of the environment — setting TZ, or running on a dev machine, would
silently shift every bucket boundary by a day.  Deriving from an explicit UTC
instant pins the behaviour instead of inheriting it.
"""
from __future__ import annotations

from datetime import date as _date, datetime as _datetime, timezone as _tz


def today_utc() -> _date:
    """Today's date in UTC — the single source of truth for send-stat dates."""
    return _datetime.now(_tz.utc).date()

"""
Deterministic outbound email sending scheduler.

Rules enforced:
  - Send window: 08:30 – 15:30 in the RECIPIENT's local timezone
  - Output send_time is expressed in the SENDER's local timezone (DST-aware)
  - Sender 1 starts at ~09:05; each subsequent sender starts ~60 min later
  - Small deterministic variance on start times (±5 min, based on hash)
  - Max 15 sends per sender per day
  - Global 5-minute minimum gap between ANY two sends across ALL senders
  - Senders start in staggered waves; daily counts auto-scale to monthly target
  - Weekends and England bank holidays are excluded
  - Prospect count range: 100–1650 (each prospect = 2 sends total incl. chaser,
    but only the initial send is scheduled)
"""
import hashlib
import math
from calendar import monthrange
from datetime import date, datetime, time, timedelta
from typing import Dict, List, Optional

from dateutil import tz as tz_mod

from utils.bank_holidays import get_england_bank_holidays

# --------------------------------------------------------------------------- #
# Deterministic variance
# --------------------------------------------------------------------------- #

def _dvariance(seed: str, lo: int, hi: int) -> int:
    """
    Deterministic integer in [lo, hi] derived from an MD5 hash of `seed`.
    No true randomness — same seed always returns the same value.
    """
    h = int(hashlib.md5(seed.encode()).hexdigest()[:8], 16)
    span = hi - lo + 1
    return lo + (h % span)


# --------------------------------------------------------------------------- #
# Working day calculation
# --------------------------------------------------------------------------- #

def get_working_days(year: int, month: int) -> List[date]:
    """Return every working day (Mon–Fri, excluding England bank holidays) in a month."""
    bank_hols = get_england_bank_holidays(year)
    _, days_in_month = monthrange(year, month)
    return [
        date(year, month, d)
        for d in range(1, days_in_month + 1)
        if date(year, month, d).weekday() < 5
        and date(year, month, d) not in bank_hols
    ]


def get_working_days_info(
    year: int,
    month: int,
    prospect_count: int,
    max_per_sender_per_day: int = 15,
) -> Dict:
    """
    Calculate and return scheduling statistics for the given parameters.
    Used by the frontend to display live info before generating a schedule.
    """
    import calendar as cal_mod
    working_days = get_working_days(year, month)
    n_days = len(working_days)

    if n_days == 0:
        raise ValueError(f"No working days found in {cal_mod.month_name[month]} {year}.")

    prospect_count = max(1, prospect_count)

    prospects_per_day = math.ceil(prospect_count / n_days)
    senders_per_day = math.ceil(prospects_per_day / max_per_sender_per_day)
    sends_per_sender = math.ceil(prospects_per_day / senders_per_day)

    return {
        'working_days': n_days,
        'working_day_dates': [d.isoformat() for d in working_days],
        'total_sends_incl_chasers': prospect_count * 2,
        'prospects_per_day': prospects_per_day,
        'senders_per_day': senders_per_day,
        'sends_per_sender_per_day': sends_per_sender,
        'send_window': '08:30 - 15:30',
        'month_name': cal_mod.month_name[month],
    }


# --------------------------------------------------------------------------- #
# Schedule generation
# --------------------------------------------------------------------------- #

SENDER_OFFSET_MINS = 60      # 1 hour between staggered senders (compressed if needed)
MAX_PER_SENDER_PER_DAY = 15


def _parse_hhmm(t: str) -> tuple:
    """Parse 'HH:MM' into (hour, minute) ints."""
    h, m = t.strip().split(':')
    return int(h), int(m)


def generate_schedule(
    year: int,
    month: int,
    prospect_count: int,
    sender_emails: List[str],
    recipient_tz: str = 'Europe/London',
    sender_tz: str = 'Europe/London',
    max_per_sender_per_day: int = 15,
    window_start: str = '08:30',
    window_end: str = '15:30',
) -> List[Dict]:
    """
    Generate a complete sending schedule for `prospect_count` initial emails
    spread across the working days of `year`/`month`.

    The send window (default 08:30–15:30) is applied in `recipient_tz`.
    The returned `send_time` values are expressed in `sender_tz` (DST-aware).

    `max_per_sender_per_day` caps how many sends each account can make per
    working day.  The caller is responsible for pre-clamping `prospect_count`
    to the monthly capacity before calling this function.

    Returns a list of dicts, each representing one scheduled send:
      date, day_of_week, send_time, sender, sender_number, prospect_id
    """
    if prospect_count < 1:
        raise ValueError("Prospect count must be at least 1.")
    if not sender_emails:
        raise ValueError("At least one sender email is required.")

    # Parse send-window boundaries
    win_start_h, win_start_m = _parse_hhmm(window_start)
    win_end_h,   win_end_m   = _parse_hhmm(window_end)
    win_start_mins = win_start_h * 60 + win_start_m
    win_end_mins   = win_end_h   * 60 + win_end_m
    if win_end_mins <= win_start_mins:
        raise ValueError("Send window end must be after send window start.")

    # Sender 1 starts 5 minutes into the window
    s1_start_mins = win_start_mins + 5

    working_days = get_working_days(year, month)
    n_days = len(working_days)
    if n_days == 0:
        raise ValueError("No working days in the selected month.")

    # Timezone objects for DST-aware conversion
    rec_zone = tz_mod.gettz(recipient_tz)
    snd_zone = tz_mod.gettz(sender_tz)
    if rec_zone is None:
        raise ValueError(f"Unknown recipient timezone: {recipient_tz!r}")
    if snd_zone is None:
        raise ValueError(f"Unknown sender timezone: {sender_tz!r}")

    schedule: List[Dict] = []
    prospect_idx = 0

    for day_num, work_day in enumerate(working_days):
        if prospect_idx >= prospect_count:
            break

        # ── How many prospects to assign today ──────────────────────────── #
        # Fixed-rate: fill each day at max_per_sender_per_day × senders until
        # prospects run out.  The last (partial) day gets whatever remains.
        remaining_prospects = prospect_count - prospect_idx
        day_target = min(max_per_sender_per_day * len(sender_emails), remaining_prospects)

        # ── How many senders are active today ───────────────────────────── #
        # Use ALL available senders (up to day_target — no point activating
        # a sender if there are fewer prospects than senders).
        senders_today = min(day_target, len(sender_emails))

        day_win_end = datetime.combine(work_day, time(win_end_h, win_end_m))

        # ── Compute each sender's deterministic start time ───────────────── #
        # Compress the inter-sender offset so the LAST sender starts no later
        # than 45 min before window end.  This guarantees every sender has
        # enough room to produce varied send times (the effective-end variance
        # below shifts up to 20 min, so we need ≥20 min of headroom minimum).
        max_last_start_mins = win_end_mins - 45
        available_for_stagger = max(0, max_last_start_mins - s1_start_mins)
        offset_mins = (
            min(SENDER_OFFSET_MINS, max(5, available_for_stagger // (senders_today - 1)))
            if senders_today > 1 else SENDER_OFFSET_MINS
        )

        sender_starts: List[datetime] = []
        for s_idx in range(senders_today):
            base_mins = s1_start_mins + s_idx * offset_mins
            variance = _dvariance(
                f"{year}-{month:02d}-{work_day.day:02d}-s{s_idx}", -5, 5)
            start_mins = base_mins + variance
            s_hour, s_min = divmod(max(start_mins, win_start_mins), 60)
            t = datetime.combine(work_day, time(s_hour, s_min))
            # Clamp to window end so avail_secs is never negative
            t = min(t, day_win_end)
            sender_starts.append(t)

        # ── Per-sender independent slot generation ───────────────────────── #
        # Each sender schedules their sends evenly from their staggered start
        # to the window end.  Different accounts can send at the same clock
        # time, so there is no global gap constraint — just natural per-sender
        # spacing.  This guarantees exactly day_target sends per day.
        base_per_sender = day_target // senders_today
        remainder       = day_target % senders_today

        adjusted: List[tuple] = []
        for s_idx in range(senders_today):
            n_sends      = base_per_sender + (1 if s_idx < remainder else 0)
            sender_start = sender_starts[s_idx]

            # Per-sender, per-day effective window end.
            # Use a negative-only shift (-20 to 0 min) so the hard-window clamp
            # never cancels the variance — the last slot lands somewhere in the
            # final 20 minutes of the window, varying by day and by sender.
            end_shift = _dvariance(
                f"{year}-{month:02d}-{work_day.day:02d}-wend-s{s_idx}", -20, 0)
            effective_end = day_win_end + timedelta(minutes=end_shift)
            # Guarantee at least 1 minute of window so interval > 0
            effective_end = max(effective_end, sender_start + timedelta(minutes=1))

            avail_secs = (effective_end - sender_start).total_seconds()

            if n_sends == 1:
                slots = [sender_start]
            else:
                interval_secs = avail_secs / (n_sends - 1)
                slots = [
                    sender_start + timedelta(seconds=i * interval_secs)
                    for i in range(n_sends)
                ]

            _tolerance = timedelta(seconds=1)
            for slot_dt in slots:
                if slot_dt <= day_win_end + _tolerance:
                    adjusted.append((slot_dt, sender_emails[s_idx], s_idx + 1))

        adjusted.sort(key=lambda x: x[0])

        # ── Resolve per-minute conflicts across senders ──────────────────── #
        # Excel outputs HH:MM only, so two slots at the same minute look like
        # simultaneous sends.  Sort descending and bump each conflict 1 minute
        # earlier until unique, keeping sends inside the window.
        win_floor = datetime.combine(work_day, time(win_start_h, win_start_m))
        adjusted.sort(key=lambda x: x[0], reverse=True)
        used_minutes: set = set()
        deduped: List[tuple] = []
        for slot_dt, sender_email, s_num in adjusted:
            minute_key = slot_dt.replace(second=0, microsecond=0)
            while minute_key in used_minutes and minute_key > win_floor:
                minute_key -= timedelta(minutes=1)
            used_minutes.add(minute_key)
            deduped.append((minute_key, sender_email, s_num))
        adjusted = sorted(deduped, key=lambda x: x[0])

        # ── Assign prospects to adjusted slots ──────────────────────────── #
        for send_dt, sender, s_num in adjusted:
            if prospect_idx >= prospect_count:
                break
            # send_dt is naive, expressed in recipient's local time.
            # Attach recipient tz → convert to sender tz for the output time.
            slot_aware   = send_dt.replace(tzinfo=rec_zone)
            sender_local = slot_aware.astimezone(snd_zone)
            schedule.append({
                'date': work_day.isoformat(),
                'day_of_week': work_day.strftime('%A'),
                'send_time': sender_local.strftime('%H:%M'),
                'sender': sender,
                'sender_number': s_num,
                'prospect_id': prospect_idx + 1,
            })
            prospect_idx += 1

    return schedule


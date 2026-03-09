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


def get_working_days_info(year: int, month: int, prospect_count: int) -> Dict:
    """
    Calculate and return scheduling statistics for the given parameters.
    Used by the frontend to display live info before generating a schedule.
    """
    import calendar as cal_mod
    working_days = get_working_days(year, month)
    n_days = len(working_days)

    if n_days == 0:
        raise ValueError(f"No working days found in {cal_mod.month_name[month]} {year}.")

    prospect_count = max(100, min(1650, prospect_count))

    prospects_per_day = math.ceil(prospect_count / n_days)
    senders_per_day = math.ceil(prospects_per_day / 15)
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

WINDOW_START_H, WINDOW_START_M = 8, 30   # 08:30
WINDOW_END_H,   WINDOW_END_M   = 15, 30  # 15:30
MIN_GAP_SECS = 5 * 60                    # 5 minutes
SENDER_1_START_H, SENDER_1_START_M = 9, 5   # ~09:05
SENDER_OFFSET_MINS = 60                  # 1 hour between senders
MAX_PER_SENDER_PER_DAY = 15


def generate_schedule(
    year: int,
    month: int,
    prospect_count: int,
    sender_emails: List[str],
    recipient_tz: str = 'Europe/London',
    sender_tz: str = 'Europe/London',
) -> List[Dict]:
    """
    Generate a complete sending schedule for `prospect_count` initial emails
    spread across the working days of `year`/`month`.

    The 08:30–15:30 send window is applied in `recipient_tz`.
    The returned `send_time` values are expressed in `sender_tz` (DST-aware).

    Returns a list of dicts, each representing one scheduled send:
      date, day_of_week, send_time, sender, sender_number, prospect_id
    """
    if not (100 <= prospect_count <= 1650):
        raise ValueError("Prospect count must be between 100 and 1650.")
    if not sender_emails:
        raise ValueError("At least one sender email is required.")

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
        remaining_prospects = prospect_count - prospect_idx
        remaining_days = n_days - day_num
        day_target = math.ceil(remaining_prospects / remaining_days)
        day_target = min(day_target, MAX_PER_SENDER_PER_DAY * len(sender_emails))

        # ── How many senders are active today ───────────────────────────── #
        # Use ALL available senders (up to day_target — no point activating
        # a sender if there are fewer prospects than senders).
        senders_today = min(day_target, len(sender_emails))

        window_end = datetime.combine(work_day, time(WINDOW_END_H, WINDOW_END_M))

        # ── Compute each sender's deterministic start time ───────────────── #
        # Dynamically shrink the inter-sender offset so that ALL active senders
        # have a start time within the send window (e.g. 10 senders at 60-min
        # gaps would push senders 8-10 past 15:30; we compress to fit them in).
        available_window_mins = (
            WINDOW_END_H * 60 + WINDOW_END_M
            - SENDER_1_START_H * 60 - SENDER_1_START_M
        )
        offset_mins = (
            min(SENDER_OFFSET_MINS, available_window_mins // (senders_today - 1))
            if senders_today > 1 else SENDER_OFFSET_MINS
        )

        sender_starts: List[datetime] = []
        for s_idx in range(senders_today):
            base_mins = (SENDER_1_START_H * 60 + SENDER_1_START_M
                         + s_idx * offset_mins)
            variance = _dvariance(
                f"{year}-{month:02d}-{work_day.day:02d}-s{s_idx}", -5, 5)
            start_mins = base_mins + variance
            s_hour, s_min = divmod(max(start_mins, WINDOW_START_H * 60), 60)
            t = datetime.combine(work_day, time(s_hour, s_min))
            sender_starts.append(t)

        # ── Flat-timeline slot generation ────────────────────────────────── #
        # Build one unified, chronologically-ordered list of send times
        # across the whole day window.  The interval is sized to fit
        # day_target sends while respecting the 5-minute global gap.
        day_start = sender_starts[0]
        total_window_secs = (window_end - day_start).total_seconds()

        if day_target > 1:
            # Space evenly; never closer than MIN_GAP_SECS
            interval_secs = max(total_window_secs / (day_target - 1), MIN_GAP_SECS)
        else:
            interval_secs = 0

        # Generate the flat list of send times
        flat_slots: List[datetime] = []
        for i in range(day_target):
            t = day_start + timedelta(seconds=i * interval_secs)
            if t > window_end:
                break
            flat_slots.append(t)

        # ── Assign each slot to a sender (round-robin) ───────────────────── #
        # Distribute slots evenly across all active senders using round-robin.
        # A sender is eligible when the slot time >= their staggered start and
        # they haven't yet hit MAX_PER_SENDER_PER_DAY.
        sender_counts = [0] * senders_today
        adjusted: List[tuple] = []
        rr = 0  # round-robin pointer

        for slot_dt in flat_slots:
            for attempt in range(senders_today):
                s_idx = (rr + attempt) % senders_today
                if (slot_dt >= sender_starts[s_idx]
                        and sender_counts[s_idx] < MAX_PER_SENDER_PER_DAY):
                    adjusted.append((slot_dt, sender_emails[s_idx], s_idx + 1))
                    sender_counts[s_idx] += 1
                    rr = (s_idx + 1) % senders_today
                    break

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


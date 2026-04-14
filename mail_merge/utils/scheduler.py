"""
Deterministic outbound email sending scheduler.

Rules enforced:
  - Send window: 08:30 – 15:30 in the RECIPIENT's local timezone
  - Output send_time is expressed in the SENDER's local timezone (DST-aware)
  - Sender 1 starts at ~09:05; each subsequent sender starts ~60 min later
  - Small deterministic variance on start times (±5 min, based on hash)
  - Weekends excluded; no bank holiday exclusion
  - Sends fill consecutive working days from start_date until all prospects
    are scheduled — no monthly cap
  - Per-day total is always exactly nominal × num_senders; variance
    redistributes who sends more/fewer each day within the allowed band
"""
import hashlib
from datetime import date, datetime, time, timedelta
from typing import Dict, List, Optional

from dateutil import tz as tz_mod


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
# Working day helpers
# --------------------------------------------------------------------------- #

def _next_working_day(d: date) -> date:
    """Advance d until it lands on a weekday (Mon–Fri)."""
    while d.weekday() >= 5:
        d += timedelta(days=1)
    return d


# --------------------------------------------------------------------------- #
# Schedule generation
# --------------------------------------------------------------------------- #

SENDER_OFFSET_MINS = 60      # 1 hour between staggered senders (compressed if needed)
MAX_PER_SENDER_PER_DAY = 15

# Variance band for each nominal sends-per-day setting.
# Each sender's daily count = nominal ± variance, but the SUM across all
# senders always equals nominal × num_senders (variance is redistributed).
#   5/day  → 3–7   (± 2)
#   10/day → 8–12  (± 2)
#   15/day → 13–17 (± 2)
SENDS_VARIANCE: dict = {5: 2, 10: 2, 15: 2}


def _parse_hhmm(t: str) -> tuple:
    """Parse 'HH:MM' into (hour, minute) ints."""
    h, m = t.strip().split(':')
    return int(h), int(m)


def generate_schedule(
    prospect_count: int,
    sender_emails: List[str],
    start_date: Optional[date] = None,
    campaign_seed: str = '',
    recipient_tz: str = 'Europe/London',
    sender_tz: str = 'Europe/London',
    max_per_sender_per_day: int = 15,
    window_start: str = '08:30',
    window_end: str = '15:30',
) -> List[Dict]:
    """
    Generate a complete sending schedule for `prospect_count` initial emails,
    filling consecutive working weekdays (Mon–Fri) from `start_date` until
    all prospects are scheduled.  No monthly cap is applied.

    `start_date` defaults to the next working weekday from today.

    The send window (default 08:30–15:30) is applied in `recipient_tz`.
    The returned `send_time` values are expressed in `sender_tz` (DST-aware).

    On each full working day the total sends equals exactly
    `max_per_sender_per_day × len(sender_emails)`; variance is redistributed
    across sender accounts so individual counts differ while the total is fixed.

    Returns a list of dicts, each representing one scheduled send:
      date, day_of_week, send_time, sender, sender_number, prospect_id
    """
    if prospect_count < 1:
        raise ValueError("Prospect count must be at least 1.")
    if not sender_emails:
        raise ValueError("At least one sender email is required.")

    # Resolve start date
    if start_date is None:
        start_date = _next_working_day(date.today())

    # Parse send-window boundaries
    win_start_h, win_start_m = _parse_hhmm(window_start)
    win_end_h,   win_end_m   = _parse_hhmm(window_end)
    win_start_mins = win_start_h * 60 + win_start_m
    win_end_mins   = win_end_h   * 60 + win_end_m
    if win_end_mins <= win_start_mins:
        raise ValueError("Send window end must be after send window start.")
    if win_end_mins - win_start_mins < 420:
        raise ValueError(
            f"Send window must be at least 7 hours. "
            f"Current window is {(win_end_mins - win_start_mins) // 60}h "
            f"{(win_end_mins - win_start_mins) % 60}m."
        )

    # Sender 1 starts 5 minutes into the window
    s1_start_mins = win_start_mins + 5

    # Timezone objects for DST-aware conversion
    rec_zone = tz_mod.gettz(recipient_tz)
    snd_zone = tz_mod.gettz(sender_tz)
    if rec_zone is None:
        raise ValueError(f"Unknown recipient timezone: {recipient_tz!r}")
    if snd_zone is None:
        raise ValueError(f"Unknown sender timezone: {sender_tz!r}")

    schedule: List[Dict] = []
    prospect_idx = 0
    work_day = start_date

    while prospect_idx < prospect_count:

        # Skip weekends
        if work_day.weekday() >= 5:
            work_day += timedelta(days=1)
            continue

        # ── How many prospects to assign today ──────────────────────────── #
        remaining_prospects = prospect_count - prospect_idx
        full_day_target     = max_per_sender_per_day * len(sender_emails)
        day_target          = min(full_day_target, remaining_prospects)
        senders_today       = min(day_target, len(sender_emails))

        v = SENDS_VARIANCE.get(max_per_sender_per_day, 0)

        if day_target == full_day_target and v > 0 and senders_today > 1:
            # Full day: generate per-sender offsets and balance them to sum=0
            # so the total stays exactly full_day_target.
            offsets = [
                _dvariance(
                    f"{campaign_seed}-{work_day.isoformat()}-limit-s{s_idx}", -v, v
                )
                for s_idx in range(senders_today)
            ]
            surplus = sum(offsets)
            if surplus != 0:
                order = sorted(
                    range(senders_today),
                    key=lambda i: -offsets[i] if surplus > 0 else offsets[i],
                )
                for i in order:
                    if surplus == 0:
                        break
                    if surplus > 0:
                        cut = min(offsets[i] - (-v), surplus)
                        offsets[i] -= cut
                        surplus    -= cut
                    else:
                        add = min(v - offsets[i], -surplus)
                        offsets[i] += add
                        surplus    += add
            per_sender_counts = [max_per_sender_per_day + off for off in offsets]
        else:
            # Partial last day: distribute evenly across active senders
            base = day_target // senders_today
            rem  = day_target % senders_today
            per_sender_counts = [
                base + (1 if s < rem else 0) for s in range(senders_today)
            ] + [0] * (len(sender_emails) - senders_today)

        day_win_end = datetime.combine(work_day, time(win_end_h, win_end_m))

        # ── Compute each sender's deterministic start time ───────────────── #
        max_last_start_mins   = win_end_mins - 45
        available_for_stagger = max(0, max_last_start_mins - s1_start_mins)
        offset_mins = (
            min(SENDER_OFFSET_MINS, max(5, available_for_stagger // (senders_today - 1)))
            if senders_today > 1 else SENDER_OFFSET_MINS
        )

        sender_starts: List[datetime] = []
        for s_idx in range(senders_today):
            base_mins = s1_start_mins + s_idx * offset_mins
            variance  = _dvariance(
                f"{campaign_seed}-{work_day.isoformat()}-s{s_idx}", -5, 5)
            start_mins = base_mins + variance
            s_hour, s_min = divmod(max(start_mins, win_start_mins), 60)
            t = datetime.combine(work_day, time(s_hour, s_min))
            t = min(t, day_win_end)
            sender_starts.append(t)

        # ── Per-sender independent slot generation ───────────────────────── #
        adjusted: List[tuple] = []
        for s_idx in range(senders_today):
            n_sends = per_sender_counts[s_idx]
            if n_sends == 0:
                continue
            sender_start = sender_starts[s_idx]

            end_shift = _dvariance(
                f"{campaign_seed}-{work_day.isoformat()}-wend-s{s_idx}", -20, 0)
            effective_end = day_win_end + timedelta(minutes=end_shift)
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
            slot_aware   = send_dt.replace(tzinfo=rec_zone)
            sender_local = slot_aware.astimezone(snd_zone)
            schedule.append({
                'date':        work_day.isoformat(),
                'day_of_week': work_day.strftime('%A'),
                'send_time':   sender_local.strftime('%H:%M'),
                'sender':      sender,
                'sender_number': s_num,
                'prospect_id': prospect_idx + 1,
            })
            prospect_idx += 1

        work_day += timedelta(days=1)

    return schedule

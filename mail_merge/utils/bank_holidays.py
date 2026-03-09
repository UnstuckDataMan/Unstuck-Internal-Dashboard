"""
England & Wales bank holiday calculator.
Computes holidays algorithmically — no external API or library required.
"""
from datetime import date, timedelta
import calendar


def _easter(year: int) -> date:
    """Anonymous Gregorian algorithm for Easter Sunday."""
    a = year % 19
    b = year // 100
    c = year % 100
    d = b // 4
    e = b % 4
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - d - g + 15) % 30
    i = c // 4
    k = c % 4
    l = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * l) // 451
    month = (h + l - 7 * m + 114) // 31
    day = ((h + l - 7 * m + 114) % 31) + 1
    return date(year, month, day)


def _first_monday(year: int, month: int) -> date:
    """Return the first Monday of a given month."""
    d = date(year, month, 1)
    while d.weekday() != 0:  # 0 = Monday
        d += timedelta(days=1)
    return d


def _last_monday(year: int, month: int) -> date:
    """Return the last Monday of a given month."""
    last_day = date(year, month, calendar.monthrange(year, month)[1])
    while last_day.weekday() != 0:
        last_day -= timedelta(days=1)
    return last_day


def get_england_bank_holidays(year: int) -> set:
    """
    Return a set of date objects representing England & Wales bank holidays
    for the given calendar year.
    """
    holidays = set()

    # --- New Year's Day ---
    ny = date(year, 1, 1)
    wd = ny.weekday()
    if wd == 5:      # Saturday → Monday 3rd
        holidays.add(date(year, 1, 3))
    elif wd == 6:    # Sunday → Monday 2nd
        holidays.add(date(year, 1, 2))
    else:
        holidays.add(ny)

    # --- Easter ---
    easter_sunday = _easter(year)
    holidays.add(easter_sunday - timedelta(days=2))   # Good Friday
    holidays.add(easter_sunday + timedelta(days=1))   # Easter Monday

    # --- Early May Bank Holiday (first Monday in May) ---
    holidays.add(_first_monday(year, 5))

    # --- Spring Bank Holiday (last Monday in May) ---
    holidays.add(_last_monday(year, 5))

    # --- Summer Bank Holiday (last Monday in August) ---
    holidays.add(_last_monday(year, 8))

    # --- Christmas & Boxing Day ---
    christmas = date(year, 12, 25)
    xmas_wd = christmas.weekday()
    if xmas_wd == 4:    # Friday: Christmas on 25th, Boxing Day sub → Mon 28th
        holidays.add(date(year, 12, 25))
        holidays.add(date(year, 12, 28))
    elif xmas_wd == 5:  # Saturday: Christmas sub → Mon 27th, Boxing → Tue 28th
        holidays.add(date(year, 12, 27))
        holidays.add(date(year, 12, 28))
    elif xmas_wd == 6:  # Sunday: Christmas sub → Mon 26th, Boxing → Tue 27th
        holidays.add(date(year, 12, 26))
        holidays.add(date(year, 12, 27))
    else:               # Mon–Thu: both on their actual dates
        holidays.add(date(year, 12, 25))
        holidays.add(date(year, 12, 26))

    return holidays


def is_working_day(d: date) -> bool:
    """Return True if `d` is a weekday that is not an England bank holiday."""
    if d.weekday() >= 5:
        return False
    holidays = get_england_bank_holidays(d.year)
    return d not in holidays

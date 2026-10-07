"""The per-client Performance Tracker sheet: how good were the leads?

Every client has a Google Sheet an account manager maintains by hand, grading
each lead against that client's ICP. This module turns its "All leads" tab into
figures. It does no I/O — it takes the rows a reader handed back — so every rule
in here is testable against a list of dicts.

WHY NOT THE "PERFORMANCE TRACKER" TAB. That tab already holds a total, an
average and three charts, and it is tempting to read them instead of deriving
our own. Do not. It is a hand-built layout of scattered cells and chart ranges
that breaks the first time somebody drags a column, and it is all-time where a
report needs a period. Everything on it is derivable from "All leads", which is
a plain table. We read its *conventions* from it — the canonical band and region
lists below are its own — and nothing else.

"RAW LEADS" IS NOT OURS. It is locked, has a different schema, and must never
be read or written by this code.

WHAT COUNTS AS A ROW. The real sheet is ~2,500 rows long because a formula is
filled down it; only a few dozen are leads. The test is that the date parses.
Not "the Lead cell is filled" — a real lead in the live sheet has a blank Lead
cell, and filtering on it silently drops them and disagrees with the client's
own total. Not "any cell is non-empty" — that admits all 2,500.

UNGRADED IS NOT ZERO. A third of the rows carry "-" in the ICP column. The
sheet's own average excludes them and so does ours, which is the only way the
two agree. An ungraded lead is carried as None everywhere below and is never
coerced, imputed or defaulted — a 0 would drag an average down by a third and
nobody would be able to see why.
"""
from __future__ import annotations

import logging
import re
from datetime import date, datetime, timedelta

logger = logging.getLogger(__name__)

# The data tab. Named here rather than in the caller so there is one spelling.
DEFAULT_TAB = "All leads"

# Never read this one, whatever a caller passes.
FORBIDDEN_TABS = frozenset({"raw leads"})


# ── Column resolution ────────────────────────────────────────────────────────
#
# Exact (normalised) header names, not substrings. Substring matching is how
# "Sale stage" ends up being read as a date column.
_ALIASES: dict[str, tuple[str, ...]] = {
    "lead":      ("lead", "company", "domain", "website"),
    "channel":   ("channel", "source"),
    "headcount": ("headcount", "company size", "employees", "size"),
    "role":      ("job role", "job title", "role", "title"),
    "industry":  ("industry", "lead industry"),
    "location":  ("location", "country", "region"),
    "icp":       ("icp rating", "icp score", "icp"),
    "category":  ("category", "lead status", "status"),
    "stage":     ("sale stage", "sales stage", "stage"),
    "notes":     ("additional notes", "notes"),
}

# The date column is named after the client — "Dovetail Date/Time" — so it is
# the one column matched by shape. Deliberately anchored at the end: "Date
# added" is a different thing and we would rather report no date column than
# quietly scope a client's report by the wrong one.
_DATE_HEADER = re.compile(r"(^|\s|/)date(\s*/\s*time)?$")


def _norm(header) -> str:
    return " ".join(str(header or "").strip().lower().split()).rstrip(":")


# How far down to look for the header row before giving up.
MAX_HEADER_SCAN = 10

_KNOWN_HEADERS = frozenset(a for names in _ALIASES.values() for a in names)


def header_score(row: list[str]) -> int:
    """How much this row looks like a header: the column names it carries."""
    score = 0
    for cell in row:
        name = _norm(cell)
        if not name:
            continue
        if name in _KNOWN_HEADERS:
            score += 1
        elif _DATE_HEADER.search(name):
            score += 1
    return score


def find_header(values: list[list[str]]) -> tuple[int, list[str]]:
    """Which row holds the column names. Returns (index, headers).

    ROW 1 IS NOT THE HEADER on a real tracker. The live sheets open with a
    banner naming the client — "Dovetail" alone in A1, the rest of the row
    blank — and put the column names on row 2. Taking row 1 regardless gave a
    sheet with one usable column and no ICP rating at all, and said so in a
    warning rather than failing, which is how it reached production looking
    linked but empty.

    Scored rather than hardcoded to "row 2", because the banner is a
    convention and not a rule: the row that names the most columns we
    recognise is the header, whichever row that is. On the live sheet row 2
    scores 11 and nothing else scores above 1, so there is no close call to
    get wrong.
    """
    best, best_score = 0, -1
    for index, row in enumerate(values[:MAX_HEADER_SCAN]):
        score = header_score(row)
        if score > best_score:
            best, best_score = index, score
    # Nothing recognisable anywhere near the top: fall back to the first row
    # so resolve_columns can report what it could not find, by name.
    if best_score <= 0:
        return 0, list(values[0]) if values else []
    return best, list(values[best])


def rows_to_dicts(values: list[list[str]]) -> tuple[list[dict], list[str]]:
    """The grid below the header row, keyed by it. Blank and repeated column
    names are dropped, keeping the first — two columns called the same thing
    cannot both be read, and the first is the one a person means."""
    if not values:
        return [], []
    index, headers = find_header(values)
    seen: set[str] = set()
    keys: list[str | None] = []
    for cell in headers:
        name = str(cell or "").strip()
        if name and name not in seen:
            seen.add(name)
            keys.append(name)
        else:
            keys.append(None)
    out = [{k: v for k, v in zip(keys, row) if k is not None}
           for row in values[index + 1:]]
    return out, [k for k in keys if k is not None]


def resolve_columns(headers: list[str]) -> tuple[dict[str, str], list[str]]:
    """Map our field names onto this sheet's headers. Returns (columns, warnings).

    Every field is optional. A missing one disables its breakdown and is named
    in the warnings, rather than being guessed at by position — a breakdown
    built from the wrong column is worse than an absent one.
    """
    warnings: list[str] = []
    normalised = [(h, _norm(h)) for h in headers]
    taken: set[str] = set()
    columns: dict[str, str] = {}

    for field, names in _ALIASES.items():
        for want in names:
            hit = next((h for h, n in normalised if n == want and h not in taken), None)
            if hit is not None:
                columns[field] = hit
                taken.add(hit)
                break

    # The date column, by shape, among the headers nothing else claimed.
    candidates = [h for h, n in normalised
                  if h not in taken and _DATE_HEADER.search(n)]
    if candidates:
        columns["date"] = candidates[0]
        if len(candidates) > 1:
            warnings.append(
                "More than one column looks like the date — using "
                f"“{candidates[0]}”. The others: "
                + ", ".join(f"“{c}”" for c in candidates[1:]) + ".")
    elif headers and headers[0] not in taken:
        # No header matched, so fall back to the first column — but only on
        # evidence, otherwise a sheet whose first column is "Notes" becomes the
        # thing every report period is scoped by.
        columns["date"] = headers[0]
        warnings.append(
            f"No column is named like a date, so “{headers[0]}” is being "
            "used as one.")

    for field in ("date", "icp"):
        if field not in columns:
            warnings.append(
                f"No {'date' if field == 'date' else 'ICP rating'} column found."
                + (" Leads cannot be matched to a report period."
                   if field == "date" else " Nothing can be graded."))
    return columns, warnings


# ── Dates ────────────────────────────────────────────────────────────────────

_SLASHED = re.compile(r"^\s*(\d{1,4})\s*/\s*(\d{1,2})\s*/\s*(\d{2,4})\s*$")
# Google's epoch. Only reachable if a caller switches to unformatted values.
_SERIAL_EPOCH = date(1899, 12, 30)


def detect_date_order(values) -> str:
    """'mdy' or 'dmy', decided across the whole column rather than per row.

    "9/23/2025" can only be month-first and "23/9/2025" can only be day-first,
    but "1/5/2026" is both. Deciding per row would read a column as a mixture of
    the two and put leads in the wrong month; deciding per sheet from the rows
    that *are* unambiguous gets every row right or none of them.

    Month-first when nothing settles it — that is what the trackers use today.
    """
    first_over_12 = second_over_12 = False
    for value in values:
        m = _SLASHED.match(str(value or ""))
        if not m:
            continue
        a, b = int(m.group(1)), int(m.group(2))
        if a > 12:
            first_over_12 = True
        if b > 12:
            second_over_12 = True
    if first_over_12 and not second_over_12:
        return "dmy"
    return "mdy"


def parse_day(value, order: str = "mdy") -> date | None:
    """One cell to a date, or None. Never guesses a date for an unreadable cell."""
    text = str(value or "").strip()
    if not text:
        return None
    m = _SLASHED.match(text)
    if m:
        a, b, c = (int(m.group(1)), int(m.group(2)), int(m.group(3)))
        day, month = (a, b) if order == "dmy" else (b, a)
        year = c + 2000 if c < 100 else c
        try:
            return date(year, month, day)
        except ValueError:
            return None
    try:                                    # a real date cell, or ISO text
        return date.fromisoformat(text[:10])
    except ValueError:
        pass
    if re.fullmatch(r"\d{5}", text):        # a Sheets serial number
        return _SERIAL_EPOCH + timedelta(days=int(text))
    return None


# ── ICP ──────────────────────────────────────────────────────────────────────

ICP_MIN, ICP_MAX = 1, 10

# Everything an account manager types to mean "I have not graded this one".
_UNRATED = frozenset({"", "-", "–", "—", "n/a", "na", "tbc", "tbd", "?"})

# Plain numeric ranges, no adjectives. These labels go in front of a paying
# client, and "Weak fit" against a lead they are excited about is a conversation
# nobody wants to have on our behalf.
ICP_BANDS: tuple[tuple[str, str, int, int], ...] = (
    ("9_10", "ICP 9–10", 9, 10),
    ("7_8",  "ICP 7–8",  7, 8),
    ("5_6",  "ICP 5–6",  5, 6),
    ("1_4",  "ICP 1–4",  1, 4),
)
UNRATED_BAND = ("unrated", "Not yet rated")

# What "top rated" means wherever it is counted.
TOP_RATED_FROM = 8


def parse_icp(value) -> int | None:
    """A grade, or None for ungraded. Never 0 — see the module docstring."""
    text = str(value or "").strip().lower()
    if text in _UNRATED:
        return None
    try:
        score = int(round(float(text)))
    except (TypeError, ValueError):
        return None
    return score if ICP_MIN <= score <= ICP_MAX else None


def icp_band(score: int | None) -> str:
    if score is None:
        return UNRATED_BAND[0]
    for key, _label, lo, hi in ICP_BANDS:
        if lo <= score <= hi:
            return key
    return UNRATED_BAND[0]


# ── Headcount ────────────────────────────────────────────────────────────────

# Taken from the Performance Tracker tab's own list, so a band reads and sorts
# the way it does on the sheet the account manager is looking at.
CANONICAL_HEADCOUNT_BANDS: tuple[str, ...] = (
    "1-10", "11-50", "51-200", "201-500", "501-1000",
    "1001-5000", "5001-10000", "10001+",
)

UNSPECIFIED = "Unspecified"
UNRECOGNISED = "Unrecognised"

# Past any real headcount, and finite so the two stay distinct.
_SORT_UNSPECIFIED = 1e15
_SORT_UNRECOGNISED = 1e16

_BAND_RANGE = re.compile(r"^\s*(\d[\d,]*)\s*[-–—]\s*(\d[\d,]*)\s*$")
_BAND_OPEN = re.compile(r"^\s*(\d[\d,]*)\s*\+\s*$")


def headcount_order(band: str) -> float:
    """Sort key. The lower bound, never the string.

    Sorted as text, "1001-5000" lands between "1-10" and "11-50" and the chart
    reads as nonsense.
    """
    text = (band or "").strip()
    m = _BAND_RANGE.match(text) or _BAND_OPEN.match(text)
    if m:
        return float(m.group(1).replace(",", ""))
    # Finite sentinels: float("inf") - 1 IS float("inf"), so the two collapsed
    # into one and the pair had no order between them.
    if text == UNSPECIFIED:
        return _SORT_UNSPECIFIED
    return _SORT_UNRECOGNISED


def clean_band(value) -> str:
    """Normalise a headcount cell, keeping an unreadable one visible.

    Google will coerce a typed "1-10" into a date, so bands do arrive looking
    like "1/10/2025". That lands in Unrecognised and gets surfaced, because the
    fix is in the sheet and silently dropping it would hide that.
    """
    text = " ".join(str(value or "").strip().split())
    if not text:
        return UNSPECIFIED
    m = _BAND_RANGE.match(text)
    if m:
        return f"{int(m.group(1).replace(',', ''))}-{int(m.group(2).replace(',', ''))}"
    m = _BAND_OPEN.match(text)
    if m:
        return f"{int(m.group(1).replace(',', ''))}+"
    return UNRECOGNISED


# ── Seniority ────────────────────────────────────────────────────────────────
#
# A heuristic, and labelled as one. 34 distinct job titles across 45 leads means
# a breakdown by raw title is one bar per lead and says nothing; "are we
# reaching decision makers" is the question the ICP grade is actually about.
#
# ORDER IS THE WHOLE THING. First match wins, most senior first, so "Managing
# Director" is a director and not caught by "manager", and "Founding Partner" is
# a founder rather than a partner. Expect to edit this table as real titles
# arrive — that is what it is for.
SENIORITY_RULES: tuple[tuple[str, str, tuple[str, ...]], ...] = (
    ("founder",  "Founder / Owner", ("founder", "founding", "co-founder",
                                     "cofounder", "owner", "proprietor")),
    ("c_level",  "C-level",         ("ceo", "cto", "coo", "cfo", "cmo", "cio",
                                     "cpo", "chro", "cro", "chief", "president")),
    ("vp",       "VP",              ("vp", "svp", "evp", "vice president")),
    ("partner",  "Partner",         ("partner",)),
    ("director", "Director",        ("director", "md")),
    ("head",     "Head of",         ("head",)),
    ("manager",  "Manager",         ("manager", "supervisor")),
    ("ic",       "Individual contributor",
     ("specialist", "executive", "coordinator", "analyst", "engineer",
      "consultant", "associate", "officer")),
)

SENIORITY_ORDER: tuple[str, ...] = tuple(key for key, _l, _k in SENIORITY_RULES) + \
    ("unclassified",)
SENIORITY_LABELS: dict[str, str] = {k: l for k, l, _ in SENIORITY_RULES}
SENIORITY_LABELS["unclassified"] = "Not classified"

_COMPILED = [(key, [re.compile(rf"\b{re.escape(kw)}\b") for kw in words])
             for key, _label, words in SENIORITY_RULES]

# Trailing geography, so "Head of People UK" and "Head of People" are one row in
# the ranked title list.
_GEO_SUFFIX = re.compile(
    r"[\s,(\-]+(uk|us|usa|emea|apac|dach|anz|eu|europe|global|"
    r"united kingdom|united states|ireland)\)?\s*$", re.I)


def clean_role(value) -> str:
    text = " ".join(str(value or "").strip().split())
    previous = None
    while text and text != previous:
        previous = text
        text = _GEO_SUFFIX.sub("", text).strip(" ,-(")
    return text


def seniority(title) -> str:
    """Bucket a job title. 'unclassified' is a real answer, not a dustbin.

    Folding unmatched titles into individual contributor would invent a figure.
    An honest "we could not classify these seven" is defensible; a wrong IC
    count on a client's report is not.
    """
    text = clean_role(title).lower()
    if not text:
        return "unclassified"
    for key, patterns in _COMPILED:
        if any(p.search(text) for p in patterns):
            return key
    return "unclassified"


# ── Channel ──────────────────────────────────────────────────────────────────

NOT_RECORDED = "Not recorded"

_CHANNELS: dict[str, tuple[str, str]] = {
    "smartlead":    ("smartlead", "Smartlead"),
    "linkedin":     ("meet_alfred", "LinkedIn"),
    "meet alfred":  ("meet_alfred", "LinkedIn"),
    "meetalfred":   ("meet_alfred", "LinkedIn"),
    "manual":       ("manual", "DNC & Merger"),
    "email":        ("smartlead", "Smartlead"),
}


def clean_channel(value) -> tuple[str, str]:
    """(key, label). A blank is its own visible bucket, never a dropped row —
    the channel split has to add up to the total or a reader will notice."""
    text = " ".join(str(value or "").strip().split())
    if not text:
        return "unspecified", NOT_RECORDED
    hit = _CHANNELS.get(text.lower())
    if hit:
        return hit
    return re.sub(r"[^a-z0-9]+", "_", text.lower()).strip("_"), text


# ── Rows ─────────────────────────────────────────────────────────────────────

def _text(row: dict, columns: dict, field: str) -> str:
    column = columns.get(field)
    if not column:
        return ""
    return " ".join(str(row.get(column, "") or "").strip().split())


def parse_rows(values: list[list[str]]) -> tuple[list[dict], dict]:
    """A sheet's grid → lead rows + metadata about the sheet itself.

    Takes the raw grid rather than pre-keyed dicts, because finding the header
    row is itself one of this module's jobs — see find_header.

    Every value on a returned row is a JSON primitive, because the same list is
    both what `summarise` counts and what gets handed to the browser for
    filtering. The browser then needs no business rules at all: it filters and
    counts fields that are already decided here.
    """
    header_index, header_row = find_header(values or [])
    raw, headers = rows_to_dicts(values or [])
    columns, warnings = resolve_columns(headers)
    meta = {
        "columns":     columns,
        "header_row":  header_index + 1,     # 1-based, as the sheet numbers them
        "warnings":  warnings,
        "has_dates": "date" in columns,
        "date_order": "mdy",
        "skipped":   0,
        "unrecognised_bands": 0,
    }
    if "date" not in columns:
        # Without a date nothing can be tied to a reporting period, and a
        # period figure built from undated rows would be fiction.
        return [], meta

    date_column = columns["date"]
    order = detect_date_order([r.get(date_column) for r in raw])
    meta["date_order"] = order

    rows: list[dict] = []
    for raw_row in raw:
        day = parse_day(raw_row.get(date_column), order)
        if day is None:
            # Not a lead. The live sheet is ~2,500 rows of filled-down formula
            # under a few dozen real ones.
            meta["skipped"] += 1
            continue
        score = parse_icp(_text(raw_row, columns, "icp"))
        band = clean_band(_text(raw_row, columns, "headcount"))
        if band == UNRECOGNISED:
            meta["unrecognised_bands"] += 1
        channel_key, channel_label = clean_channel(_text(raw_row, columns, "channel"))
        role = clean_role(_text(raw_row, columns, "role"))
        rows.append({
            "day":             day.isoformat(),
            "lead":            _text(raw_row, columns, "lead"),
            "channel":         channel_key,
            "channel_label":   channel_label,
            "headcount":       band,
            "headcount_order": headcount_order(band),
            "role":            role,
            "seniority":       seniority(role),
            "industry":        _text(raw_row, columns, "industry") or UNSPECIFIED,
            "location":        _text(raw_row, columns, "location") or UNSPECIFIED,
            "icp":             score,
            "icp_band":        icp_band(score),
            "category":        _text(raw_row, columns, "category") or UNSPECIFIED,
            "stage":           _text(raw_row, columns, "stage"),
        })
    rows.sort(key=lambda r: r["day"], reverse=True)
    return rows, meta


def in_period(rows: list[dict], start: date | None, end: date | None) -> list[dict]:
    """The leads dated inside a report's period. Inclusive at both ends."""
    lo = start.isoformat() if start else ""
    hi = end.isoformat() if end else ""
    return [r for r in rows
            if (not lo or r["day"] >= lo) and (not hi or r["day"] <= hi)]


# ── Summarising ──────────────────────────────────────────────────────────────

TOP_N = 8


def _share(count: int, total: int) -> float:
    return round(count * 100.0 / total, 1) if total else 0.0


def _tally(rows: list[dict], field: str, *, order=None, top: int | None = None,
           labels: dict | None = None) -> list[dict]:
    """Count by a field. Ordered by `order` if given, else by count descending.

    `top` collapses the tail into "Other" — but never a tail of one, because
    "Other: 1" hides a name for no gain.
    """
    counts: dict[str, int] = {}
    for row in rows:
        counts[row.get(field) or UNSPECIFIED] = counts.get(row.get(field) or UNSPECIFIED, 0) + 1
    total = len(rows)

    if order is not None:
        items = sorted(counts.items(), key=lambda kv: order(kv[0]))
    else:
        items = sorted(counts.items(), key=lambda kv: (-kv[1], kv[0].lower()))

    out = [{"key": k, "label": (labels or {}).get(k, k),
            "count": n, "pct": _share(n, total)} for k, n in items]

    if top is not None and len(out) > top + 1:
        head, tail = out[:top], out[top:]
        spare = sum(e["count"] for e in tail)
        head.append({"key": "other", "label": "Other", "count": spare,
                     "pct": _share(spare, total), "distinct": len(tail)})
        return head
    return out


def summarise(rows: list[dict]) -> dict:
    """The figures every view of this data is built from.

    One function, three consumers: the account manager's panel, the client's
    report, and the tests. They cannot disagree because there is nothing to
    disagree with.
    """
    total = len(rows)
    graded = [r["icp"] for r in rows if r["icp"] is not None]
    graded.sort()

    average = round(sum(graded) / len(graded), 2) if graded else None
    if graded:
        middle = len(graded) // 2
        median = float(graded[middle]) if len(graded) % 2 else \
            round((graded[middle - 1] + graded[middle]) / 2, 1)
    else:
        median = None

    band_counts = {key: 0 for key, _l, _lo, _hi in ICP_BANDS}
    band_counts[UNRATED_BAND[0]] = 0
    for row in rows:
        band_counts[row["icp_band"]] = band_counts.get(row["icp_band"], 0) + 1
    bands = [{"key": key, "label": label, "count": band_counts.get(key, 0),
              "pct": _share(band_counts.get(key, 0), total)}
             for key, label, _lo, _hi in ICP_BANDS]
    bands.append({"key": UNRATED_BAND[0], "label": UNRATED_BAND[1],
                  "count": band_counts.get(UNRATED_BAND[0], 0),
                  "pct": _share(band_counts.get(UNRATED_BAND[0], 0), total)})

    roles: dict[str, int] = {}
    for row in rows:
        if row["role"]:
            roles[row["role"]] = roles.get(row["role"], 0) + 1

    return {
        "total":     total,
        "rated":     len(graded),
        # Stated everywhere the average is, so nobody reads a grade for a third
        # of the leads as a grade for all of them.
        "unrated":   total - len(graded),
        "average":   average,
        "median":    median,
        "top_rated": sum(1 for s in graded if s >= TOP_RATED_FROM),
        "bands":     bands,
        "headcount": _tally(rows, "headcount", order=headcount_order),
        "industry":  _tally(rows, "industry", top=TOP_N),
        "seniority": _tally(rows, "seniority",
                            order=lambda k: SENIORITY_ORDER.index(k)
                            if k in SENIORITY_ORDER else len(SENIORITY_ORDER),
                            labels=SENIORITY_LABELS),
        "location":  _tally(rows, "location", top=TOP_N),
        "channel":   _tally(rows, "channel", labels={
            r["channel"]: r["channel_label"] for r in rows}),
        "category":  _tally(rows, "category"),
        "roles":     [{"label": l, "count": n} for l, n in
                      sorted(roles.items(), key=lambda kv: (-kv[1], kv[0].lower()))],
    }

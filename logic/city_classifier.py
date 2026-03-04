from __future__ import annotations

import json
import re
import unicodedata
from pathlib import Path
from typing import Iterable, Optional, Set, Tuple


def _strip_accents(text: str) -> str:
    """
    Convert accented characters to their closest ASCII equivalents.
    Deterministic and offline.
    """
    norm = unicodedata.normalize("NFKD", text)
    return "".join(ch for ch in norm if not unicodedata.combining(ch))


def normalize_city_key(value: Optional[str]) -> str:
    """
    Normalise a location string for deterministic matching.

    - Lowercase
    - Strip accents
    - Remove punctuation
    - Collapse whitespace
    """
    if value is None:
        return ""
    s = str(value).strip()
    if not s:
        return ""
    s = _strip_accents(s).lower()

    # Common cleanup: treat "&" as "and", remove punctuation
    s = s.replace("&", " and ")
    s = re.sub(r"[^\w\s]", " ", s)  # punctuation -> space
    s = re.sub(r"\s+", " ", s).strip()
    return s


def load_city_whitelist(path: str | Path) -> Tuple[Set[str], Set[str]]:
    """
    Load a JSON whitelist and build:
      - raw_set: original entries (for reference)
      - key_set: normalised keys (used for matching)
    """
    p = Path(path)
    cities = json.loads(p.read_text(encoding="utf-8"))
    raw_set = set(str(c).strip() for c in cities if str(c).strip())
    key_set = set(normalize_city_key(c) for c in raw_set if normalize_city_key(c))
    return raw_set, key_set


def is_business_city(city_value: Optional[str], whitelist_key_set: Set[str]) -> bool:
    """
    Conservative whitelist-first classification.
    """
    key = normalize_city_key(city_value)
    if not key:
        return False
    return key in whitelist_key_set


def choose_location_output(
    city_value: Optional[str],
    state_value: Optional[str],
    whitelist_key_set: Set[str],
) -> str:
    """
    Output either city or state/region based on whitelist classification.

    Rules:
      - If city is whitelisted -> return city (trimmed)
      - Else if state is present -> return state (trimmed)
      - Else -> return empty string
    """
    city_str = "" if city_value is None else str(city_value).strip()
    state_str = "" if state_value is None else str(state_value).strip()

    if is_business_city(city_str, whitelist_key_set):
        return city_str
    if state_str:
        return state_str
    return ""

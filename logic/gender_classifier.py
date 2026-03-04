from __future__ import annotations

from typing import Optional

# NOTE:
# This module intentionally does NOT implement or change any gender classification logic.
# It is a thin wrapper around the existing deterministic classifier implementation.
#
# Required file (must exist in the repo):
#   classifier.py  -> must expose: classify_first_name(value) -> "m" | "f" | "unknown"

try:
    from classifier import classify_first_name  # type: ignore
except Exception as e:  # pragma: no cover
    classify_first_name = None  # type: ignore
    _IMPORT_ERROR = e
else:
    _IMPORT_ERROR = None


def classify(value: Optional[str]) -> str:
    """Classify a first name as 'm', 'f', or 'unknown' using the existing deterministic classifier."""
    if classify_first_name is None:  # pragma: no cover
        raise ImportError(
            "Missing dependency: classifier.py (expected classify_first_name). "
            "Add the original classifier file(s) to the repo root."
        ) from _IMPORT_ERROR

    return classify_first_name(value)

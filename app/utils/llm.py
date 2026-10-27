"""
Server-side LLM calls (OpenAI Chat Completions).

Plain `requests`, like every other integration here, rather than the OpenAI
SDK: no new dependency, and the offline test harness already patches
`requests`, so a test that forgets to stub this fails loudly instead of
calling out.

Every call asks for JSON that matches a schema (structured outputs), so the
caller gets a dict back or an LLMError — never free text to parse.

Configuration:
  OPENAI_API_KEY   required; is_configured() is False without it
  OPENAI_MODEL     optional, default gpt-4o (the model Reply Bank uses)
  OPENAI_API_BASE  optional, default https://api.openai.com/v1
"""
from __future__ import annotations

import json
import logging
import os
import re
import time

import requests

logger = logging.getLogger(__name__)

DEFAULT_MODEL = "gpt-4o"
_TIMEOUT_S = 120
# Calls run in background threads, so waiting out a rate limit is cheaper
# than failing: a new OpenAI account allows about one long transcript a
# minute, and a meeting's analysis and its notes draft go back to back.
_MAX_ATTEMPTS = 5
_MAX_WAIT_S = 65


class LLMError(RuntimeError):
    """The model could not be reached or did not return usable JSON."""


class RequestTooLarge(LLMError):
    """The prompt is bigger than the account's per-minute token limit.

    Waiting won't help; the caller has to send less.
    """


class OutOfCredit(LLMError):
    """The OpenAI account has no credit or billing set up."""


def is_configured() -> bool:
    return bool(os.environ.get("OPENAI_API_KEY", "").strip())


def model_name() -> str:
    return os.environ.get("OPENAI_MODEL", "").strip() or DEFAULT_MODEL


def _duration_s(value: str) -> float | None:
    """Seconds from OpenAI's reset headers: "1m20.5s", "6s", "250ms"."""
    if not value:
        return None
    total, found = 0.0, False
    for number, unit in re.findall(r"([\d.]+)\s*(ms|h|m|s)", value):
        found = True
        total += float(number) * {"ms": 0.001, "s": 1, "m": 60, "h": 3600}[unit]
    return total if found else None


def _error(r) -> tuple[str, str]:
    try:
        err = (r.json() or {}).get("error") or {}
    except ValueError:
        return "", (r.text or "")[:300]
    return str(err.get("code") or err.get("type") or ""), str(err.get("message") or "")


def _rate_limit_wait(r, message: str) -> float:
    """How long OpenAI asks us to wait, from its headers or its message."""
    headers = r.headers or {}
    waits: list[float] = []
    for name, scale in (("retry-after-ms", 0.001), ("retry-after", 1.0)):
        raw = str(headers.get(name, "")).strip()
        if raw.replace(".", "", 1).isdigit():
            waits.append(float(raw) * scale)
    for name in ("x-ratelimit-reset-tokens", "x-ratelimit-reset-requests"):
        wait = _duration_s(str(headers.get(name, "")))
        if wait is not None:
            waits.append(wait)
    match = re.search(r"try again in ([\d.]+\s*(?:ms|s|m))", message)
    if match:
        waits.append(_duration_s(match.group(1)) or 0)
    return min(_MAX_WAIT_S, max(waits) + 1) if waits else 20


def complete_json(system: str, user: str, *, schema: dict, schema_name: str,
                  max_tokens: int = 4000) -> dict:
    """One chat completion constrained to `schema`. Returns the parsed object.

    Rate limits are waited out (OpenAI says how long); timeouts and 5xx are
    retried; a request too large for the account, or an account with no
    credit, fails at once with an error that says which.
    """
    key = os.environ.get("OPENAI_API_KEY", "").strip()
    if not key:
        raise LLMError("OpenAI isn't connected. Add OPENAI_API_KEY in Render.")
    base = (os.environ.get("OPENAI_API_BASE", "").strip()
            or "https://api.openai.com/v1").rstrip("/")
    body = {
        "model": model_name(),
        "temperature": 0.2,
        "max_tokens": max_tokens,
        "messages": [
            {"role": "system", "content": system},
            {"role": "user", "content": user},
        ],
        "response_format": {
            "type": "json_schema",
            "json_schema": {"name": schema_name, "strict": True, "schema": schema},
        },
    }
    last = ""
    for attempt in range(_MAX_ATTEMPTS):
        final = attempt == _MAX_ATTEMPTS - 1
        try:
            r = requests.post(
                f"{base}/chat/completions",
                headers={"Authorization": f"Bearer {key}",
                         "Content-Type": "application/json"},
                json=body,
                timeout=_TIMEOUT_S,
            )
        except requests.RequestException as exc:
            last = f"Couldn't reach OpenAI ({exc.__class__.__name__})."
            if not final:
                time.sleep(5)
            continue

        if r.status_code == 429:
            code, message = _error(r)
            logger.warning("OpenAI 429 (%s): %s", code, message[:300])
            if code == "insufficient_quota" or "quota" in message.lower():
                raise OutOfCredit(
                    "OpenAI says this account is out of credit. Add credit or a payment "
                    "method at platform.openai.com (Settings → Billing), then retry.")
            # "Request too large ... Limit 30000, Requested 34567" can never
            # succeed; "Rate limit reached ... try again in 6s" will.
            if "request too large" in message.lower():
                raise RequestTooLarge(
                    "This meeting is too long for the OpenAI account's per-minute limit.")
            last = ("OpenAI's rate limit was hit and didn't clear after several tries. "
                    "Retry in a few minutes.")
            if not final:
                time.sleep(_rate_limit_wait(r, message))
            continue

        if r.status_code >= 500:
            last = f"OpenAI had a server error (HTTP {r.status_code}). Retry in a few minutes."
            if not final:
                time.sleep(5)
            continue

        if r.status_code >= 400:
            # The body names the problem (bad key, unknown model); it goes to the
            # log, and the UI gets a sentence without the raw upstream text.
            code, message = _error(r)
            logger.warning("OpenAI HTTP %s (%s): %s", r.status_code, code, message[:500])
            if r.status_code == 401:
                raise LLMError("OpenAI rejected the API key. Check OPENAI_API_KEY in Render.")
            if code == "model_not_found":
                raise LLMError(f"This OpenAI account can't use the model {model_name()}. "
                               "Check OPENAI_MODEL in Render.")
            raise LLMError(f"OpenAI refused the request (HTTP {r.status_code}). "
                           "Check OPENAI_API_KEY and OPENAI_MODEL in Render.")
        try:
            choice = r.json()["choices"][0]
            message = choice["message"]
            if message.get("refusal"):
                raise LLMError("The model declined to summarise this meeting.")
            if choice.get("finish_reason") == "length":
                raise LLMError("The model's answer was cut off before it finished.")
            return json.loads(message["content"])
        except LLMError:
            raise
        except (KeyError, IndexError, TypeError, ValueError) as exc:
            raise LLMError("OpenAI returned an answer that wasn't valid JSON.") from exc
    raise LLMError(last or "OpenAI did not respond.")

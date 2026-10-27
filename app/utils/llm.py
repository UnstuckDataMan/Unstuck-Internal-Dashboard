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
import time

import requests

logger = logging.getLogger(__name__)

DEFAULT_MODEL = "gpt-4o"
_TIMEOUT_S = 120


class LLMError(RuntimeError):
    """The model could not be reached or did not return usable JSON."""


def is_configured() -> bool:
    return bool(os.environ.get("OPENAI_API_KEY", "").strip())


def model_name() -> str:
    return os.environ.get("OPENAI_MODEL", "").strip() or DEFAULT_MODEL


def complete_json(system: str, user: str, *, schema: dict, schema_name: str,
                  max_tokens: int = 4000) -> dict:
    """One chat completion constrained to `schema`. Returns the parsed object.

    Retries once on a timeout, a 429 or a 5xx; anything else fails at once,
    because a bad request will not get better by being sent again.
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
    for attempt in range(2):
        try:
            r = requests.post(
                f"{base}/chat/completions",
                headers={"Authorization": f"Bearer {key}",
                         "Content-Type": "application/json"},
                json=body,
                timeout=_TIMEOUT_S,
            )
        except requests.RequestException as exc:
            last = f"OpenAI request failed: {exc.__class__.__name__}"
            time.sleep(2)
            continue
        if r.status_code == 429 or r.status_code >= 500:
            last = f"OpenAI returned HTTP {r.status_code}"
            time.sleep(5 if attempt == 0 else 0)
            continue
        if r.status_code >= 400:
            # The body names the problem (bad key, unknown model); it goes to the
            # log, and the UI gets a sentence without the raw upstream text.
            logger.warning("OpenAI HTTP %s: %s", r.status_code, (r.text or "")[:500])
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

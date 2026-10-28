"""Small, server-only writing provider boundary; no provider owns the corpus.

The studio deliberately has its own model setting. A model is selected and
evaluated by the company, not inferred from an employee's chat subscription.
Neither adapter logs prompts, source text, credentials or upstream error bodies.
"""
from __future__ import annotations

import json
import os
import time
from typing import Protocol

import requests


class ProviderError(RuntimeError):
    pass


class WritingProvider(Protocol):
    name: str
    model: str

    def complete_json(self, system: str, user: str, *, schema: dict,
                      schema_name: str, max_tokens: int = 5000) -> dict: ...


def validate_json(value, schema: dict, path: str = "$", *, depth: int = 0):
    """Validate the deliberately small JSON-schema vocabulary used by the studio.

    This is applied to both providers and fixture adapters. It is not a general
    replacement for JSON Schema and refuses unsupported schema keywords.
    """
    allowed = {"type", "properties", "required", "additionalProperties", "items",
               "enum", "minimum", "maximum", "minItems", "maxItems", "minLength",
               "maxLength", "description"}
    if set(schema) - allowed or depth > 20:
        raise ProviderError("Unsupported output schema.")
    typ = schema.get("type")
    valid = {"object": isinstance(value, dict), "array": isinstance(value, list),
             "string": isinstance(value, str), "integer": type(value) is int,
             "number": type(value) in (int, float), "boolean": type(value) is bool,
             "null": value is None}
    if typ and not valid.get(typ, False):
        raise ProviderError(f"Invalid model output at {path}: expected {typ}.")
    if "enum" in schema and value not in schema["enum"]:
        raise ProviderError(f"Invalid model output at {path}: unsupported value.")
    if typ == "object":
        props = schema.get("properties", {})
        if any(key not in value for key in schema.get("required", [])):
            raise ProviderError(f"Invalid model output at {path}: required field missing.")
        if schema.get("additionalProperties") is False and set(value) - set(props):
            raise ProviderError(f"Invalid model output at {path}: unexpected field.")
        for key, item in value.items():
            if key in props:
                validate_json(item, props[key], path + "." + key, depth=depth + 1)
    elif typ == "array":
        if len(value) < schema.get("minItems", 0) or len(value) > schema.get("maxItems", 100000):
            raise ProviderError(f"Invalid model output at {path}: wrong number of items.")
        for i, item in enumerate(value):
            validate_json(item, schema.get("items", {}), f"{path}[{i}]", depth=depth + 1)
    elif typ == "string":
        if len(value) < schema.get("minLength", 0) or len(value) > schema.get("maxLength", 1000000):
            raise ProviderError(f"Invalid model output at {path}: text length outside limits.")
    elif typ in ("integer", "number"):
        if not schema.get("minimum", float("-inf")) <= value <= schema.get("maximum", float("inf")):
            raise ProviderError(f"Invalid model output at {path}: number outside limits.")
    return value


def _post(url: str, headers: dict, body: dict):
    # At most three attempts and bounded waits; a durable job handles later retry.
    from .costs import record_usage, start_call
    provider = "anthropic" if url.startswith("https://api.anthropic.com/") else "openai"
    for attempt in range(3):
        # Persist an unknown-cost attempt before transmission. A worker crash
        # after charging cannot incorrectly make the job appear free.
        call_id = start_call(provider, body.get("model"))
        try:
            response = requests.post(url, headers=headers, json=body, timeout=(10, 120))
        except requests.RequestException:
            record_usage(provider, body.get("model"), None, status="transport_error", call_id=call_id)
            if attempt == 2:
                raise ProviderError("The writing provider could not be reached. Retry this job.") from None
            time.sleep(2 ** attempt)
            continue
        try:
            payload = response.json()
        except (ValueError, TypeError):
            payload = None
        if not isinstance(payload, dict):
            payload = None
        # Meter even incomplete/refused responses, and before schema parsing.
        # A ledger failure propagates: a charged call must not disappear quietly.
        record_usage(provider, body.get("model"), payload.get("usage") if payload else None,
                     request_id=(payload.get("id") if payload else None)
                     or response.headers.get("x-request-id") or response.headers.get("request-id"),
                     status=(str(payload.get("status") or payload.get("stop_reason") or "received")
                             if payload and response.status_code < 400 else f"http_{response.status_code}"),
                     call_id=call_id)
        if response.status_code == 429 or response.status_code >= 500:
            if attempt == 2:
                raise ProviderError("The writing provider is rate-limited or unavailable. Retry later.")
            try:
                delay = min(15, max(1, float(response.headers.get("Retry-After", 2 ** attempt))))
            except (ValueError, TypeError):
                delay = 2 ** attempt
            time.sleep(delay)
            continue
        if response.status_code in (401, 403):
            raise ProviderError("The company writing API credential was refused. Check server configuration.")
        if response.status_code >= 400:
            raise ProviderError("The writing provider rejected this request. Check the configured model and account limits.")
        if payload is None:
            raise ProviderError("The writing provider returned an invalid response.")
        return payload
    raise ProviderError("The writing provider is unavailable.")


class OpenAIProvider:
    name = "openai"

    def __init__(self, model: str, key: str):
        self.model, self._key = model, key

    def complete_json(self, system, user, *, schema, schema_name, max_tokens=5000):
        payload = _post("https://api.openai.com/v1/responses", {
            "Authorization": f"Bearer {self._key}", "Content-Type": "application/json",
        }, {"model": self.model, "store": False,
            "input": [{"role": "system", "content": system}, {"role": "user", "content": user}],
            "max_output_tokens": max_tokens,
            "text": {"format": {"type": "json_schema", "name": schema_name,
                                  "strict": True, "schema": schema}}})
        if payload.get("status") in ("incomplete", "failed"):
            raise ProviderError("The writing response was incomplete. Reduce the request or increase its output budget.")
        texts = [part.get("text", "") for item in payload.get("output", [])
                 for part in item.get("content", []) if part.get("type") == "output_text"]
        try:
            result = json.loads("".join(texts))
        except (ValueError, TypeError):
            raise ProviderError("The writing provider did not return a complete structured draft.") from None
        return validate_json(result, schema)


class AnthropicProvider:
    name = "anthropic"

    def __init__(self, model: str, key: str):
        self.model, self._key = model, key

    def complete_json(self, system, user, *, schema, schema_name, max_tokens=5000):
        payload = _post("https://api.anthropic.com/v1/messages", {
            "x-api-key": self._key, "anthropic-version": "2023-06-01", "Content-Type": "application/json",
        }, {"model": self.model, "system": system, "max_tokens": max_tokens,
            "messages": [{"role": "user", "content": user}],
            "tools": [{"name": schema_name, "description": "Return the requested structured result.",
                       "input_schema": schema}],
            "tool_choice": {"type": "tool", "name": schema_name}})
        if payload.get("stop_reason") == "max_tokens":
            raise ProviderError("The writing response was incomplete. Reduce the request or increase its output budget.")
        matches = [part.get("input") for part in payload.get("content", [])
                   if part.get("type") == "tool_use" and part.get("name") == schema_name]
        if len(matches) != 1:
            raise ProviderError("The writing provider did not return one structured draft.")
        return validate_json(matches[0], schema)


def provider_status() -> dict:
    selected = os.environ.get("COPY_STUDIO_PROVIDER", "openai").strip().lower()
    model = (os.environ.get("COPY_STUDIO_MODEL", "").strip()
             or os.environ.get("OPENAI_MODEL" if selected == "openai" else "ANTHROPIC_MODEL", "").strip())
    key_name = "COPY_STUDIO_OPENAI_API_KEY" if selected == "openai" else "COPY_STUDIO_ANTHROPIC_API_KEY"
    fallback_key_name = "OPENAI_API_KEY" if selected == "openai" else "ANTHROPIC_API_KEY"
    supported = selected in {"openai", "anthropic"}
    key_set = bool((os.environ.get(key_name) or os.environ.get(fallback_key_name) or "").strip()) if supported else False
    return {"provider": selected, "model": model or None, "configured": supported and key_set and bool(model),
            "credential_configured": key_set, "model_configured": bool(model),
            "credential_variable": key_name, "credential_fallback": fallback_key_name,
            "evaluation_status": "Human blind copy review pending",
            "assembly_available": True,
            "message": "Ready for draft requests; writing quality requires human evaluation." if supported and key_set and model
            else f"Set COPY_STUDIO_MODEL and the company API credential in {key_name} (or existing {fallback_key_name}) on the server. Approved-template assembly works without a model."}


def get_provider() -> WritingProvider:
    info = provider_status()
    if not info["configured"]:
        raise ProviderError(info["message"])
    if info["provider"] == "openai":
        return OpenAIProvider(info["model"], (os.environ.get("COPY_STUDIO_OPENAI_API_KEY") or os.environ["OPENAI_API_KEY"]).strip())
    return AnthropicProvider(info["model"], (os.environ.get("COPY_STUDIO_ANTHROPIC_API_KEY") or os.environ["ANTHROPIC_API_KEY"]).strip())

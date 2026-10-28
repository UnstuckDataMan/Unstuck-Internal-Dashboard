"""Canonical Client Profiles metadata, without inventing verified website facts."""
from __future__ import annotations

import ipaddress
import hashlib
import json
import re
from urllib.parse import urlsplit, urlunsplit

from app.utils.pulse.readai import FREE_MAIL_DOMAINS


def profile_domain(value) -> str:
    """A public-looking profile domain; never infer one from a sending inbox."""
    text = str(value or "").strip()
    if not text or "@" in text or any(c.isspace() for c in text):
        return ""
    if "://" in text:
        try:
            parsed = urlsplit(text)
        except ValueError:
            return ""
        if parsed.scheme.lower() not in {"http", "https"} or parsed.username or parsed.password:
            return ""
        host = parsed.hostname or ""
    else:
        # Profiles store domains, not arbitrary URL paths/ports or email URLs.
        host = text.split("/", 1)[0]
    host = host.lower().rstrip(".")
    if host.startswith("www."):
        host = host[4:]
    try:
        host = host.encode("idna").decode("ascii")
        ipaddress.ip_address(host)
    except ValueError:
        pass
    except UnicodeError:
        return ""
    else:
        return ""  # an IP address is not the client's canonical domain
    if len(host) > 253 or "." not in host or host.endswith((".local", ".localhost", ".internal", ".test", ".invalid")):
        return ""
    labels = host.split(".")
    if any(not re.fullmatch(r"[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?", label) for label in labels):
        return ""
    return host


def _website(value) -> str:
    text = str(value or "").strip()
    try:
        parsed = urlsplit(text)
        if parsed.scheme.lower() not in {"http", "https"} or parsed.username or parsed.password or parsed.port not in (None, 80, 443):
            return ""
        domain = profile_domain(text)
        if not domain or domain in FREE_MAIL_DOMAINS:
            return ""
        return urlunsplit((parsed.scheme.lower(), domain, parsed.path or "", parsed.query, ""))
    except ValueError:
        return ""


def profile_metadata(row: dict) -> dict:
    raw_domains = row.get("domains") if isinstance(row.get("domains"), list) else []
    domains = list(dict.fromkeys(domain for domain in map(profile_domain, raw_domains) if domain and domain not in FREE_MAIL_DOMAINS))
    explicit = _website(row.get("website_url") or row.get("website"))
    inferred = not explicit and bool(domains)
    website = explicit or ("https://"+domains[0] if domains else None)
    limitations = []
    if inferred:
        limitations.append("Website link is derived from the Client Profiles domain; the official site and its content have not been verified.")
    elif not website:
        limitations.append("Add the client's domain in Client Profiles to identify its website. Sending inboxes are not treated as website evidence.")
    else:
        limitations.append("Website comes from Client Profiles; site content has not been fetched or independently verified.")
    identifier = str(row.get("id") or "")
    fingerprint = hashlib.sha256(json.dumps({"id": identifier, "domains": domains, "updated_at": row.get("updated_at")},
        sort_keys=True, ensure_ascii=False, separators=(",", ":")).encode()).hexdigest()
    return {
        "id": identifier, "name": str(row.get("name") or row.get("id") or ""),
        "domains": domains, "website_url": website,
        "website_source": "client_profile_website" if explicit else "client_profile_domain" if inferred else "missing",
        "website_inferred": inferred, "website_verified": False,
        "profile_url": "/client-profiles", "profile_updated_at": row.get("updated_at"),
        "profile_fingerprint": fingerprint,
        "territories": [str(v) for v in row.get("territories", [])] if isinstance(row.get("territories"), list) else [],
        "industries": [str(v) for v in row.get("industries", [])] if isinstance(row.get("industries"), list) else [],
        "senders": [str(v) for v in row.get("senders", [])] if isinstance(row.get("senders"), list) else [],
        "source": "client_profiles", "limitations": limitations,
    }

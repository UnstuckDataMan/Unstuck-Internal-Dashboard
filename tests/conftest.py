"""
Offline test harness for the DNC & Merge tool.

No network access: every Supabase call goes through FakeSupabase (patched over
the `requests` module functions all routers share), and every Google Sheets
interaction is replaced with in-memory fakes.  Any request to an unexpected
host raises immediately.
"""
from __future__ import annotations

import base64
import json
import os
import sys
from pathlib import Path

import pytest

# ── Environment must be set BEFORE app modules are imported ───────────────────
REPO_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(REPO_ROOT))

TEST_SUPABASE_URL = "http://supabase.test"
os.environ["SUPABASE_URL"]      = TEST_SUPABASE_URL
os.environ["SUPABASE_ANON_KEY"] = "test-anon-key"
# Auth is off by default for the existing offline endpoint tests; the dedicated
# auth/RBAC tests re-enable it per-test via monkeypatch. SESSION_SECRET is fixed
# so SessionMiddleware is deterministic across the suite.
os.environ.setdefault("AUTH_DISABLED", "1")
os.environ.setdefault("SESSION_SECRET", "test-session-secret")
# Disable auto-sync pacing so the sequential-loop tests don't sleep between
# campaigns (pacing is a production quota measure, not test-relevant).
os.environ["AUTO_SYNC_PACING_MS"] = "0"
# Non-empty so google_sheets.is_configured() returns True; only decoded if a
# test forgets to patch the gspread layer (decoding works, auth never runs).
os.environ["GOOGLE_SHEETS_SA_JSON"] = base64.b64encode(
    json.dumps({"type": "service_account"}).encode()
).decode()

import requests  # noqa: E402


# ── Fake HTTP layer ────────────────────────────────────────────────────────────

class FakeResponse:
    def __init__(self, status_code=200, json_data=None, headers=None, text=None):
        self.status_code = status_code
        self._json   = [] if json_data is None else json_data
        self.headers = headers or {}
        self.text    = json.dumps(self._json) if text is None else text
        self.ok      = status_code < 400

    def json(self):
        return self._json

    def raise_for_status(self):
        if self.status_code >= 400:
            raise requests.HTTPError(f"HTTP {self.status_code}: {self.text[:120]}")


def params_list(call) -> list[tuple[str, str]]:
    """Normalise a recorded call's params (dict or list of tuples) to a list."""
    p = call["params"]
    if p is None:
        return []
    if isinstance(p, dict):
        return list(p.items())
    return list(p)


def param_values(call, key: str) -> list[str]:
    return [v for k, v in params_list(call) if k == key]


def in_list_values(raw: str) -> list[str]:
    """Parse a PostgREST `in.(...)` filter value into its items (unquoting)."""
    assert raw.startswith("in.(") and raw.endswith(")"), raw
    inner = raw[4:-1]
    if not inner:
        return []
    items = []
    for part in inner.split(","):
        part = part.strip()
        if part.startswith('"') and part.endswith('"'):
            part = part[1:-1].replace('\\"', '"').replace("\\\\", "\\")
        items.append(part)
    return items


class FakeSupabase:
    """Recording fake for the Supabase REST API.

    Tests register handlers per (method, table); unmatched table calls get a
    safe default (GET -> [], writes -> 204).  Calls to any URL outside the
    fake Supabase host raise — guaranteeing the suite runs fully offline.
    """

    def __init__(self):
        self.calls: list[dict] = []
        self._routes: dict[tuple[str, str], callable] = {}

    def route(self, method: str, table: str, handler):
        """handler(call) -> FakeResponse.  `table` matches the path segment."""
        self._routes[(method.upper(), table)] = handler

    def calls_to(self, method: str, table: str) -> list[dict]:
        return [c for c in self.calls
                if c["method"] == method.upper() and c["table"] == table]

    def _dispatch(self, method: str, url: str, **kwargs):
        if not url.startswith(TEST_SUPABASE_URL):
            raise AssertionError(f"Unexpected network call in offline test: {method} {url}")
        table = url.rsplit("/rest/v1/", 1)[-1].split("?")[0]
        call = {
            "method":  method.upper(),
            "url":     url,
            "table":   table,
            "params":  kwargs.get("params"),
            "json":    kwargs.get("json"),
            "headers": kwargs.get("headers") or {},
        }
        self.calls.append(call)
        handler = self._routes.get((method.upper(), table))
        if handler is not None:
            return handler(call)
        if method.upper() == "GET":
            return FakeResponse(200, [])
        return FakeResponse(204, [])


@pytest.fixture
def fake_sb(monkeypatch):
    fake = FakeSupabase()
    monkeypatch.setattr(requests, "get",    lambda url, **kw: fake._dispatch("GET", url, **kw))
    monkeypatch.setattr(requests, "post",   lambda url, **kw: fake._dispatch("POST", url, **kw))
    monkeypatch.setattr(requests, "patch",  lambda url, **kw: fake._dispatch("PATCH", url, **kw))
    monkeypatch.setattr(requests, "delete", lambda url, **kw: fake._dispatch("DELETE", url, **kw))
    return fake


@pytest.fixture
def client():
    """TestClient WITHOUT lifespan (no APScheduler) — endpoints only."""
    from starlette.testclient import TestClient
    from app.main import app
    return TestClient(app, raise_server_exceptions=True)


@pytest.fixture(autouse=True)
def _reset_sheets_state():
    """Clear the google_sheets module caches and quota windows between tests.

    They are deliberately process-global in production (one service account =
    one shared quota), so without this a cached worksheet handle or a spent
    read budget would leak from one test into the next.
    """
    import app.utils.google_sheets as gs

    def _clear():
        gs._records_cache.clear()
        gs._ws_cache.clear()
        gs._read_limiter._hits.clear()
        gs._write_limiter._hits.clear()
        gs.reset_sheet_reads()

    _clear()
    yield
    _clear()


# ── Fake gspread / Sheets REST layer ──────────────────────────────────────────

class FakeWorksheet:
    def __init__(self, headers: list[str], rows: list[list[str]], gid: int = 777,
                 col_count: int | None = None):
        self.headers = headers
        self.rows    = rows               # data rows (row 2 onward)
        self.id      = gid
        # Physical grid width — defaults to exactly the header width, which is
        # the tight-grid case that used to 400 on column appends.
        self.col_count = col_count if col_count is not None else max(1, len(headers))
        self.added_cols: list[int] = []
        self.batch_updates: list[list] = []
        self.updates: list[tuple] = []    # generic recorded calls
        # batch_get payload (read_stats_cells); set to an Exception to raise.
        self.batch_get_result: object = [[]]

    def add_cols(self, n: int):
        self.added_cols.append(n)
        self.col_count += n

    def row_values(self, n: int):
        if n == 1:
            return list(self.headers)
        return list(self.rows[n - 2])

    def col_values(self, idx: int):
        col = idx - 1
        out = [self.headers[col] if col < len(self.headers) else ""]
        for r in self.rows:
            out.append(r[col] if col < len(r) else "")
        return out

    def batch_update(self, updates):
        self.batch_updates.append(updates)

    def get_all_values(self, value_render_option=None, **kwargs):
        return [list(self.headers)] + [list(r) for r in self.rows]

    def batch_get(self, ranges, **kwargs):
        self.updates.append(("batch_get", ranges))
        if isinstance(self.batch_get_result, Exception):
            raise self.batch_get_result
        return self.batch_get_result

    def get_all_records(self, **kwargs):
        return [dict(zip(self.headers, r)) for r in self.rows]

    # create_outreach_sheet surface
    def update_title(self, title):  self.updates.append(("title", title))
    def update(self, values, rng):  self.updates.append(("update", rng, values))
    def format(self, rng, fmt):     self.updates.append(("format", rng, fmt))
    def freeze(self, rows):         self.updates.append(("freeze", rows))
    def update_cell(self, r, c, v): self.updates.append(("cell", r, c, v))


class FakeSpreadsheet:
    """Stands in for gspread's Spreadsheet.

    `sheet1` is a PROPERTY, not an attribute, because that is what it is in
    gspread: reading it calls fetch_sheet_metadata() — a real API request against
    the read quota.  Tests count those accesses (`metadata_fetches`) to prove the
    sync isn't re-fetching metadata it already holds.
    """

    def __init__(self, ws: FakeWorksheet, tabs: dict | None = None):
        self._ws = ws
        self._tabs = tabs or {}
        self.shares: list[tuple] = []
        self.metadata_fetches = 0

    @property
    def sheet1(self) -> FakeWorksheet:
        self.metadata_fetches += 1
        return self._ws

    def worksheet(self, title: str) -> FakeWorksheet:
        self.metadata_fetches += 1
        if title not in self._tabs:
            raise KeyError(f"no worksheet named {title!r}")
        return self._tabs[title]

    def share(self, *a, **kw):
        self.shares.append((a, kw))


class FakeGC:
    def __init__(self, sheets: dict):
        self._sheets = sheets

    def open_by_key(self, key):
        return self._sheets[key]


class FakeSession:
    """Stands in for AuthorizedSession; records Sheets REST calls."""
    def __init__(self, get_json=None):
        self.posts: list[tuple[str, dict]] = []
        self.gets:  list[tuple[str, dict]] = []
        self._get_json = get_json if get_json is not None else {}

    def post(self, url, params=None, json=None, **kw):
        self.posts.append((url, json))
        return FakeResponse(200, {})

    def get(self, url, params=None, **kw):
        self.gets.append((url, params))
        return FakeResponse(200, self._get_json)

    def delete(self, url, **kw):
        return FakeResponse(204, {})


def cf_requests(session: FakeSession) -> list[dict]:
    """Flatten all batchUpdate requests posted through a FakeSession."""
    out = []
    for _url, body in session.posts:
        out.extend(body.get("requests", []))
    return out

"""
One-time import: load Credit Control's existing data.json into the
credit_control_store Supabase table created by
migrations/credit_control_schema.sql.

Safe to re-run — upserts the single row (id=1), so re-running after fixing a
problem just overwrites it again rather than erroring on a duplicate key. A
backup row is still written first, same as a normal save would.

Usage (from the dashboard repo root, with SUPABASE_URL / SUPABASE_SERVICE_ROLE_KEY
already set in the environment — e.g. via the same .env-loading one-liner
serve uses in .claude/launch.json):

    python scripts/import_credit_control_data.py "C:\\path\\to\\data.json"

Then verify: GET /api/credit-control/data on the running app should return
byte-identical JSON to the source file (the script does this check itself,
re-fetching after the write).
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

import requests

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from app.utils.supabase import SUPABASE_URL, sb_service_headers as sb_headers, sb_service_configured as sb_configured  # noqa: E402


def main() -> int:
    if not sb_configured():
        print("ERROR: SUPABASE_URL / SUPABASE_SERVICE_ROLE_KEY are not set in the environment.")
        return 1

    if len(sys.argv) != 2:
        print("Usage: python scripts/import_credit_control_data.py <path-to-data.json>")
        return 1

    src = Path(sys.argv[1])
    if not src.is_file():
        print(f"ERROR: {src} does not exist.")
        return 1

    raw = src.read_text(encoding="utf-8")
    data = json.loads(raw)  # fails loudly on invalid JSON, same as serve.ps1's /save does
    for key in ("settings", "clients", "invoices"):
        if key not in data:
            print(f"ERROR: source JSON is missing top-level key '{key}' — refusing to import "
                  f"something that doesn't look like a Credit Control data.json.")
            return 1
    print(f"Loaded {src}: {len(data['clients'])} clients, {len(data['invoices'])} invoices.")

    # Write a backup row first, same as every real save will — but only if a
    # row already exists (first-ever run has nothing to back up yet).
    existing = requests.get(
        f"{SUPABASE_URL}/rest/v1/credit_control_store",
        headers=sb_headers(),
        params={"select": "data", "id": "eq.1"},
        timeout=30,
    )
    existing.raise_for_status()
    rows = existing.json()
    if rows:
        requests.post(
            f"{SUPABASE_URL}/rest/v1/credit_control_backups",
            headers=sb_headers("return=minimal"),
            json={"data": rows[0]["data"], "saved_by": "import_script (pre-import snapshot)"},
            timeout=30,
        ).raise_for_status()
        print("Backed up existing row before overwriting.")

    # Upsert the single row.
    r = requests.post(
        f"{SUPABASE_URL}/rest/v1/credit_control_store",
        headers=sb_headers("resolution=merge-duplicates,return=minimal"),
        params={"on_conflict": "id"},
        json={"id": 1, "data": data, "updated_by": "import_script"},
        timeout=30,
    )
    r.raise_for_status()
    print("Imported into credit_control_store.")

    # Round-trip verification: re-fetch and diff against the source.
    check = requests.get(
        f"{SUPABASE_URL}/rest/v1/credit_control_store",
        headers=sb_headers(),
        params={"select": "data", "id": "eq.1"},
        timeout=30,
    )
    check.raise_for_status()
    stored = check.json()[0]["data"]
    if stored == data:
        print("Round-trip check PASSED — stored data is byte-identical to the source file.")
        return 0
    else:
        print("Round-trip check FAILED — stored data differs from the source file. "
              "Do not treat Supabase as the source of truth yet; investigate before proceeding.")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())

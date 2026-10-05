"""
One-off data fix: for every PAID invoice due in a previous month, set its
"Paid on" date to its due date.

Does nothing without --apply (dry run: prints exactly what would change).
With --apply it backs up the live row to credit_control_backups first, only
writes if nobody has saved in the meantime, changes ONLY paidDate on the
listed invoices, and re-reads the result to check nothing else moved.

Usage (from the dashboard repo root, with SUPABASE_URL and
SUPABASE_SERVICE_ROLE_KEY set in the environment):

    python scripts/set_paid_dates_to_due.py            # dry run against Supabase
    python scripts/set_paid_dates_to_due.py --apply    # make the change

    python scripts/set_paid_dates_to_due.py --file C:\\path\\to\\data.json
        # dry run against a local file instead (no Supabase needed)
"""
from __future__ import annotations

import argparse
import copy
import json
import sys
from datetime import date, datetime, timezone
from pathlib import Path

import requests

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from app.utils.supabase import SUPABASE_URL, sb_service_headers as sb_headers, sb_service_configured as sb_configured  # noqa: E402


def find_changes(data: dict, current_month: str) -> list[dict]:
    names = {c["id"]: c.get("companyName", c["id"]) for c in data.get("clients", [])}
    out = []
    for inv in data.get("invoices", []):
        due, paid_on = inv.get("dueDate"), inv.get("paidDate")
        if not (inv.get("paid") and due and paid_on):
            continue
        if due[:7] >= current_month:  # current/future months are left alone
            continue
        if paid_on != due:
            out.append({"id": inv["id"], "client": names.get(inv["clientId"], inv["clientId"]),
                        "due": due, "paid_on": paid_on})
    return out


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--apply", action="store_true")
    ap.add_argument("--file")
    args = ap.parse_args()
    current_month = date.today().strftime("%Y-%m")

    if args.file:
        if args.apply:
            print("--file is dry-run only.")
            return 1
        data = json.loads(Path(args.file).read_text(encoding="utf-8"))
        version = None
    else:
        if not sb_configured():
            print("ERROR: SUPABASE_URL / SUPABASE_SERVICE_ROLE_KEY are not set in the environment.")
            return 1
        r = requests.get(f"{SUPABASE_URL}/rest/v1/credit_control_store", headers=sb_headers(),
                         params={"select": "data,updated_at", "id": "eq.1"}, timeout=30)
        r.raise_for_status()
        rows = r.json()
        if not rows:
            print("ERROR: no data in credit_control_store.")
            return 1
        data, version = rows[0]["data"], rows[0]["updated_at"]

    changes = find_changes(data, current_month)
    paid_prev = sum(1 for i in data["invoices"] if i.get("paid") and (i.get("dueDate") or "9")[:7] < current_month)
    print(f"{paid_prev} paid invoices due before {current_month}; {len(changes)} have a paid date different from the due date:\n")
    for c in sorted(changes, key=lambda c: (c["due"], c["client"])):
        print(f"  {c['id']:<16} {c['client']:<28} due {c['due']}  paid on {c['paid_on']}  ->  {c['due']}")

    if not changes:
        print("\nNothing to change.")
        return 0
    if not args.apply:
        print("\nDRY RUN — nothing written. Re-run with --apply to make this change.")
        return 0

    fixed = copy.deepcopy(data)
    ids = {c["id"] for c in changes}
    for inv in fixed["invoices"]:
        if inv["id"] in ids:
            inv["paidDate"] = inv["dueDate"]

    requests.post(f"{SUPABASE_URL}/rest/v1/credit_control_backups", headers=sb_headers("return=minimal"),
                  json={"data": data, "saved_by": "set_paid_dates_to_due (pre-change snapshot)"},
                  timeout=30).raise_for_status()
    print("\nBacked up the current data to credit_control_backups.")

    now = datetime.now(timezone.utc).isoformat()
    w = requests.patch(f"{SUPABASE_URL}/rest/v1/credit_control_store",
                       headers=sb_headers("return=representation"),
                       params={"id": "eq.1", "updated_at": f"eq.{version}"},
                       json={"data": fixed, "updated_at": now, "updated_by": "set_paid_dates_to_due"},
                       timeout=30)
    w.raise_for_status()
    if not w.json():
        print("ABORTED: someone saved between reading and writing. Nothing was changed. Re-run.")
        return 1

    chk = requests.get(f"{SUPABASE_URL}/rest/v1/credit_control_store", headers=sb_headers(),
                       params={"select": "data", "id": "eq.1"}, timeout=30).json()[0]["data"]
    ok = (chk == fixed
          and len(chk["invoices"]) == len(data["invoices"]) and len(chk["clients"]) == len(data["clients"])
          and chk["settings"] == data["settings"] and chk["clients"] == data["clients"])
    print("Verification PASSED — only the listed paid dates changed." if ok
          else "Verification FAILED — the stored data is not what was expected. Stop and check the backup.")
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())

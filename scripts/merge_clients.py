"""
Merge duplicate Credit Control clients into one, in the live data in Supabase.

Edit GROUPS below: each group lists the existing client names to combine and the name the single
merged client should have. For every group:

  * the client that started earliest is kept (its ID survives); the others are removed and all of
    their invoices are moved onto it
  * if invoices from DIFFERENT source clients fall due in the same month they become ONE invoice
    whose amount is the total of them (amount landed too, once all of them are paid). Several invoices
    that already sat under the same source client in one month are left exactly as they were
  * invoices are renumbered 1, 2, 3 ... in date order across the merged client
  * invoices that followed their old client's account manager are frozen to that account manager, so
    the history doesn't jump to the merged client's current one
  * the client record keeps the earliest start month and original lead-gen / converted-by; takes fees,
    schedule, account manager and industry tags from the most recent member; is churned only if every
    member is churned (churn month = the latest); manual LTV/tenure adjustments are added together;
    and its notes record what it was merged from

Does nothing without --apply: the dry run prints exactly what would change. With --apply it backs
up the live row to credit_control_backups first, only writes if nobody has saved in the meantime, and
re-reads the result to check that nothing but the merged clients' records and invoices moved, that the
total invoiced is unchanged, and that every other client and invoice is untouched.

Usage (from the dashboard repo root, with SUPABASE_URL and SUPABASE_SERVICE_ROLE_KEY set):

    python scripts/merge_clients.py
    python scripts/merge_clients.py --apply

    python scripts/merge_clients.py --data-file C:\\path\\to\\data.json
        # dry run against a local data.json instead (no Supabase needed)
"""
from __future__ import annotations

import argparse
import copy
import json
import re
import sys
from collections import defaultdict
from datetime import datetime, timezone
from pathlib import Path

import requests

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from app.utils.supabase import SUPABASE_URL, sb_service_headers as sb_headers, sb_service_configured as sb_configured  # noqa: E402

GROUPS = [
    {"name": "Grantify", "members": ["Grantify", "Grantify - Grant Funding", "Grantify - UK Govt. Contracts Campaign"]},
    {"name": "Film Farmers", "members": ["Film Farmers", "The Film Farmers"]},
]

EARLIEST_KEYS = ("contractTerms", "agreedTrialMonths", "leadGenBy", "convertedBy")
FEE_KEYS = ("feeGBP", "feeUSD", "feeEUR")
HANDLED = {"id", "companyName", "startMonth", "status", "churnMonth", "notes", "ltvAdjust", "tenureAdjust", *EARLIEST_KEYS, *FEE_KEYS}


def clean(v) -> str:
    return re.sub(r"\s+", " ", str(v)).strip() if v is not None else ""


def norm(v) -> str:
    return clean(v).casefold()


def num(v) -> float:
    return float(v) if isinstance(v, (int, float)) else 0.0


def blank(v) -> bool:
    return v is None or v == ""


def money(v) -> str:
    return f"£{(v or 0):,.2f}"


def find_members(data: dict, names: list[str]) -> list[dict]:
    found = []
    for n in names:
        hits = [c for c in data["clients"] if norm(c["companyName"]) == norm(n)]
        if len(hits) != 1:
            raise SystemExit(f'ERROR: expected exactly one client called "{n}" but found {len(hits)}. Check the name in GROUPS.')
        found.append(hits[0])
    if len({c["id"] for c in found}) != len(found):
        raise SystemExit("ERROR: the same client is listed twice in one group.")
    return sorted(found, key=lambda c: (c.get("startMonth") or "9999-99", c["id"]))


def merged_client(members: list[dict], target: str, report: list[str]) -> dict:
    first, last = members[0], members[-1]
    keys = []
    for c in members:
        for k in c:
            if k not in keys:
                keys.append(k)
    out = dict(first)

    def latest(key):
        for c in reversed(members):
            if not blank(c.get(key)):
                return c.get(key)
        return first.get(key)

    def earliest(key):
        for c in members:
            if not blank(c.get(key)):
                return c.get(key)
        return first.get(key)

    for k in keys:
        if k not in HANDLED:
            out[k] = latest(k)
    for k in EARLIEST_KEYS:
        if k in keys:
            out[k] = earliest(k)
    fee_src = next((c for c in reversed(members) if any(c.get(k) for k in FEE_KEYS)), last)
    for k in FEE_KEYS:
        if k in keys:
            out[k] = fee_src.get(k)
    for k in ("accountManager", "industry", "country", "companySize"):
        if not blank(first.get(k)):
            out[k] = first[k]          # the surviving client keeps its own mapping; only blanks are filled from the others
    out["companyName"] = target
    out["startMonth"] = min((c["startMonth"] for c in members if c.get("startMonth")), default=first.get("startMonth"))
    if all(c.get("status") == "churned" for c in members):
        out["status"] = "churned"
        out["churnMonth"] = max((c["churnMonth"] for c in members if c.get("churnMonth")), default="")
    else:
        live = [c for c in members if c.get("status") != "churned"]
        out["status"] = live[-1].get("status")
        out["churnMonth"] = ""
    for k in ("ltvAdjust", "tenureAdjust"):
        vals = [c[k] for c in members if isinstance(c.get(k), (int, float))]
        if vals:
            out[k] = sum(vals)
            if len([v for v in vals if v]) > 1:
                report.append(f"   note: {k} values {vals} were added together -> {out[k]}")
    stamp = datetime.now().strftime("%Y-%m-%d")
    notes = [n for n in (clean(c.get("notes")) for c in members) if n]
    notes.append(f"Merged on {stamp} from: " + "; ".join(c["companyName"] for c in members) + ".")
    out["notes"] = " | ".join(notes)
    return out


def merge_invoices(group: list[dict], client_names: dict[str, str], report: list[str]) -> dict:
    group = sorted(group, key=lambda i: (i.get("issueDate") or "9999", i["id"]))
    base = dict(group[0])
    base["amountGBP"] = sum(i.get("amountGBP") or 0 for i in group)
    for k in ("amountUSD", "amountEUR"):
        vals = [i[k] for i in group if i.get(k) is not None]
        base[k] = sum(vals) if vals else None
    issue = [i["issueDate"] for i in group if i.get("issueDate")]
    due = [i["dueDate"] for i in group if i.get("dueDate")]
    base["issueDate"] = min(issue) if issue else base.get("issueDate")
    base["dueDate"] = max(due) if due else base.get("dueDate")
    base["issued"] = all(i.get("issued") for i in group)
    paid_all = all(i.get("paid") for i in group)
    base["paid"] = paid_all
    if paid_all:
        pd = [i["paidDate"] for i in group if i.get("paidDate")]
        base["paidDate"] = max(pd) if pd else base.get("paidDate")
        if any(i.get("amountLanded") is not None for i in group):
            base["amountLanded"] = sum(i["amountLanded"] if i.get("amountLanded") is not None else (i.get("amountGBP") or 0) for i in group)
    else:
        unpaid = next(i for i in group if not i.get("paid"))
        base["paidDate"] = unpaid.get("paidDate")
        base.pop("amountLanded", None)
        if any(i.get("paid") for i in group):
            report.append(f"   !! {base['month']}: part-paid month — some of these invoices are paid and some are not; the merged invoice is shown as UNPAID. Check it.")
    for k in ("issuedLate", "uncertain"):
        if any(i.get(k) for i in group):
            base[k] = True
    ams = {clean(i.get("accountManager")) for i in group} - {""}
    if len(ams) > 1:
        biggest = max(group, key=lambda i: i.get("amountGBP") or 0)
        base["accountManager"] = clean(biggest.get("accountManager"))
        report.append(f"   note: {base['month']}: different account managers on the merged invoices ({', '.join(sorted(ams))}); kept the one on the largest ({base['accountManager']}).")
    elif ams:
        base["accountManager"] = next(iter(ams))
    pk = {i.get("package") for i in group if i.get("package")}
    if len(pk) > 1:
        report.append(f"   note: {base['month']}: different packages on the merged invoices ({', '.join(sorted(pk))}); kept the first.")
    notes = [n for n in (clean(i.get("notes")) for i in group) if n]
    notes.append("Merged invoices: " + " + ".join(f"{money(i.get('amountGBP'))} ({client_names[i['clientId']]})" for i in group))
    base["notes"] = " | ".join(notes)
    return base


def merge_group(data: dict, group: dict, report: list[str]) -> None:
    members = find_members(data, group["members"])
    target = group["name"]
    ids = {c["id"] for c in members}
    names = {c["id"]: c["companyName"] for c in members}
    survivor = members[0]
    new_client = merged_client(members, target, report)
    report.append(f'\n=== {target} — {len(members)} clients -> 1 (keeps ID {survivor["id"]}) ===')
    for c in members:
        n = sum(1 for i in data["invoices"] if i["clientId"] == c["id"])
        report.append(f'   {c["id"]}  {c["companyName"][:42]:<42} {c.get("startMonth") or "?"} -> {c.get("churnMonth") or ("(" + str(c.get("status")) + ")")}  AM: {clean(c.get("accountManager")) or "(none)"}  fee {money(c.get("feeGBP"))}  invoices: {n}')
    report.append(f'   result: status {new_client["status"]}, start {new_client["startMonth"]}, churn {new_client.get("churnMonth") or "—"}, '
                  f'AM {clean(new_client.get("accountManager")) or "(none)"}, fee {money(new_client.get("feeGBP"))}, '
                  f'industry {new_client.get("industry") or "—"}, lead gen {new_client.get("leadGenBy") or "—"}, converted by {new_client.get("convertedBy") or "—"}')

    invs = [copy.deepcopy(i) for i in data["invoices"] if i["clientId"] in ids]
    merged_am = clean(new_client.get("accountManager"))
    frozen = 0
    for i in invs:
        if clean(i.get("accountManager")):
            continue
        src = clean(next(c for c in members if c["id"] == i["clientId"]).get("accountManager"))
        if src and src != merged_am:
            i["accountManager"] = src
            frozen += 1
        elif not src and merged_am:
            report.append(f"   note: invoice {names[i['clientId']]} {i['month']} had no account manager and will now follow {merged_am}.")
    if frozen:
        report.append(f"   {frozen} invoice(s) frozen to their old client's account manager so history doesn't move.")

    by_month = defaultdict(list)
    for i in invs:
        by_month[i["month"]].append(i)
    final: list[dict] = []
    replaced: dict[str, dict] = {}
    removed: set[str] = set()
    for month in sorted(by_month):
        grp = by_month[month]
        if len({i["clientId"] for i in grp}) >= 2:
            m = merge_invoices(grp, names, report)
            report.append(f'   MERGED {month}: ' + " + ".join(f'{money(i.get("amountGBP"))} ({names[i["clientId"]]})' for i in grp) + f' = {money(m["amountGBP"])}')
            replaced[m["id"]] = m
            removed |= {i["id"] for i in grp if i["id"] != m["id"]}
            final.append(m)
        else:
            if len(grp) > 1:
                report.append(f"   (kept as-is: {len(grp)} separate invoices already in {month} under {names[grp[0]['clientId']]} — {', '.join(money(i.get('amountGBP')) for i in grp)})")
            final.extend(grp)
    final.sort(key=lambda i: (i["month"], num(i.get("monthNumber")), i.get("issueDate") or "", i["id"]))
    renum = 0
    for n, i in enumerate(final, 1):
        if i.get("monthNumber") != n:
            renum += 1
        i["monthNumber"] = n
        i["clientId"] = survivor["id"]
        replaced[i["id"]] = i
    if not any(len({i["clientId"] for i in g}) >= 2 for g in by_month.values()):
        report.append("   No month had invoices from more than one of these clients, so no invoices needed adding together.")
    report.append(f"   {len(invs)} invoices -> {len(final)}; {renum} renumbered. Timeline (month · no. · amount · paid):")
    line = []
    for i in final:
        line.append(f'{i["month"]} #{i["monthNumber"]} {money(i.get("amountGBP"))}{"" if i.get("paid") else " UNPAID"}')
    for k in range(0, len(line), 3):
        report.append("      " + "   |   ".join(line[k:k + 3]))

    out_inv = []
    for i in data["invoices"]:
        if i["id"] in removed:
            continue
        out_inv.append(replaced.get(i["id"], i))
    data["invoices"] = out_inv
    data["clients"] = [new_client if c["id"] == survivor["id"] else c for c in data["clients"] if c["id"] == survivor["id"] or c["id"] not in ids]

    blob = json.dumps(data["settings"], ensure_ascii=False)
    for c in members:
        if c["id"] in blob or json.dumps(c["companyName"], ensure_ascii=False)[1:-1] in blob:
            report.append(f'   !! "{c["companyName"]}" / {c["id"]} is mentioned in Settings — check it by hand.')


def total(data: dict) -> float:
    return round(sum(i.get("amountGBP") or 0 for i in data["invoices"]), 2)


def explain_bad_response(r) -> None:
    """Supabase answered, but not with data — say what it did answer, so the cause (almost always the URL or key) is obvious."""
    host = re.sub(r"^https?://", "", SUPABASE_URL).split("/")[0] or "(empty)"
    body = (r.text or "").strip().replace("\n", " ")[:200]
    print("ERROR: the database did not answer with data.")
    print(f"  address used : {host}   (status {r.status_code}, type {r.headers.get('content-type', '?')})")
    print(f"  it replied   : {body or '(nothing)'}")
    print("  Check SUPABASE_URL is your Supabase PROJECT address — it looks like https://xxxxxxxx.supabase.co")
    print("  (not the dashboard's own address, and with nothing after .supabase.co), and that the key is the service_role key.")


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--apply", action="store_true")
    ap.add_argument("--data-file")
    args = ap.parse_args()

    version = None
    if args.data_file:
        if args.apply:
            print("--data-file is dry-run only.")
            return 1
        data = json.loads(Path(args.data_file).read_text(encoding="utf-8"))
    else:
        if not sb_configured():
            print("ERROR: SUPABASE_URL / SUPABASE_SERVICE_ROLE_KEY are not set in the environment.")
            return 1
        r = requests.get(f"{SUPABASE_URL}/rest/v1/credit_control_store", headers=sb_headers(),
                         params={"select": "data,updated_at", "id": "eq.1"}, timeout=30)
        try:
            r.raise_for_status()
            rows = r.json()
        except (ValueError, requests.HTTPError):
            explain_bad_response(r)
            return 1
        if not rows:
            print("ERROR: no data in credit_control_store.")
            return 1
        data, version = rows[0]["data"], rows[0]["updated_at"]

    fixed = copy.deepcopy(data)
    report: list[str] = []
    for g in GROUPS:
        merge_group(fixed, g, report)
    print("\n".join(report))

    n_before, n_after = len(data["clients"]), len(fixed["clients"])
    i_before, i_after = len(data["invoices"]), len(fixed["invoices"])
    t_before, t_after = total(data), total(fixed)
    print(f"\nClients {n_before} -> {n_after}.  Invoices {i_before} -> {i_after}.  Total invoiced {money(t_before)} -> {money(t_after)}.")
    if t_before != t_after:
        print("!! The total invoiced changed — this should never happen. Stopping.")
        return 1
    touched = {c["id"] for g in GROUPS for c in find_members(data, g["members"])}
    if not args.apply:
        print("\nDRY RUN — nothing written. Re-run with --apply to make these changes.")
        return 0

    requests.post(f"{SUPABASE_URL}/rest/v1/credit_control_backups", headers=sb_headers("return=minimal"),
                  json={"data": data, "saved_by": "merge_clients (pre-change snapshot)"}, timeout=30).raise_for_status()
    print("\nBacked up the current data to credit_control_backups.")
    w = requests.patch(f"{SUPABASE_URL}/rest/v1/credit_control_store", headers=sb_headers("return=representation"),
                       params={"id": "eq.1", "updated_at": f"eq.{version}"},
                       json={"data": fixed, "updated_at": datetime.now(timezone.utc).isoformat(), "updated_by": "merge_clients"}, timeout=30)
    w.raise_for_status()
    if not w.json():
        print("ABORTED: someone saved between reading and writing. Nothing was changed. Re-run.")
        return 1
    chk = requests.get(f"{SUPABASE_URL}/rest/v1/credit_control_store", headers=sb_headers(),
                       params={"select": "data", "id": "eq.1"}, timeout=30).json()[0]["data"]
    survivors = {find_members(data, g["members"])[0]["id"] for g in GROUPS}
    others_same = ({c["id"]: c for c in chk["clients"] if c["id"] not in touched} == {c["id"]: c for c in data["clients"] if c["id"] not in touched}
                   and {i["id"]: i for i in chk["invoices"] if i["clientId"] not in survivors} ==
                       {i["id"]: i for i in data["invoices"] if i["clientId"] not in touched})
    ok = chk == fixed and chk["settings"] == data["settings"] and total(chk) == t_before and others_same
    print("Verification PASSED — only the merged clients and their invoices changed; the total invoiced is the same; every other client, invoice and setting is untouched."
          if ok else "Verification FAILED — stored data is not what was expected. Stop and check the backup.")
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())

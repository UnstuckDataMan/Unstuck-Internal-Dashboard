"""
Reset the live Credit Control data (Supabase) so the history matches the old Credit Control Google Sheet.

Everything BEFORE --keep-from (default 2026-07, i.e. up to and including June 2026) is rebuilt from the sheet.
July 2026 onwards — July, August, September, October, November and anything later — is never touched.

What it does, in order:
  1. Grantify and Film Farmers are merged into one client each (the same merge as merge_clients.py), if they aren't already,
     so the sheet's several Grantify / Film Farmers names all land on one client.
  2. Every invoice in the app dated before --keep-from is removed and replaced by the sheet's rows for those months.
     A sheet row counts as an invoice when its £ cell is a number (or a number with a stray letter, like "A2430" or
     "1090A"). Rows where the £ cell says something else — "THEY HAVE LEFT", "#2840" (paused), "1090(REFUNDED)",
     "RENEGOTIATE" — are not invoices, and are listed. Different sheet names for the same client in the same month
     (the Grantify and Film Farmers variants) become one invoice with the amounts added together.
  3. Clients are matched to the sheet by name (ignoring case and a trailing bracket). A client only in the sheet is
     created. A client whose only invoices are old gets its fee / package / terms / start / churn month refreshed from
     the sheet; a client that is invoiced from --keep-from onwards keeps its live settings (only the start month is
     recalculated). A client left with no invoices at all is listed, not deleted (add --delete-empty-clients to remove them).
  4. THE CLIENT MAPPING IS KEPT: account manager (on the client and on each invoice), industry / country / company size,
     manual LTV and tenure adjustments, lead-gen / converted-by, targets and every setting. Re-created invoices carry over
     the account manager, notes, paid date, landed amount, uncertain flag and package of the app invoice they replace
     (matched by client + month).
  5. Invoices are numbered 1, 2, 3 … per client in date order. Invoices from --keep-from onwards keep their number; any
     that no longer follow on from the rebuilt history are counted in the report (add --renumber-kept to fix those too).

Does nothing without --apply: the dry run prints what would change month by month and writes a full line-by-line CSV
(--report). With --apply it backs up the live row to credit_control_backups first, only writes if nobody has saved in the
meantime, and re-reads the result to check that everything from --keep-from onwards, the client mapping and the settings
are exactly as they were.

Usage (from the dashboard repo root, with SUPABASE_URL and SUPABASE_SERVICE_ROLE_KEY set):

    python scripts/reset_from_sheet.py "C:\\path\\Credit Control (1).xlsx" --report "C:\\path\\reset-report.csv"
    python scripts/reset_from_sheet.py "C:\\path\\Credit Control (1).xlsx" --apply

    python scripts/reset_from_sheet.py "C:\\path\\sheet.xlsx" --data-file C:\\path\\data.json
        # dry run against a local data.json instead (no Supabase needed)
"""
from __future__ import annotations

import argparse
import copy
import csv
import datetime
import difflib
import json
import re
import sys
from collections import defaultdict
from pathlib import Path

import requests

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent))
sys.path.insert(0, str(HERE))
from app.utils.supabase import SUPABASE_URL, sb_service_headers as sb_headers, sb_service_configured as sb_configured  # noqa: E402
import merge_clients as mc  # noqa: E402
import sheet_rows as sr  # noqa: E402

GROUP_TARGETS = {"grantify": "Grantify", "film farmers": "Film Farmers"}
MAPPING_FIELDS = ("accountManager", "industry", "country", "companySize", "ltvAdjust", "tenureAdjust", "leadGenBy", "convertedBy")
CARRY_FIELDS = ("accountManager", "uncertain", "issuedLate", "package", "packageOverride")
AMT_FLAG = re.compile(r"^\s*[A-Za-z]?\s*[£$€]?\s*\d[\d,]*(?:\.\d+)?\s*[A-Za-z]?\s*$")
AMT_NOTE = re.compile(r"^\s*\(.*\)\s*\d[\d,]*(?:\.\d+)?\s*$")


def money(v) -> str:
    return f"£{(v or 0):,.2f}"


def add_month(mk: str, n: int = 1) -> str:
    y, m = int(mk[:4]), int(mk[5:7]) + n
    y += (m - 1) // 12
    m = (m - 1) % 12 + 1
    return f"{y}-{m:02d}"


def canon(key: str) -> str:
    if key.startswith("grantify"):
        return "grantify"
    if key in ("film farmers", "the film farmers"):
        return "film farmers"
    return key


def app_key(c: dict) -> str:
    return canon(sr.norm_key(c["companyName"]))


def not_an_invoice(r: dict):
    """None if the sheet row is an invoice; otherwise why it isn't."""
    raw = r["amt_raw"]
    if not r["issue"] and not r["due"] and r["email"] and "@" not in r["email"]:
        return "looks like a note, not an invoice (no dates, and the email cell is text)"
    if isinstance(raw, (int, float)):
        return None
    if raw is None:
        return "no £ typed"
    s = str(raw).strip()
    if AMT_FLAG.match(s) or AMT_NOTE.match(s):
        return None
    return f'£ cell reads "{s}"'


def merge_same_month(rows: list[dict]) -> dict:
    """Several sheet names for one client in one month -> one invoice, amounts added."""
    rows = sorted(rows, key=lambda r: (r["issue"] or "9999", r["row"]))
    out = dict(rows[0])
    out["amount"] = round(sum(r["amount"] for r in rows), 2)
    for k in ("usd", "eur"):
        vals = [r[k] for r in rows if r[k] is not None]
        out[k] = round(sum(vals), 2) if vals else None
    issues = [r["issue"] for r in rows if r["issue"]]
    dues = [r["due"] for r in rows if r["due"]]
    out["issue"] = min(issues) if issues else None
    out["due"] = max(dues) if dues else None
    paid = [True if r["paid"] is None else r["paid"] for r in rows]
    out["paid"] = all(paid)
    out["issued"] = all(r["issued"] for r in rows)
    out["merged"] = " + ".join(f"{money(r['amount'])} ({r['company']})" for r in rows)
    return out


def plan(data: dict, rows: list[dict], keep_from: str, aliases: dict, renumber_kept: bool, delete_empty: bool) -> dict:
    fixed = copy.deepcopy(data)
    log: list[str] = []
    problems: list[str] = []
    merged_ids: set[str] = set()

    # 1 ── merge the duplicate clients first (same merge as merge_clients.py), so the sheet's variants share one client
    for gk, target in GROUP_TARGETS.items():
        members = [c["companyName"] for c in fixed["clients"] if app_key(c) == gk]
        if len(members) > 1:
            lines: list[str] = []
            mc.merge_group(fixed, {"name": target, "members": members}, lines)
            merged_ids.add(next(c["id"] for c in fixed["clients"] if c["companyName"] == target))
            log.append(f"Merged {len(members)} clients into {target}: " + "; ".join(members))

    idx: dict[str, dict] = {}
    for c in fixed["clients"]:
        k = app_key(c)
        if k in idx:
            problems.append(f'two clients in the app share the name "{c["companyName"]}" / "{idx[k]["companyName"]}"')
        idx[k] = c

    # 2 ── which sheet rows are invoices, and which app client each belongs to
    alias_by_key = {sr.norm_key(k): sr.norm_key(v) for k, v in aliases.items()}
    skipped: list[tuple[dict, str]] = []
    reset_rows = [r for r in rows if r["month"] < keep_from]
    usable: list[dict] = []
    for r in reset_rows:
        why = not_an_invoice(r)
        if why:
            skipped.append((r, why))
        else:
            usable.append(r)
    for r in usable:
        r["ck"] = canon(alias_by_key.get(r["key"], r["key"]))

    grouped: dict[tuple, list[dict]] = defaultdict(list)
    for r in usable:
        grouped[(r["ck"], r["month"])].append(r)
    sheet_inv: dict[str, dict[str, list[dict]]] = defaultdict(lambda: defaultdict(list))   # ck -> month -> rows
    merged_rows = 0
    for (ck, month), rs in grouped.items():
        if ck in GROUP_TARGETS and len({r["key"] for r in rs}) > 1:
            sheet_inv[ck][month].append(merge_same_month(rs))
            merged_rows += len(rs) - 1
        else:
            sheet_inv[ck][month].extend(sorted(rs, key=lambda r: (r["issue"] or "9999", r["row"])))

    # 3 ── clients that only the sheet knows about
    nums = [int(c["id"][1:]) for c in fixed["clients"] if re.fullmatch(r"c\d+", c["id"])]
    next_c = max(nums or [0]) + 1
    new_clients: list[dict] = []
    for ck in sorted(sheet_inv):
        if ck in idx:
            continue
        latest_month = max(sheet_inv[ck])
        latest = sheet_inv[ck][latest_month][-1]
        first_month = min(sheet_inv[ck])
        issue_day, due_days = 1, 7
        for m in sorted(sheet_inv[ck], reverse=True):
            done = False
            for r in reversed(sheet_inv[ck][m]):
                if r["issue"]:
                    issue_day = int(r["issue"][8:10])
                    if r["due"]:
                        gap = (datetime.date.fromisoformat(r["due"]) - datetime.date.fromisoformat(r["issue"])).days
                        if 0 <= gap <= 60:
                            due_days = gap
                    done = True
                    break
            if done:
                break
        c = {
            "id": f"c{next_c:03d}", "companyName": GROUP_TARGETS.get(ck, latest["company"]), "qbName": latest["qb"], "email": latest["email"],
            "accountManager": latest["am"], "package": latest["pkg"], "feeGBP": latest["amount"], "feeUSD": latest["usd"], "feeEUR": latest["eur"],
            "issueDay": issue_day, "dueDays": due_days, "contractTerms": latest["terms"], "agreedTrialMonths": latest["agreed"],
            "startMonth": first_month, "status": "churned", "churnMonth": add_month(latest_month),
            "leadGenBy": "", "convertedBy": "", "autoIssue": False, "mailMerge": False, "notes": "",
        }
        next_c += 1
        fixed["clients"].append(c)
        idx[ck] = c
        new_clients.append(c)

    # 4 ── take out every app invoice before keep_from, remembering them to carry details across
    old_by: dict[tuple, list[dict]] = defaultdict(list)
    kept: list[dict] = []
    removed_total = 0
    for i in fixed["invoices"]:
        if i["month"] < keep_from:
            old_by[(i["clientId"], i["month"])].append(i)
            removed_total += 1
        else:
            kept.append(i)
    sort_old = lambda l: sorted(l, key=lambda i: (i.get("issueDate") or "", i.get("dueDate") or "", i.get("amountGBP") or 0))
    for k in old_by:
        old_by[k] = sort_old(old_by[k])

    # 5 ── build the new invoices
    inum = [int(i["id"][1:]) for i in fixed["invoices"] if re.fullmatch(r"i\d+", i["id"])]
    next_i = max(inum or [0]) + 1
    new_invoices: list[dict] = []
    diff: list[dict] = []
    for ck, months in sheet_inv.items():
        c = idx[ck]
        issue_day = c.get("issueDay") or 1
        due_days = c.get("dueDays") if c.get("dueDays") is not None else 7
        for month, rs in months.items():
            olds = old_by.get((c["id"], month), [])
            for n, r in enumerate(rs):
                old = olds[n] if n < len(olds) else None
                issue = r["issue"] or f"{month}-{min(issue_day, 28):02d}"
                due = r["due"] or (datetime.date.fromisoformat(issue) + datetime.timedelta(days=due_days)).isoformat()
                paid = True if r["paid"] is None else bool(r["paid"])
                issued = True if paid else bool(r["issued"])
                inv = {
                    "id": old["id"] if old else f"i{next_i:04d}", "clientId": c["id"], "month": month, "monthNumber": None,
                    "issueDate": issue, "dueDate": due, "amountGBP": r["amount"], "amountUSD": r["usd"], "amountEUR": r["eur"],
                    "issued": issued, "paid": paid, "paidDate": due if paid else None, "notes": "",
                }
                if not old:
                    next_i += 1
                else:
                    for f in CARRY_FIELDS:
                        if f in old and old[f] not in (None, ""):
                            if f in ("uncertain",) and paid:
                                continue
                            inv[f] = old[f]
                    if old.get("notes"):
                        inv["notes"] = old["notes"]
                    if paid and old.get("paid") and old.get("paidDate"):
                        inv["paidDate"] = old["paidDate"]
                    if paid and old.get("paid") and "amountLanded" in old and old.get("amountGBP") == inv["amountGBP"]:
                        inv["amountLanded"] = old["amountLanded"]
                new_invoices.append(inv)
                what = []
                if old:
                    for label, a, b in (("£", old.get("amountGBP"), inv["amountGBP"]), ("issue date", old.get("issueDate"), inv["issueDate"]),
                                        ("due date", old.get("dueDate"), inv["dueDate"]), ("issued", bool(old.get("issued")), inv["issued"]),
                                        ("paid", bool(old.get("paid")), inv["paid"])):
                        if a != b:
                            what.append(label)
                diff.append({"month": month, "action": "ADDED" if not old else ("CHANGED" if what else "same"), "client": c["companyName"],
                             "what": ", ".join(what), "old": old, "new": inv, "merged": r.get("merged", "")})
            for old in olds[len(rs):]:
                diff.append({"month": month, "action": "REMOVED", "client": c["companyName"], "what": "", "old": old, "new": None, "merged": ""})
    for (cid, month), olds in old_by.items():
        c = next((x for x in fixed["clients"] if x["id"] == cid), None)
        ck = app_key(c) if c else None
        if c and ck in sheet_inv and month in sheet_inv[ck]:
            continue
        for old in olds:
            diff.append({"month": month, "action": "REMOVED", "client": c["companyName"] if c else cid, "what": "month not in the sheet for this client",
                         "old": old, "new": None, "merged": ""})

    fixed["invoices"] = kept + new_invoices

    # 6 ── month numbers: 1, 2, 3 … per client in date order (invoices from keep_from on keep theirs unless --renumber-kept)
    by_client: dict[str, list[dict]] = defaultdict(list)
    for i in fixed["invoices"]:
        by_client[i["clientId"]].append(i)
    boundary: list[str] = []
    cname = {c["id"]: c["companyName"] for c in fixed["clients"]}
    for cid, invs in by_client.items():
        invs.sort(key=lambda i: (i["month"], i.get("issueDate") or "", i.get("amountGBP") or 0))
        for pos, i in enumerate(invs, 1):
            if i["month"] < keep_from:
                i["monthNumber"] = pos
            elif i.get("monthNumber") != pos:
                if renumber_kept:
                    i["monthNumber"] = pos
                else:
                    boundary.append(f'{cname[cid]} {i["month"]}: invoice is numbered {i.get("monthNumber")}, the rebuilt history makes it {pos}')

    # 7 ── client records
    recent = {i["clientId"] for i in kept}
    changed_clients: list[str] = []
    empty: list[dict] = []
    flagged: list[str] = []
    for c in fixed["clients"]:
        invs = by_client.get(c["id"], [])
        if not invs:
            empty.append(c)
            continue
        first = min(i["month"] for i in invs)
        before = dict(c)
        if c.get("startMonth") != first:
            c["startMonth"] = first
        ck = app_key(c)
        if c["id"] not in recent and ck in sheet_inv:
            latest_month = max(sheet_inv[ck])
            latest = sheet_inv[ck][latest_month][-1]
            c["feeGBP"], c["feeUSD"], c["feeEUR"] = latest["amount"], latest["usd"], latest["eur"]
            for f, v in (("package", latest["pkg"]), ("contractTerms", latest["terms"])):
                if v:
                    c[f] = v
            if latest["agreed"] is not None:
                c["agreedTrialMonths"] = latest["agreed"]
            for f, v in (("qbName", latest["qb"]), ("email", latest["email"])):
                if v and not c.get(f):
                    c[f] = v
            last = max(i["month"] for i in invs)
            if c.get("status") == "churned":
                c["churnMonth"] = add_month(last)
            else:
                flagged.append(f'{c["companyName"]}: marked active in the app but its last invoice is {last}')
        if c != before:
            changed_clients.append(c["companyName"])
    if delete_empty and empty:
        gone = {c["id"] for c in empty}
        fixed["clients"] = [c for c in fixed["clients"] if c["id"] not in gone]

    # similar-name hints for sheet names that became new clients
    hints = {}
    app_names = [c["companyName"] for c in data["clients"]]
    for c in new_clients:
        n = sr.norm_key(c["companyName"])
        close = [a for a in app_names if sr.norm_key(a) and sr.norm_key(a) != n and (sr.norm_key(a) in n or n in sr.norm_key(a))]
        close += [a for a in difflib.get_close_matches(c["companyName"], app_names, n=3, cutoff=0.72) if a not in close]
        if close:
            hints[c["companyName"]] = close[:3]

    return {"fixed": fixed, "log": log, "problems": problems, "skipped": skipped, "diff": diff, "new_clients": new_clients,
            "empty": empty, "changed_clients": changed_clients, "boundary": boundary, "flagged": flagged, "hints": hints,
            "removed_total": removed_total, "added_total": len(new_invoices), "merged_rows": merged_rows, "reset_rows": reset_rows,
            "merged_ids": merged_ids}


def violations(data: dict, fixed: dict, keep_from: str, renumber_kept: bool, merged_ids=()) -> list[str]:
    """Things the reset promises never to change; returns what (if anything) it would break."""
    bad: list[str] = []
    strip = lambda i: {k: v for k, v in i.items() if k not in ("clientId", "monthNumber")}
    before = {i["id"]: i for i in data["invoices"] if i["month"] >= keep_from}
    after = {i["id"]: i for i in fixed["invoices"] if i["month"] >= keep_from}
    if set(before) != set(after):
        bad.append(f"invoices from {keep_from} on would be added or removed ({len(set(before) ^ set(after))})")
    changed = [k for k in before if k in after and strip(before[k]) != strip(after[k])]
    if changed:
        bad.append(f"{len(changed)} invoice(s) from {keep_from} on would change")
    if not renumber_kept:
        renum = [k for k in before if k in after and before[k].get("monthNumber") != after[k].get("monthNumber")]
        if renum:
            bad.append(f"{len(renum)} invoice(s) from {keep_from} on would be renumbered")
    if fixed["settings"] != data["settings"]:
        bad.append("settings would change")
    now = {c["id"]: c for c in fixed["clients"]}
    lost = []
    for c in data["clients"]:
        if c["id"] not in now:
            continue
        # a merged client's manual LTV / tenure adjustments are legitimately added together
        fields = [f for f in MAPPING_FIELDS if not (c["id"] in merged_ids and f in ("ltvAdjust", "tenureAdjust"))]
        diffs = [f'{f}: "{c.get(f)}" -> "{now[c["id"]].get(f)}"' for f in fields if now[c["id"]].get(f) != c.get(f)]
        if diffs:
            lost.append(f'{c["companyName"]} ({"; ".join(diffs)})')
    if lost:
        bad.append(f"client mapping would change for {len(lost)} client(s): " + " | ".join(lost[:5]))
    return bad


def write_report(path: str, diff: list[dict]) -> None:
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["Month", "Change", "Client", "What differs", "App £ (before)", "Sheet £ (after)", "App issue", "Sheet issue", "App due", "Sheet due",
                    "App issued", "Sheet issued", "App paid", "Sheet paid", "Combined from"])
        for d in sorted(diff, key=lambda d: (d["month"], d["client"].lower(), d["action"])):
            o, n = d["old"] or {}, d["new"] or {}
            w.writerow([d["month"], d["action"], d["client"], d["what"], o.get("amountGBP", ""), n.get("amountGBP", ""),
                        o.get("issueDate", ""), n.get("issueDate", ""), o.get("dueDate", ""), n.get("dueDate", ""),
                        o.get("issued", ""), n.get("issued", ""), o.get("paid", ""), n.get("paid", ""), d["merged"]])


def explain_bad_response(r) -> None:
    """Supabase answered, but not with data — say what it did answer, so the cause (almost always the URL or key) is obvious."""
    host = re.sub(r"^https?://", "", SUPABASE_URL).split("/")[0] or "(empty)"
    body = (r.text or "").strip().replace("\n", " ")[:200]
    print("ERROR: the database did not answer with data.")
    print(f"  address used : {host}   (status {r.status_code}, type {r.headers.get('content-type', '?')})")
    print(f"  it replied   : {body or '(nothing)'}")
    print("  Check SUPABASE_URL is your Supabase PROJECT address — it looks like https://xxxxxxxx.supabase.co")
    print("  (not the dashboard's own address, and with nothing after .supabase.co), and that the key is the service_role key.")


def load_live():
    if not sb_configured():
        print("ERROR: SUPABASE_URL / SUPABASE_SERVICE_ROLE_KEY are not set in the environment.")
        raise SystemExit(1)
    r = requests.get(f"{SUPABASE_URL}/rest/v1/credit_control_store", headers=sb_headers(),
                     params={"select": "data,updated_at", "id": "eq.1"}, timeout=30)
    try:
        r.raise_for_status()
        rows = r.json()
    except ValueError:
        explain_bad_response(r)
        raise SystemExit(1)
    except requests.HTTPError:
        explain_bad_response(r)
        raise SystemExit(1)
    if not rows:
        print("ERROR: no data in credit_control_store.")
        raise SystemExit(1)
    return rows[0]["data"], rows[0]["updated_at"]


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("xlsx")
    ap.add_argument("--apply", action="store_true")
    ap.add_argument("--data-file")
    ap.add_argument("--keep-from", default="2026-07", help="months from here on are never touched (default 2026-07)")
    ap.add_argument("--report", default="reset-report.csv")
    ap.add_argument("--aliases", help="JSON file: {\"name as written in the sheet\": \"client name in the app\"}")
    ap.add_argument("--renumber-kept", action="store_true")
    ap.add_argument("--delete-empty-clients", action="store_true")
    ap.add_argument("--create-similar", action="store_true", help="allow --apply even when a client to be created has a similar name in the app")
    args = ap.parse_args()

    # the name-mapping file is picked up automatically when it sits next to the sheet (reset-aliases.json)
    alias_path = Path(args.aliases) if args.aliases else Path(args.xlsx).with_name("reset-aliases.json")
    aliases = json.loads(alias_path.read_text(encoding="utf-8")) if alias_path.exists() else {}
    if aliases:
        print(f"Using name mappings from {alias_path.name}: " + "; ".join(f'"{k}" = "{v}"' for k, v in aliases.items()))
    rows, warnings, tabs = sr.parse_sheet(args.xlsx)
    version = None
    if args.data_file:
        if args.apply:
            print("--data-file is dry-run only.")
            return 1
        data = json.loads(Path(args.data_file).read_text(encoding="utf-8"))
    else:
        data, version = load_live()

    p = plan(data, rows, args.keep_from, aliases, args.renumber_kept, args.delete_empty_clients)
    fixed, diff = p["fixed"], p["diff"]
    totals = sr.sheet_totals(args.xlsx)

    print(f"Sheet: {len(tabs)} month tabs, {len(rows)} lines. App: {len(data['clients'])} clients, {len(data['invoices'])} invoices.")
    print(f"Months BEFORE {args.keep_from} are rebuilt from the sheet; {args.keep_from} onwards is left exactly as it is.")
    for w in warnings:
        print("  sheet warning:", w)
    for line in p["log"]:
        print("  " + line)
    for pr in p["problems"]:
        print("!! " + pr)

    counts = defaultdict(lambda: defaultdict(int))
    app_sum, new_sum = defaultdict(float), defaultdict(float)
    for d in diff:
        counts[d["month"]][d["action"]] += 1
        if d["old"]:
            app_sum[d["month"]] += d["old"].get("amountGBP") or 0
        if d["new"]:
            new_sum[d["month"]] += d["new"].get("amountGBP") or 0
    print("\nMonth      app £ -> sheet £      added removed changed same   (sheet's own Total, where it has one)")
    for m in sorted(counts):
        c = counts[m]
        if not (c["ADDED"] or c["REMOVED"] or c["CHANGED"]):
            continue
        t = totals.get(m)
        note = ""
        if t is not None:
            note = f"   sheet Total {money(t)}" + ("" if abs(t - new_sum[m]) < 0.5 else f"  (rebuilt {money(new_sum[m])} — see skipped rows)")
        print(f"{m}  {money(app_sum[m]):>11} -> {money(new_sum[m]):>11}   {c['ADDED']:>5} {c['REMOVED']:>7} {c['CHANGED']:>7} {c['same']:>4}{note}")
    same_months = [m for m in sorted(counts) if not (counts[m]["ADDED"] or counts[m]["REMOVED"] or counts[m]["CHANGED"])]
    print(f"({len(same_months)} months already match the sheet exactly)")
    tot = {a: sum(c[a] for c in counts.values()) for a in ("ADDED", "REMOVED", "CHANGED", "same")}
    print(f"\nInvoices: {p['removed_total']} removed from the app before {args.keep_from}, {p['added_total']} put back from the sheet "
          f"({tot['same']} identical, {tot['CHANGED']} changed, {tot['ADDED']} new, {tot['REMOVED']} gone for good).")
    if p["merged_rows"]:
        print(f"{p['merged_rows']} sheet row(s) for the same Grantify / Film Farmers month were added together into one invoice.")

    if p["skipped"]:
        print(f"\n{len(p['skipped'])} sheet row(s) are NOT counted as invoices:")
        for r, why in p["skipped"]:
            print(f"   {r['month']}  {r['company'][:34]:<34} {why}")
    if p["new_clients"]:
        print(f"\n{len(p['new_clients'])} client(s) are in the sheet but not the app — they will be CREATED (check none is just a different spelling):")
        for c in p["new_clients"]:
            h = p["hints"].get(c["companyName"])
            print(f"   {c['companyName']:<36} {c['startMonth']} -> {c['churnMonth']}" + (f"   similar in app: {', '.join(h)}" if h else ""))
        if p["hints"]:
            print("   If any of those is the SAME client as one in the app, tell me and I'll map them (aliases), e.g.:")
            for nm, h in list(p["hints"].items())[:6]:
                print(f'      "{nm}": "{h[0]}"')
    if p["empty"]:
        print(f"\n{len(p['empty'])} app client(s) have NO invoices after the reset" + (" — they will be DELETED:" if args.delete_empty_clients else " — they are left in place (use --delete-empty-clients to remove them):"))
        for c in p["empty"][:40]:
            print(f"   {c['companyName']}")
        if len(p["empty"]) > 40:
            print(f"   … and {len(p['empty']) - 40} more")
    print(f"\nClients whose record changes: {len(p['changed_clients'])}  (start month recalculated; old-only clients also refreshed from the sheet)")
    for f in p["flagged"][:15]:
        print("   check:", f)
    if p["boundary"]:
        print(f"\n{len(p['boundary'])} invoice(s) from {args.keep_from} onwards would no longer follow on from the rebuilt history (numbers left alone; --renumber-kept fixes them):")
        for b in p["boundary"][:12]:
            print("   " + b)
        if len(p["boundary"]) > 12:
            print(f"   … and {len(p['boundary']) - 12} more (they are all counted, not changed)")
    write_report(args.report, diff)
    print(f"\nFull line-by-line list written to {args.report}")

    broken = violations(data, fixed, args.keep_from, args.renumber_kept, p["merged_ids"])
    if p["log"] and broken:
        # the Grantify / Film Farmers merge legitimately re-numbers those clients' invoices — only flag anything else
        broken = [x for x in broken if "renumbered" not in x]
    for x in broken:
        print("!! " + x)
    if broken or p["problems"]:
        print("\nREFUSING to go on while the problems above remain.")
        return 1
    if not args.apply:
        print("\nDRY RUN — nothing written. Re-run with --apply to make these changes.")
        return 0
    if p["hints"] and not args.create_similar:
        print("\nREFUSING to apply: these sheet names would be created as NEW clients but look like clients already in the app:")
        for nm, h in p["hints"].items():
            print(f"   {nm}  ~  {', '.join(h)}")
        print("Add them to the aliases file (--aliases) if they are the same client, or add --create-similar if they really are different clients.")
        return 1

    requests.post(f"{SUPABASE_URL}/rest/v1/credit_control_backups", headers=sb_headers("return=minimal"),
                  json={"data": data, "saved_by": "reset_from_sheet (pre-change snapshot)"}, timeout=30).raise_for_status()
    print("\nBacked up the current data to credit_control_backups.")
    w = requests.patch(f"{SUPABASE_URL}/rest/v1/credit_control_store", headers=sb_headers("return=representation"),
                       params={"id": "eq.1", "updated_at": f"eq.{version}"},
                       json={"data": fixed, "updated_at": datetime.datetime.now(datetime.timezone.utc).isoformat(), "updated_by": "reset_from_sheet"}, timeout=30)
    w.raise_for_status()
    if not w.json():
        print("ABORTED: someone saved between reading and writing. Nothing was changed. Re-run.")
        return 1
    chk = requests.get(f"{SUPABASE_URL}/rest/v1/credit_control_store", headers=sb_headers(),
                       params={"select": "data", "id": "eq.1"}, timeout=30).json()[0]["data"]
    ok = chk == fixed
    print("Verification PASSED — what is stored is exactly what the dry run showed: the months from " + args.keep_from + " on, the client mapping and every setting are untouched."
          if ok else "Verification FAILED — stored data is not what was expected. Stop and check the backup.")
    return 0 if ok else 1


class _Tee:
    """Everything printed also goes to a log file next to the report, so a run can be reviewed (or debugged) afterwards."""

    def __init__(self, stream, path):
        self.stream, self.file = stream, open(path, "w", encoding="utf-8")

    def write(self, text):
        self.stream.write(text)
        self.file.write(text)

    def flush(self):
        self.stream.flush()
        self.file.flush()


if __name__ == "__main__":
    import traceback

    _report = "reset-report.csv"
    for _i, _a in enumerate(sys.argv):
        if _a == "--report" and _i + 1 < len(sys.argv):
            _report = sys.argv[_i + 1]
    _log = str(Path(_report).with_suffix(".log"))
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stdout = _Tee(sys.stdout, _log)
    try:
        _code = main()
    except SystemExit as e:
        _code = e.code if isinstance(e.code, int) else (print(e.code) or 1)
    except Exception:
        traceback.print_exc(file=sys.stdout)
        _code = 1
    print(f"\n(This output is also saved in {_log})")
    raise SystemExit(_code)

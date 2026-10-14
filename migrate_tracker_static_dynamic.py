#!/usr/bin/env python3
"""
Migration: Old tracker -> New 3-sheet template.

Sheets populated:
  1) App Inventory
  2) Scans
  3) Vulnerabilities

Run Static and Dynamic separately, but keep the SAME migration_state.json.
That makes App ID, Scan ID and Vul ID continue globally across both trackers.

Examples:
  python migrate_tracker_static_dynamic.py --source static.xlsx --template new_template.xlsx --type Static --output static_migrated.xlsx
  python migrate_tracker_static_dynamic.py --source dynamic.xlsx --template new_template.xlsx --type Dynamic --output dynamic_migrated.xlsx

Important rules implemented:
- Latest Scan? is COPIED from the old tracker; it is NOT recalculated.
- Static: Instance Count expands one old row into N vulnerability rows.
- Dynamic: every old finding row becomes one vulnerability row.
- Static Finding From = old Assessment Type (SAST/SCA/IaC/Container).
- Dynamic Finding From: starts with M -> Manual Scan; starts with A -> Automated Scan.
- First scan = New. Later scan: if finding name existed in immediately previous scan,
  mark Recurring; otherwise New.
- First Identified Date -> Discovery Date.
- Static and Dynamic are separate assessment sources even when application names match.
- IDs remain one continuous global sequence.
"""

from pathlib import Path
from datetime import datetime, date
from copy import copy
import argparse
import json
import re
from openpyxl import load_workbook

STATE_FILE = "migration_state.json"


# ---------- Helpers ----------

def text(v):
    return "" if v is None else str(v).strip()


def norm(v):
    return re.sub(r"\s+", " ", text(v)).strip().lower()


def date_value(v):
    if v is None or v == "":
        return None
    if isinstance(v, (datetime, date)):
        return v
    s = text(v)
    for fmt in ("%d-%b-%Y", "%d-%B-%Y", "%d/%m/%Y",
                "%m/%d/%Y", "%Y-%m-%d", "%d-%m-%Y"):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            pass
    return v


def int_value(v, default=1):
    try:
        n = int(float(text(v)))
        return n if n > 0 else default
    except Exception:
        return default


def headers(ws):
    return {norm(c.value): c.column for c in ws[1] if text(c.value)}


def col(h, names):
    if isinstance(names, str):
        names = [names]
    for n in names:
        if norm(n) in h:
            return h[norm(n)]
    return None


def get(row, h, names, default=""):
    c = col(h, names)
    return default if c is None else row[c - 1].value


def put(ws, h, row_no, name, value):
    c = col(h, name)
    if c is None:
        raise ValueError(f"Missing column '{name}' in sheet '{ws.title}'.")
    ws.cell(row_no, c).value = value


def copy_style(ws, source_row, target_row):
    if source_row < 2 or target_row == source_row:
        return
    for c in range(1, ws.max_column + 1):
        a = ws.cell(source_row, c)
        b = ws.cell(target_row, c)
        if a.has_style:
            b._style = copy(a._style)
        b.number_format = a.number_format
        b.alignment = copy(a.alignment)
        b.protection = copy(a.protection)


def next_row(ws):
    return max(2, ws.max_row + 1)


def load_state(path):
    if not path.exists():
        return {"app": 1, "scan": 1, "finding": 1}
    data = json.loads(path.read_text(encoding="utf-8"))
    return {
        "app": int(data.get("app", 1)),
        "scan": int(data.get("scan", 1)),
        "finding": int(data.get("finding", 1)),
    }


def save_state(path, state):
    path.write_text(json.dumps(state, indent=2), encoding="utf-8")


def new_id(state, kind):
    n = state[kind]
    prefix = {"app": "APP", "scan": "SCAN", "finding": "FIND"}[kind]
    state[kind] += 1
    return f"{prefix}-{n:06d}"


def scan_number_key(v):
    s = text(v)
    try:
        return (0, int(float(s)))
    except Exception:
        return (1, s.lower())


def latest_value(v):
    """Preserve old tracker Latest Scan? while normalizing common values."""
    n = norm(v)
    if n in {"latest scan", "latest", "yes", "true", "y"}:
        return "Yes"
    if n in {"older scan", "older", "no", "false", "n"}:
        return "No"
    return text(v)


def finding_key(v):
    return norm(v)


# ---------- Source column aliases ----------

ALIASES = {
    "app": ["Application Name", "App Name", "Application"],
    "unique_app": ["Unique Application Name", "Unique App Name"],
    "scan_no": ["Scan Relevancy", "Scan Number", "Scan No", "Scan"],
    "latest": ["Latest Scan?", "Latest Scan", "Is Latest Scan"],
    "first_date": ["First Identified Date", "Discovery Date", "First Identified"],
    "scan_date": ["Scan Date", "Date", "Scan Completion Date"],
    "report_date": ["Report Shared Date", "Report Date"],
    "closed_date": ["Closed date", "Closed Date", "Closure Date"],
    "severity": ["Vulnerability Severity", "Severity"],
    "finding": ["Vulnerability Name", "Finding Title", "Finding Name"],
    "status": ["Remediation Status", "Status"],
    "assessment": ["Assessment Type", "Finding From", "Scan Type", "Execution Type"],
    "comment": ["Comment (if any)", "Comment", "Remarks"],
    "instances": ["Instance Count", "Instances", "Instance"],
    "url": ["App URL", "Application URL", "URL", "URL Link"],
    "bu": ["Business Unit", "BU", "First BU"],
    "criticality": ["Application Criticality", "App Criticality", "Criticality"],
    "frequency": ["Scan Frequency", "Frequency"],
    "tool": ["Tool", "Scan Tool"],
    "execution": ["Execution Type"],
}


def read_source(source):
    wb = load_workbook(source, data_only=False)
    ws = wb["Vulnerabilities"] if "Vulnerabilities" in wb.sheetnames else wb[wb.sheetnames[0]]
    h = headers(ws)

    required = {
        "app": "Application Name",
        "scan_no": "Scan Relevancy",
        "latest": "Latest Scan?",
        "first_date": "First Identified Date",
        "scan_date": "Scan Date",
        "severity": "Vulnerability Severity",
        "finding": "Vulnerability Name",
        "status": "Remediation Status",
        "assessment": "Assessment Type",
    }

    missing = [label for key, label in required.items()
               if col(h, ALIASES[key]) is None]
    if missing:
        raise ValueError(
            "Source tracker is missing required columns: " + ", ".join(missing)
        )

    records = []

    for r in range(2, ws.max_row + 1):
        row = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]

        app = text(get(row, h, ALIASES["app"]))
        if not app:
            continue

        assessment = text(get(row, h, ALIASES["assessment"]))

        records.append({
            "app": app,
            "unique_app": text(get(row, h, ALIASES["unique_app"], app)) or app,
            "scan_no": text(get(row, h, ALIASES["scan_no"])),
            "latest": latest_value(get(row, h, ALIASES["latest"])),
            "first_date": date_value(get(row, h, ALIASES["first_date"])),
            "scan_date": date_value(get(row, h, ALIASES["scan_date"])),
            "report_date": date_value(get(row, h, ALIASES["report_date"])),
            "closed_date": date_value(get(row, h, ALIASES["closed_date"])),
            "severity": text(get(row, h, ALIASES["severity"])),
            "finding": text(get(row, h, ALIASES["finding"])),
            "status": text(get(row, h, ALIASES["status"])),
            "assessment": assessment,
            "comment": text(get(row, h, ALIASES["comment"])),
            "instances": int_value(get(row, h, ALIASES["instances"]), 1),
            "url": text(get(row, h, ALIASES["url"])),
            "bu": text(get(row, h, ALIASES["bu"])),
            "criticality": text(get(row, h, ALIASES["criticality"])),
            "frequency": text(get(row, h, ALIASES["frequency"])),
            "tool": text(get(row, h, ALIASES["tool"])),
            "execution": text(get(row, h, ALIASES["execution"])),
        })

    return records


# ---------- Migration ----------

def migrate_one(source, template_wb, tracker_type, state):
    """
    Migrate one source workbook into the already-open combined template.
    IDs are allocated from the shared state, so Static + Dynamic are continuous.
    """
    source = Path(source)
    tracker_type = tracker_type.lower()
    scan_type = "Static" if tracker_type == "static" else "Dynamic"

    records = read_source(source)
    if not records:
        print(f"WARNING: No data found in {source.name}; skipping.")
        return {"apps": 0, "scans": 0, "findings": 0}

    wb = template_wb
    app_ws = wb["App Inventory"]
    scan_ws = wb["Scans"]
    vul_ws = wb["Vulnerabilities"]

    ah, sh, vh = headers(app_ws), headers(scan_ws), headers(vul_ws)

    required_new = {
        "App Inventory": ["App ID", "Application Name",
                          "Unique Application Name", "Scan Type"],
        "Scans": ["Scan ID", "App ID", "App Name", "Date",
                  "Scan Type", "Latest Scan?"],
        "Vulnerabilities": ["Vul ID", "App ID", "Scan ID",
                            "Finding Title", "Severity",
                            "Vulnerability Type", "Discovery Date",
                            "Finding From", "Scan Type", "Status",
                            "Latest Scan?"],
    }

    for sheet, needed in required_new.items():
        hm = {"App Inventory": ah, "Scans": sh, "Vulnerabilities": vh}[sheet]
        missing = [x for x in needed if col(hm, x) is None]
        if missing:
            raise ValueError(f"{sheet} missing columns: {', '.join(missing)}")

    # Application IDs: same application name in Static and Dynamic is treated
    # as two assessment-source records, but IDs continue globally.
    app_id = {}
    app_first_record = {}

    for rec in records:
        key = (norm(rec["app"]), tracker_type)
        if key not in app_id:
            app_id[key] = new_id(state, "app")
            app_first_record[key] = rec

    # Scan groups are application + scan number + assessment source.
    groups = {}
    group_order = []

    for rec in records:
        key = (norm(rec["app"]), rec["scan_no"], tracker_type)

        if key not in groups:
            groups[key] = {
                "app": rec["app"],
                "scan_no": rec["scan_no"],
                "latest": rec["latest"],       # OLD TRACKER VALUE
                "scan_date": rec["scan_date"],
                "report_date": rec["report_date"],
                "url": rec["url"],
                "bu": rec["bu"],
                "criticality": rec["criticality"],
                "frequency": rec["frequency"],
                "tool": rec["tool"],
                "execution": rec["execution"],
                "records": [],
                "app_key": (norm(rec["app"]), tracker_type),
            }
            group_order.append(key)

        groups[key]["records"].append(rec)

        # Preserve old tracker Latest Scan?; don't calculate it from dates.
        if norm(rec["latest"]) in {"yes", "latest", "latest scan", "true", "y"}:
            groups[key]["latest"] = "Yes"

        for f in ("scan_date", "report_date", "url", "bu",
                  "criticality", "frequency", "tool", "execution"):
            if not groups[key][f] and rec[f]:
                groups[key][f] = rec[f]

    # Global Scan IDs.
    scan_id = {key: new_id(state, "scan") for key in group_order}

    # Previous-scan finding lookup for New/Recurring.
    by_app = {}
    for rec in records:
        by_app.setdefault((norm(rec["app"]), tracker_type), {}).setdefault(
            rec["scan_no"], []
        ).append(rec)

    previous_names = {}

    for app_key, scan_groups in by_app.items():
        scan_numbers = sorted(scan_groups.keys(), key=scan_number_key)

        for i, scan_no in enumerate(scan_numbers):
            group_key = (app_key[0], scan_no, tracker_type)

            if i == 0:
                previous_names[group_key] = set()
            else:
                previous_scan = scan_numbers[i - 1]
                previous_names[group_key] = {
                    finding_key(x["finding"])
                    for x in scan_groups[previous_scan]
                    if x["finding"]
                }

    # ---- App Inventory ----
    row_no = next_row(app_ws)

    for key, rec in app_first_record.items():
        copy_style(app_ws, max(2, row_no - 1), row_no)
        put(app_ws, ah, row_no, "App ID", app_id[key])
        put(app_ws, ah, row_no, "Application Name", rec["app"])
        put(app_ws, ah, row_no, "Unique Application Name", rec["unique_app"])
        put(app_ws, ah, row_no, "Scan Type", scan_type)

        optional = {
            "App URL": rec["url"],
            "Business Unit": rec["bu"],
            "Application Criticality": rec["criticality"],
            "Scan Frequency": rec["frequency"],
        }
        for name, value in optional.items():
            if col(ah, name) is not None:
                put(app_ws, ah, row_no, name, value)

        row_no += 1

    # ---- Scans ----
    row_no = next_row(scan_ws)

    for key in group_order:
        g = groups[key]
        copy_style(scan_ws, max(2, row_no - 1), row_no)

        put(scan_ws, sh, row_no, "Scan ID", scan_id[key])
        put(scan_ws, sh, row_no, "App ID", app_id[g["app_key"]])
        put(scan_ws, sh, row_no, "App Name", g["app"])
        put(scan_ws, sh, row_no, "Date", g["scan_date"])
        put(scan_ws, sh, row_no, "Scan Type", scan_type)

        # OLD TRACKER -> NEW TRACKER
        put(scan_ws, sh, row_no, "Latest Scan?", g["latest"])

        optional = {
            "Tool": g["tool"],
            "Execution Type": g["execution"],
            "Scan Completion Date": g["scan_date"],
            "Status": "Completed",
        }
        for name, value in optional.items():
            if col(sh, name) is not None:
                put(scan_ws, sh, row_no, name, value)

        row_no += 1

    # ---- Vulnerabilities ----
    row_no = next_row(vul_ws)
    finding_total = 0

    for key in group_order:
        g = groups[key]
        prev = previous_names.get(key, set())

        try:
            scan_num = int(float(g["scan_no"]))
            first_scan = scan_num == 1
        except Exception:
            first_scan = norm(g["scan_no"]) in {"first scan", "first"}

        for rec in g["records"]:
            # "Not Found" means the application was scanned but no actual
            # vulnerability finding was identified. Keep the scan in Scans,
            # but NEVER create a Vulnerabilities row for this source record.
            if norm(rec["severity"]) == "not found":
                continue

            # Static: Instance Count creates individual finding rows.
            # Dynamic: each old row is already one individual finding.
            copies = rec["instances"] if tracker_type == "static" else 1

            if tracker_type == "static":
                finding_from = rec["assessment"]
            else:
                a = rec["assessment"]
                if norm(a).startswith("m"):
                    finding_from = "Manual Scan"
                elif norm(a).startswith("a"):
                    finding_from = "Automated Scan"
                else:
                    finding_from = a

            vtype = (
                "New"
                if first_scan
                else ("Recurring" if finding_key(rec["finding"]) in prev else "New")
            )

            for _ in range(copies):
                copy_style(vul_ws, max(2, row_no - 1), row_no)

                fid = new_id(state, "finding")
                finding_total += 1

                put(vul_ws, vh, row_no, "Vul ID", fid)
                put(vul_ws, vh, row_no, "App ID", app_id[g["app_key"]])
                put(vul_ws, vh, row_no, "Scan ID", scan_id[key])
                put(vul_ws, vh, row_no, "Finding Title", rec["finding"])
                put(vul_ws, vh, row_no, "Severity", rec["severity"])
                put(vul_ws, vh, row_no, "Vulnerability Type", vtype)
                put(vul_ws, vh, row_no, "Discovery Date", rec["first_date"])
                put(vul_ws, vh, row_no, "Finding From", finding_from)
                put(vul_ws, vh, row_no, "Scan Type", scan_type)

                status = rec["status"]
                if norm(status) == "active":
                    status = "Open"
                elif norm(status) == "fixed":
                    status = "Closed"
                put(vul_ws, vh, row_no, "Status", status)

                # OLD TRACKER Latest Scan? -> EVERY vulnerability instance.
                put(vul_ws, vh, row_no, "Latest Scan?", g["latest"])

                optional = {
                    "Report Shared Date": rec["report_date"],
                    "Closure Date": rec["closed_date"],
                    "Comment": rec["comment"],
                    "App URL": rec["url"],
                }
                for name, value in optional.items():
                    if col(vh, name) is not None:
                        put(vul_ws, vh, row_no, name, value)

                row_no += 1

    return {
        "apps": len(app_id),
        "scans": len(groups),
        "findings": finding_total,
    }


def main():
    """
    ONE-GO MODE.

    Put these beside this script:
        static.xlsx
        dynamic.xlsx
        new_template.xlsx

    Then run:
        python migrate_tracker_static_dynamic.py

    The output is:
        migrated_tracker.xlsx
    """
    base = Path(__file__).resolve().parent

    static_file = base / "static.xlsx"
    dynamic_file = base / "dynamic.xlsx"
    template_file = base / "new_template.xlsx"
    output_file = base / "migrated_tracker.xlsx"
    state_file = base / STATE_FILE

    for f in (static_file, dynamic_file, template_file):
        if not f.exists():
            raise FileNotFoundError(
                f"\nMissing file: {f.name}\n"
                f"Place {f.name} in the same folder as this script."
            )

    # Start from the empty new template ONCE, then append both sources.
    wb = load_workbook(template_file)

    # Start fresh for this one combined migration.
    # If migration_state.json already exists from an earlier test, delete it
    # or use a new folder so IDs start at APP-000001 / SCAN-000001 / FIND-000001.
    state = {"app": 1, "scan": 1, "finding": 1}

    print("Starting combined migration...")
    print(f"Static source : {static_file.name}")
    print(f"Dynamic source: {dynamic_file.name}")
    print(f"Template      : {template_file.name}")
    print()

    static_result = migrate_one(
        static_file, wb, "static", state
    )

    dynamic_result = migrate_one(
        dynamic_file, wb, "dynamic", state
    )

    # Save the final combined workbook.
    wb.save(output_file)

    # Save final sequence so you can see where the IDs ended.
    save_state(state_file, state)

    print("\n========================================")
    print("COMBINED MIGRATION COMPLETED")
    print("========================================")
    print(f"Output: {output_file}")
    print()
    print("Static:")
    print(f"  Applications: {static_result['apps']}")
    print(f"  Scans       : {static_result['scans']}")
    print(f"  Findings    : {static_result['findings']}")
    print()
    print("Dynamic:")
    print(f"  Applications: {dynamic_result['apps']}")
    print(f"  Scans       : {dynamic_result['scans']}")
    print(f"  Findings    : {dynamic_result['findings']}")
    print()
    print("Final ID sequence:")
    print(f"  Next App ID    : APP-{state['app']:06d}")
    print(f"  Next Scan ID   : SCAN-{state['scan']:06d}")
    print(f"  Next Finding ID: FIND-{state['finding']:06d}")
    print()
    print("Latest Scan? is copied from the old trackers.")
    print("Static and Dynamic are separate assessment sources.")
    print("IDs are continuous across both sources.")


if __name__ == "__main__":
    main()

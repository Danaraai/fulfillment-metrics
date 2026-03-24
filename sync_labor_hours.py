"""
sync_labor_hours.py — Sync labor hours from source daily sheet → Export PowerBI 'Labor Hours 2026' tab.

Reads the source daily-tracking spreadsheet (one tab per working day, e.g. 'March 23'),
extracts the daily summary totals, and appends any new rows to the 'Labor Hours 2026'
tab in the Export PowerBI spreadsheet.

USAGE:
  python3 sync_labor_hours.py            # Sync new rows
  python3 sync_labor_hours.py --dry-run  # Preview without writing
  python3 sync_labor_hours.py --schedule # Install daily 5pm Eastern launchd job
  python3 sync_labor_hours.py --unschedule # Remove the schedule
"""

import argparse
import calendar
import json
import os
import pickle
import subprocess
import sys
import warnings
from datetime import date
from pathlib import Path

warnings.filterwarnings("ignore")

BASE_DIR      = Path(__file__).parent
CONFIG_FILE   = BASE_DIR / "config.json"
TOKEN_CACHE   = BASE_DIR / "google_token.pickle"
LOG_DIR       = BASE_DIR / "logs"
PLIST_LABEL   = "com.powerbi.sync-labor-hours"
PLIST_PATH    = Path.home() / "Library" / "LaunchAgents" / f"{PLIST_LABEL}.plist"

# ── Config ─────────────────────────────────────────────────────────────────────
def load_config() -> dict:
    with open(CONFIG_FILE) as f:
        return json.load(f)

# ── Google auth ────────────────────────────────────────────────────────────────
def get_google_creds(config: dict):
    from google.oauth2.credentials import Credentials
    from google.auth.transport.requests import Request

    SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
    creds  = None
    if TOKEN_CACHE.exists():
        with open(TOKEN_CACHE, "rb") as f:
            creds = pickle.load(f)
    if creds and creds.valid:
        return creds
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
        with open(TOKEN_CACHE, "wb") as f:
            pickle.dump(creds, f)
        return creds
    raise RuntimeError("Google token expired or missing. Run run_now.py once to re-authenticate.")

# ── Time parsing ───────────────────────────────────────────────────────────────
def parse_hms(s: str) -> float:
    """Convert 'HH:MM:SS' (HH can exceed 24) to decimal hours."""
    try:
        parts = str(s).strip().split(":")
        return int(parts[0]) + int(parts[1]) / 60 + (int(parts[2]) if len(parts) > 2 else 0) / 3600
    except Exception:
        return 0.0

# ── Read source sheet ──────────────────────────────────────────────────────────
def read_source_labor_hours(svc, labor_sheet_id: str) -> list[dict]:
    """Read all daily tabs from the source sheet. Returns list of row dicts."""
    meta     = svc.spreadsheets().get(spreadsheetId=labor_sheet_id).execute()
    all_tabs = [s["properties"]["title"] for s in meta["sheets"]]

    month_map = {}
    for i, name in enumerate(calendar.month_name):
        if name: month_map[name.lower()] = i
    for i, name in enumerate(calendar.month_abbr):
        if name: month_map[name.lower()] = i

    rows = []
    for tab in all_tabs:
        parts = tab.strip().split()
        if len(parts) != 2:
            continue
        month_num = month_map.get(parts[0].lower())
        if not month_num:
            continue
        try:
            day = int(parts[1])
        except ValueError:
            continue

        year = 2026  # Labor Hours 2026 sheet
        try:
            tab_date = date(year=year, month=month_num, day=day)
        except Exception:
            continue

        try:
            res  = svc.spreadsheets().values().get(
                spreadsheetId=labor_sheet_id,
                range=f"'{tab}'!O1:P12"
            ).execute()
            vals = res.get("values", [])
        except Exception:
            continue

        summary = {r[0]: r[1] for r in vals if len(r) >= 2}

        rows.append({
            "Date":           str(tab_date),          # "2026-03-23"
            "Total Hours":    round(parse_hms(summary.get("Total Labor Hours",
                              summary.get("Total Emp/Temp Hours", "0"))), 4),
            "Emp Hours":      round(parse_hms(summary.get("Total Emp Hours",  "0")), 4),
            "Temp Hours":     round(parse_hms(summary.get("Total Temp Hours", "0")), 4),
            "Outbound Hours": round(parse_hms(summary.get("Total Outbound Hours", "0")), 4),
            "Inbound Hours":  round(parse_hms(summary.get("Total Inbound Hours",  "0")), 4),
            "Headcount":      0,
        })

    return sorted(rows, key=lambda r: r["Date"])

# ── Sync to Export PowerBI sheet ───────────────────────────────────────────────
def sync_to_export_sheet(svc, export_sheet_id: str, source_rows: list[dict],
                          dry_run: bool = False) -> int:
    """Append rows with dates not already present. Returns count of rows added."""
    tab = "Labor Hours 2026"

    # Fetch existing dates from column A
    resp = svc.spreadsheets().values().get(
        spreadsheetId=export_sheet_id,
        range=f"{tab}!A:A"
    ).execute()
    existing_vals = resp.get("values", [])
    existing_dates = {str(r[0]).strip() for r in existing_vals if r}

    HEADERS = ["Date", "Total Hours", "Emp Hours", "Temp Hours",
               "Outbound Hours", "Inbound Hours", "Headcount"]

    new_rows = [r for r in source_rows if r["Date"] not in existing_dates]

    if not new_rows:
        print("  ✓ Labor Hours 2026 is already up to date. Nothing to append.")
        return 0

    print(f"  Found {len(new_rows)} new row(s) to append:")
    for r in new_rows:
        print(f"    {r['Date']}  Total={r['Total Hours']}h  Outbound={r['Outbound Hours']}h")

    if dry_run:
        print("  [dry-run] Skipping write.")
        return len(new_rows)

    # Append
    values_to_write = [[r[h] for h in HEADERS] for r in new_rows]
    svc.spreadsheets().values().append(
        spreadsheetId=export_sheet_id,
        range=f"{tab}!A:G",
        valueInputOption="USER_ENTERED",
        insertDataOption="INSERT_ROWS",
        body={"values": values_to_write}
    ).execute()

    print(f"  ✓ Appended {len(new_rows)} row(s) to '{tab}'.")
    return len(new_rows)

# ── Schedule helpers ───────────────────────────────────────────────────────────
def install_schedule():
    """Install a launchd plist to run daily at 5pm Eastern (= 2pm Pacific)."""
    python  = sys.executable
    script  = str(Path(__file__).resolve())
    LOG_DIR.mkdir(exist_ok=True)
    plist = f"""<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN"
    "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>Label</key><string>{PLIST_LABEL}</string>
    <key>ProgramArguments</key>
    <array>
        <string>{python}</string>
        <string>{script}</string>
    </array>
    <key>StartCalendarInterval</key>
    <dict>
        <key>Hour</key><integer>14</integer>
        <key>Minute</key><integer>5</integer>
    </dict>
    <key>WorkingDirectory</key><string>{BASE_DIR}</string>
    <key>StandardOutPath</key><string>{LOG_DIR}/sync_labor.log</string>
    <key>StandardErrorPath</key><string>{LOG_DIR}/sync_labor_error.log</string>
    <key>RunAtLoad</key><false/>
</dict>
</plist>"""
    PLIST_PATH.write_text(plist)
    subprocess.run(["launchctl", "unload", str(PLIST_PATH)],
                   capture_output=True)
    subprocess.run(["launchctl", "load", str(PLIST_PATH)], check=True)
    print(f"✓ Scheduled: daily at 2:05 PM Pacific / 5:05 PM Eastern")
    print(f"  Plist: {PLIST_PATH}")

def uninstall_schedule():
    subprocess.run(["launchctl", "unload", str(PLIST_PATH)], capture_output=True)
    if PLIST_PATH.exists():
        PLIST_PATH.unlink()
    print("✓ Schedule removed.")

# ── Main ───────────────────────────────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--dry-run",    action="store_true")
    parser.add_argument("--schedule",   action="store_true")
    parser.add_argument("--unschedule", action="store_true")
    args = parser.parse_args()

    if args.schedule:
        install_schedule()
        return
    if args.unschedule:
        uninstall_schedule()
        return

    from googleapiclient.discovery import build

    config          = load_config()
    labor_sheet_id  = config["labor_hours_sheet_id"]
    export_sheet_id = config["google_sheets"]["spreadsheet_id"]

    print("=" * 60)
    print("  Labor Hours Sync")
    print(f"  Source : {labor_sheet_id}")
    print(f"  Target : Export PowerBI → Labor Hours 2026")
    print("=" * 60)

    creds = get_google_creds(config)
    svc   = build("sheets", "v4", credentials=creds)

    print("\n[1/2] Reading source daily tabs…")
    source_rows = read_source_labor_hours(svc, labor_sheet_id)
    print(f"  Loaded {len(source_rows)} days  "
          f"({source_rows[0]['Date']} → {source_rows[-1]['Date']})")

    print("\n[2/2] Syncing to Export PowerBI…")
    added = sync_to_export_sheet(svc, export_sheet_id, source_rows, dry_run=args.dry_run)

    print(f"\n{'[dry-run] ' if args.dry_run else ''}Done — {added} row(s) added.")

if __name__ == "__main__":
    main()

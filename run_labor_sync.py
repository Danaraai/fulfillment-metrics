"""
run_labor_sync.py
Reads daily tabs from the Labor Hours source sheets and writes
consolidated rows to the destination sheet — same logic as labor_sync.gs
but runs via Python/Sheets API (no browser needed).

Usage:
  python3 run_labor_sync.py          # sync both 2025 and 2026
  python3 run_labor_sync.py --2025   # 2025 only
  python3 run_labor_sync.py --2026   # 2026 only
"""

import pickle, warnings, re, sys
from datetime import datetime, date, timedelta
from pathlib import Path

warnings.filterwarnings("ignore")
from googleapiclient.discovery import build

# ── Config ────────────────────────────────────────────────────────────────────

BASE_DIR  = Path(__file__).parent
TOKEN     = BASE_DIR / "google_token.pickle"

DEST_ID   = "1tuo7knxTvOR3snd_u1AnnW1iiN9l1-TaOLi-YaVqLM0"
SRC_2025  = "1z3SFQrxnSOVIzPizHupDMmmzl705tGn4bpsn9rEpd9w"
SRC_2026  = "1QHSkxmnuUaQtVsiTQL5YvCqOKZqDnMr1WqJODKKAqh0"

TAB_2025  = "Labor Hours 2025"
TAB_2026  = "Labor Hours 2026"

SKIP_TABS = {"EOM Report", "In/Out Hours", "Tracker"}
STATUSES  = {"present", "late", "sick", "call out", "absent"}
HEADER    = ["Date", "Total Hours", "Emp Hours", "Temp Hours",
             "Outbound Hours", "Inbound Hours", "Headcount"]

YESTERDAY = date.today() - timedelta(days=1)

# ── Helpers ───────────────────────────────────────────────────────────────────

def parse_hours(s):
    if not s or not str(s).strip():
        return 0.0
    parts = str(s).strip().split(":")
    try:
        h = int(parts[0])
        m = int(parts[1]) if len(parts) > 1 else 0
        return round(h + m / 60, 2)
    except Exception:
        return 0.0

def next_hours(row, idx):
    for v in row[idx + 1:]:
        s = str(v).strip()
        if re.match(r"^\d+:\d+", s):
            return parse_hours(s)
    return 0.0

def parse_tab_date(tab_name, default_year):
    MONTHS = {
        "jan":1,"feb":2,"mar":3,"apr":4,"may":5,"jun":6,
        "jul":7,"aug":8,"sep":9,"oct":10,"nov":11,"dec":12,
        "january":1,"february":2,"march":3,"april":4,"june":6,
        "july":7,"august":8,"september":9,"october":10,"november":11,"december":12
    }
    m = re.match(r"^(\w+)\s+(\d+)(?:,\s*(\d{4}))?$", tab_name.strip())
    if not m:
        return None
    mon = MONTHS.get(m.group(1).lower())
    if not mon:
        return None
    day  = int(m.group(2))
    year = int(m.group(3)) if m.group(3) else default_year
    try:
        return date(year, mon, day)
    except ValueError:
        return None

def extract_day(rows):
    emp = temp = total = outbound = inbound = headcount = 0.0
    for row in rows:
        # Headcount
        status = str(row[2]).strip().lower() if len(row) > 2 else ""
        name   = str(row[0]).strip() if len(row) > 0 else ""
        if name and status in STATUSES:
            headcount += 1
        # Summary labels
        for i, cell in enumerate(row):
            c = str(cell).strip().lower()
            if "total emp hours" in c and "temp" not in c:
                emp = next_hours(row, i)
            elif "total temp hours" in c:
                temp = next_hours(row, i)
            elif "total emp/temp" in c or "total emp / temp" in c:
                total = next_hours(row, i)
            elif "total labor hours" in c:
                v = next_hours(row, i)
                if v > 0:
                    total = v
            elif "total outbound" in c:
                outbound = next_hours(row, i)
            elif "total inbound" in c:
                inbound = next_hours(row, i)
    if total == 0 and (emp > 0 or temp > 0):
        total = round(emp + temp, 2)
    return {
        "total": total, "emp": emp, "temp": temp,
        "outbound": outbound, "inbound": inbound,
        "headcount": int(headcount)
    }

# ── Core sync ─────────────────────────────────────────────────────────────────

def sync(svc, source_id, dest_tab_name, year):
    print(f"\n{'='*55}")
    print(f"  Syncing {dest_tab_name}")
    print(f"  Cutoff: {YESTERDAY}")
    print(f"{'='*55}")

    # Get all source tabs
    meta = svc.spreadsheets().get(spreadsheetId=source_id).execute()
    all_tabs = [s["properties"]["title"] for s in meta["sheets"]]
    daily_tabs = [t for t in all_tabs if t not in SKIP_TABS]
    print(f"  Source tabs: {len(daily_tabs)} (skipping {len(all_tabs)-len(daily_tabs)} non-daily)")

    # Get or create destination tab
    dest_meta = svc.spreadsheets().get(spreadsheetId=DEST_ID).execute()
    dest_tab_names = [s["properties"]["title"] for s in dest_meta["sheets"]]

    if dest_tab_name not in dest_tab_names:
        print(f"  Creating tab '{dest_tab_name}'...")
        svc.spreadsheets().batchUpdate(
            spreadsheetId=DEST_ID,
            body={"requests": [{"addSheet": {"properties": {"title": dest_tab_name}}}]}
        ).execute()

    # Read existing dates
    existing_dates = set()
    try:
        existing = svc.spreadsheets().values().get(
            spreadsheetId=DEST_ID,
            range=f"'{dest_tab_name}'!A:A"
        ).execute().get("values", [])
        for row in existing[1:]:  # skip header
            if row:
                existing_dates.add(str(row[0]).strip())
    except Exception:
        pass

    # Write header if needed
    if not existing:
        svc.spreadsheets().values().update(
            spreadsheetId=DEST_ID,
            range=f"'{dest_tab_name}'!A1",
            valueInputOption="RAW",
            body={"values": [HEADER]}
        ).execute()

    # Batch-fetch all daily tabs (50 at a time)
    new_rows = []
    BATCH = 50
    for i in range(0, len(daily_tabs), BATCH):
        chunk = daily_tabs[i:i+BATCH]
        ranges = [f"'{t}'!A1:S25" for t in chunk]
        resp = svc.spreadsheets().values().batchGet(
            spreadsheetId=source_id, ranges=ranges
        ).execute()

        for vr in resp.get("valueRanges", []):
            # Extract tab name from range string e.g. "'May 1'!A1:S25"
            m = re.match(r"^'?([^'!]+)'?!", vr.get("range", ""))
            tab_name = m.group(1).strip("'") if m else ""
            tab_date = parse_tab_date(tab_name, year)
            if not tab_date:
                continue
            if tab_date > YESTERDAY:
                continue  # skip pre-created future tabs
            date_str = tab_date.isoformat()
            if date_str in existing_dates:
                continue  # already synced

            rows = vr.get("values", [])
            if not rows:
                continue
            d = extract_day(rows)
            new_rows.append([
                date_str, d["total"], d["emp"], d["temp"],
                d["outbound"], d["inbound"], d["headcount"]
            ])

        done = min(i + BATCH, len(daily_tabs))
        print(f"  Scanned {done}/{len(daily_tabs)} tabs...", end="\r")

    print()

    # Sort and append
    new_rows.sort(key=lambda r: r[0])

    if new_rows:
        # Find next empty row
        dest_data = svc.spreadsheets().values().get(
            spreadsheetId=DEST_ID, range=f"'{dest_tab_name}'!A:A"
        ).execute().get("values", [])
        next_row = len(dest_data) + 1

        svc.spreadsheets().values().update(
            spreadsheetId=DEST_ID,
            range=f"'{dest_tab_name}'!A{next_row}",
            valueInputOption="USER_ENTERED",
            body={"values": new_rows}
        ).execute()
        print(f"  ✓ Wrote {len(new_rows)} new rows  (skipped {len(existing_dates)} already present)")
    else:
        print(f"  ✓ Nothing new to add ({len(existing_dates)} rows already present)")

# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    with open(TOKEN, "rb") as f:
        creds = pickle.load(f)
    svc = build("sheets", "v4", credentials=creds)

    args = sys.argv[1:]
    do_2025 = "--2026" not in args
    do_2026 = "--2025" not in args

    if do_2025:
        sync(svc, SRC_2025, TAB_2025, 2025)
    if do_2026:
        sync(svc, SRC_2026, TAB_2026, 2026)

    print("\n✓ All done.")

if __name__ == "__main__":
    main()

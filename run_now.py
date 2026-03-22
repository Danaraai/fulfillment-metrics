"""
run_now.py — Jack Archer weekly Power BI → Google Sheets export.

HOW IT WORKS:
  1. Calculates the correct Friday→Thursday date range for the week just ended
  2. Opens the Power BI report in your real Chrome browser
  3. Shows you exactly which dates to set in the date slicer
  4. Waits (up to 10 min) for you to export the CSV
  5. Deduplicates against rows already in the sheet (by Order ID)
  6. Appends only NEW rows to the Export tab

USAGE:
  python3 run_now.py                  # Run the full export
  python3 run_now.py --dry-run        # Preview without writing to Sheets
  python3 run_now.py --upload-only    # Skip Chrome, re-upload most recent ~/Downloads/data.csv
  python3 run_now.py --schedule       # Install macOS weekly Saturday 8am job
  python3 run_now.py --unschedule     # Remove the schedule
"""

import argparse
import json
import os
import pickle
import shutil
import subprocess
import sys
import time
import warnings
from datetime import datetime, date, timedelta
from pathlib import Path

warnings.filterwarnings("ignore")

# ── Paths ─────────────────────────────────────────────────────────────────────
BASE_DIR     = Path(__file__).parent
CONFIG_FILE  = BASE_DIR / "config.json"
DOWNLOADS    = Path.home() / "Downloads"
ARCHIVE_DIR  = BASE_DIR / "downloads"
TOKEN_CACHE  = BASE_DIR / "google_token.pickle"
LOG_DIR      = BASE_DIR / "logs"
PLIST_LABEL  = "com.powerbi.jack-archer-export"
PLIST_PATH   = Path.home() / "Library" / "LaunchAgents" / f"{PLIST_LABEL}.plist"


# ── Config ────────────────────────────────────────────────────────────────────
def load_config() -> dict:
    with open(CONFIG_FILE) as f:
        return json.load(f)


# ── Date helpers ──────────────────────────────────────────────────────────────
def get_week_range(ref: date = None) -> tuple:
    """
    Returns (week_start, week_end) for the most recently completed Fri→Fri week.

    Logic (run on Saturday):
      - last Friday  = yesterday               (end of the week)
      - week_start   = last Friday minus 7     (the Friday it started)
      - week_end     = last Friday             (inclusive)

    Example — Saturday 2026-03-21:
      last_friday = 2026-03-20
      week_start  = 2026-03-13
      week_end    = 2026-03-20
    """
    if ref is None:
        ref = date.today()
    # Find the most recent Friday on or before ref
    days_since_fri = (ref.weekday() - 4) % 7   # Mon=0…Fri=4…Sat=5
    last_friday    = ref - timedelta(days=days_since_fri)
    week_end       = last_friday                # Friday (inclusive)
    week_start     = last_friday - timedelta(days=7)   # Friday 7 days earlier
    return week_start, week_end


def fmt(d: date) -> str:       return d.strftime("%B %d, %Y")    # March 13, 2026
def fmt_s(d: date) -> str:     return d.strftime("%Y-%m-%d")     # 2026-03-13
def week_label(ws, we) -> str: return f"{fmt_s(ws)} → {fmt_s(we)}"


# ── Chrome automation (AppleScript + JS injection — no Playwright needed) ─────

def _run_js(js: str) -> str:
    """Execute JS in the front Chrome tab via AppleScript. No DevTools needed."""
    import tempfile, os
    with tempfile.NamedTemporaryFile(mode='w', suffix='.js', delete=False, encoding='utf-8') as f:
        f.write(js)
        tmp = f.name
    try:
        script = f"""set jsCode to (do shell script "cat '{tmp}'")
tell application "Google Chrome"
    execute front window's active tab javascript jsCode
end tell"""
        r = subprocess.run(["osascript", "-e", script], capture_output=True, text=True, timeout=15)
        return r.stdout.strip()
    except Exception as e:
        return f"error: {e}"
    finally:
        os.unlink(tmp)


def open_chrome(url: str):
    """Open a URL in Google Chrome via AppleScript."""
    script = f'''tell application "Google Chrome"
        activate
        open location "{url}"
    end tell'''
    subprocess.run(["osascript", "-e", script])


def automate_chrome_export(report_url: str, week_end: date) -> bool:
    """
    Fully automated Chrome export via AppleScript + JS injection.
    Steps: open report → set date → click ··· → Export data → Summarized → CSV → Export.
    Returns True if export was triggered successfully.
    """
    end_str = week_end.strftime("%-m/%-d/%Y")   # e.g. "3/20/2026"
    day_num = str(week_end.day)                  # e.g. "20"

    # 1 — Open Chrome
    print("  Opening Chrome and loading report...")
    open_chrome(report_url)

    # 2 — Wait for the vcMenuBtn (the ··· button) to appear — up to 90s
    print("  Waiting for report to render...")
    for attempt in range(45):
        result = _run_js("document.querySelector('button.vcMenuBtn') ? 'ready' : 'loading'")
        if "ready" in result:
            print(f"  Report ready (after ~{attempt*2}s)")
            break
        time.sleep(2)
    else:
        print("  WARNING: Report did not fully load. Trying anyway...")

    time.sleep(2)

    # 3 — Set the end date on the Transaction Date slicer
    print(f"  Setting date range end → {end_str}...")
    set_result = _run_js(f"""
(function() {{
    var inputs = Array.from(document.querySelectorAll('input'))
                      .filter(function(i) {{ return /\\d+\\/\\d+\\/\\d+/.test(i.value); }});
    if (inputs.length < 2) {{ return 'error: only ' + inputs.length + ' date inputs found'; }}
    var endInput = inputs[inputs.length - 1];
    endInput.focus();
    endInput.select();
    document.execCommand('selectAll');
    document.execCommand('insertText', false, '{end_str}');
    endInput.dispatchEvent(new KeyboardEvent('keydown',  {{bubbles:true, cancelable:true, key:'Enter', keyCode:13}}));
    endInput.dispatchEvent(new KeyboardEvent('keypress', {{bubbles:true, cancelable:true, key:'Enter', keyCode:13}}));
    endInput.dispatchEvent(new KeyboardEvent('keyup',    {{bubbles:true, cancelable:true, key:'Enter', keyCode:13}}));
    return 'set end to {end_str}';
}})()
""")
    print(f"  Date JS: {set_result}")
    time.sleep(2)

    # 4 — If a calendar popup appeared, click the correct day number
    cal_result = _run_js(f"""
(function() {{
    var cells = Array.from(document.querySelectorAll(
        'td, [class*="day"], [class*="calendar"] button, [role="gridcell"] button'
    ));
    var target = cells.find(function(c) {{
        return c.textContent.trim() === '{day_num}' && !c.disabled && !c.classList.contains('disabled');
    }});
    if (target) {{ target.click(); return 'clicked day {day_num}'; }}
    return 'no calendar found';
}})()
""")
    print(f"  Calendar JS: {cal_result}")
    time.sleep(3)   # Let data reload with new date range

    # 5 — Click the ··· More Options button
    print("  Clicking More Options (···)...")
    _run_js("document.querySelector('button.vcMenuBtn').click()")
    time.sleep(1)

    # 6 — Click "Export data" in the context menu
    print("  Clicking Export data...")
    _run_js("""
(function() {
    var items = Array.from(document.querySelectorAll('button, li, [role="menuitem"], a'));
    var target = items.find(function(el) { return el.textContent.trim() === 'Export data'; });
    if (target) { target.click(); return 'clicked'; }
    return 'not found';
})()
""")
    time.sleep(1.5)

    # 7 — Select "Summarized data" radio button
    print("  Selecting Summarized data...")
    _run_js("""
(function() {
    var radios = Array.from(document.querySelectorAll('input[type="radio"]'));
    for (var r of radios) {
        var label = r.closest('label') || document.querySelector('label[for="' + r.id + '"]');
        if (label && label.textContent.includes('Summarized')) { r.click(); return 'clicked summarized'; }
    }
    // Fallback: click the card container
    var cards = Array.from(document.querySelectorAll('[class*="option"], [class*="card"]'));
    var c = cards.find(function(el) { return el.textContent.includes('Summarized data'); });
    if (c) { c.click(); return 'clicked card'; }
    return 'not found';
})()
""")
    time.sleep(0.5)

    # 8 — Change file format dropdown to CSV
    print("  Setting format to CSV...")
    _run_js("""
(function() {
    var sel = document.querySelector('select');
    if (!sel) return 'no select found';
    for (var i = 0; i < sel.options.length; i++) {
        if (sel.options[i].text.toLowerCase().includes('csv')) {
            sel.selectedIndex = i;
            sel.dispatchEvent(new Event('change', {bubbles: true}));
            return 'selected: ' + sel.options[i].text;
        }
    }
    return 'csv option not found';
})()
""")
    time.sleep(0.5)

    # 9 — Click the Export button
    print("  Clicking Export button...")
    _run_js("""
(function() {
    var btns = Array.from(document.querySelectorAll('button'));
    var exp = btns.find(function(b) {
        var t = b.textContent.trim();
        return t === 'Export' || t === 'Export data';
    });
    if (exp) { exp.click(); return 'clicked export'; }
    return 'button not found';
})()
""")
    print("  Export triggered — waiting for download...")
    return True


def wait_for_csv(timeout_minutes: int = 10):
    """
    Polls ~/Downloads every 3 seconds for a new data.csv.
    Ignores files older than 2 minutes at the start of polling.
    Returns Path on success, None on timeout.
    """
    deadline   = time.time() + timeout_minutes * 60
    start_time = time.time()

    print(f"\nWaiting up to {timeout_minutes} min for data.csv in ~/Downloads …")
    spinner = ["|", "/", "-", "\\"]
    i = 0

    while time.time() < deadline:
        for f in DOWNLOADS.glob("data*.csv"):
            age = time.time() - f.stat().st_mtime
            if age < (time.time() - start_time) + 5:   # file appeared after we started
                print(f"\n  ✓ Download detected: {f.name}")
                return f
        elapsed = int(time.time() - start_time)
        print(f"  {spinner[i % 4]}  Waiting… {elapsed}s elapsed", end="\r")
        i += 1
        time.sleep(3)

    print("\n  ✗ Timed out. Run with --upload-only once the file is in ~/Downloads.")
    return None


def archive_csv(src: Path, ws: date, we: date) -> Path:
    ARCHIVE_DIR.mkdir(exist_ok=True)
    dest = ARCHIVE_DIR / f"jack_archer_{fmt_s(ws)}_to_{fmt_s(we)}.csv"
    shutil.copy2(src, dest)
    return dest


# ── Google Sheets ─────────────────────────────────────────────────────────────
def get_google_creds(config: dict):
    from google.oauth2.service_account import Credentials as SA
    from google.oauth2.credentials import Credentials as UC
    from google_auth_oauthlib.flow import InstalledAppFlow
    from google.auth.transport.requests import Request

    SCOPES  = ["https://www.googleapis.com/auth/spreadsheets"]
    sa_file = BASE_DIR / "google_service_account.json"
    oa_file = BASE_DIR / "google_oauth_credentials.json"

    if sa_file.exists():
        return SA.from_service_account_file(str(sa_file), scopes=SCOPES)

    creds = None
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

    if not oa_file.exists():
        raise FileNotFoundError(
            "No Google credentials found.\n"
            "Save google_oauth_credentials.json in this folder and try again."
        )
    flow  = InstalledAppFlow.from_client_secrets_file(str(oa_file), SCOPES)
    creds = flow.run_local_server(port=0)
    with open(TOKEN_CACHE, "wb") as f:
        pickle.dump(creds, f)
    print("  ✓ Google auth saved — won't need to log in again.")
    return creds


def upload_to_sheets(csv_path: Path, config: dict, ws: date, we: date, dry_run=False) -> int:
    """
    Append new rows to the sheet, deduplicating by Order ID.
    Prepends a 'Week' column (e.g. '2026-03-13 → 2026-03-19').
    Returns number of NEW rows appended.
    """
    import pandas as pd
    from googleapiclient.discovery import build

    sheet_id  = config["google_sheets"]["spreadsheet_id"]
    tab       = config["google_sheets"]["tab_name"]
    dedup_col = config["google_sheets"].get("dedup_column", "Order ID")
    wlabel    = week_label(ws, we)

    # Read CSV
    df = pd.read_csv(csv_path, encoding="utf-8-sig")
    print(f"  CSV loaded: {len(df):,} rows × {len(df.columns)} columns")

    # Add week label column
    df.insert(0, "Week", wlabel)

    # Connect to Sheets
    creds   = get_google_creds(config)
    svc     = build("sheets", "v4", credentials=creds)
    sheets  = svc.spreadsheets()

    # Fetch existing data for dedup
    resp = sheets.values().get(
        spreadsheetId=sheet_id,
        range=f"{tab}!A1:ZZ"
    ).execute()
    existing_rows = resp.get("values", [])

    existing_ids: set = set()
    sheet_is_empty = len(existing_rows) == 0

    if not sheet_is_empty and dedup_col in existing_rows[0]:
        col_idx = existing_rows[0].index(dedup_col)
        for row in existing_rows[1:]:
            if len(row) > col_idx and row[col_idx]:
                existing_ids.add(str(row[col_idx]).strip())
        print(f"  Found {len(existing_ids):,} existing '{dedup_col}' values in sheet")
    elif not sheet_is_empty:
        print(f"  Note: '{dedup_col}' column not found in sheet — will check after first upload")

    # Deduplicate
    if dedup_col in df.columns and existing_ids:
        before = len(df)
        df = df[~df[dedup_col].astype(str).str.strip().isin(existing_ids)]
        skipped = before - len(df)
        if skipped > 0:
            print(f"  Skipped {skipped:,} duplicates already in the sheet")

    if df.empty:
        print("  ✓ Nothing new — sheet is already up to date for this week.")
        return 0

    print(f"  Appending {len(df):,} new rows to '{tab}' …")
    if dry_run:
        print("  [DRY RUN] No changes written.")
        print(df.head(3).to_string(index=False))
        return len(df)

    # Build rows
    rows = []
    if sheet_is_empty:
        rows.append(df.columns.tolist())  # Header only on first-ever write

    for _, row in df.iterrows():
        clean = []
        for v in row:
            try:
                if pd.isna(v):
                    clean.append("")
                    continue
            except (TypeError, ValueError):
                pass
            if isinstance(v, float) and v == int(v):
                clean.append(int(v))
            elif isinstance(v, (int, float, bool)):
                clean.append(v)
            else:
                clean.append(str(v))
        rows.append(clean)

    sheets.values().append(
        spreadsheetId=sheet_id,
        range=f"{tab}!A1",
        valueInputOption="USER_ENTERED",
        insertDataOption="INSERT_ROWS",
        body={"values": rows},
    ).execute()

    return len(df)


# ── Scheduling ────────────────────────────────────────────────────────────────
def install_schedule(config: dict):
    LOG_DIR.mkdir(exist_ok=True)
    weekday = config.get("schedule", {}).get("weekday", 7)   # 7 = Saturday
    hour    = config.get("schedule", {}).get("hour", 8)
    minute  = config.get("schedule", {}).get("minute", 0)

    plist = f"""<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN"
    "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>Label</key><string>{PLIST_LABEL}</string>
    <key>ProgramArguments</key>
    <array>
        <string>{sys.executable}</string>
        <string>{Path(__file__).resolve()}</string>
    </array>
    <key>StartCalendarInterval</key>
    <dict>
        <key>Weekday</key><integer>{weekday}</integer>
        <key>Hour</key><integer>{hour}</integer>
        <key>Minute</key><integer>{minute}</integer>
    </dict>
    <key>WorkingDirectory</key><string>{BASE_DIR}</string>
    <key>StandardOutPath</key><string>{LOG_DIR}/export.log</string>
    <key>StandardErrorPath</key><string>{LOG_DIR}/export_error.log</string>
    <key>RunAtLoad</key><false/>
</dict>
</plist>"""

    PLIST_PATH.parent.mkdir(parents=True, exist_ok=True)
    PLIST_PATH.write_text(plist)
    subprocess.run(["launchctl", "load", str(PLIST_PATH)], capture_output=True)

    days = {1:"Sunday", 2:"Monday", 3:"Tuesday", 4:"Wednesday",
            5:"Thursday", 6:"Friday", 7:"Saturday"}
    print(f"✓ Scheduled every {days.get(weekday, 'Saturday')} at {hour:02d}:{minute:02d}")
    print(f"  Logs → {LOG_DIR}/export.log")


def uninstall_schedule():
    if PLIST_PATH.exists():
        subprocess.run(["launchctl", "unload", str(PLIST_PATH)], capture_output=True)
        PLIST_PATH.unlink()
        print("✓ Schedule removed.")
    else:
        print("No schedule found.")


# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser(description="Jack Archer weekly Power BI → Sheets export")
    parser.add_argument("--dry-run",     action="store_true", help="Preview upload without writing")
    parser.add_argument("--upload-only", action="store_true", help="Skip Chrome, use latest ~/Downloads/data.csv")
    parser.add_argument("--schedule",    action="store_true", help="Install Saturday 8am launchd job")
    parser.add_argument("--unschedule",  action="store_true", help="Remove the scheduled job")
    args = parser.parse_args()

    if args.unschedule:
        uninstall_schedule()
        return

    config = load_config()

    if args.schedule:
        install_schedule(config)
        return

    ws, we = get_week_range()

    print("=" * 62)
    print("  Jack Archer Weekly Export")
    print(f"  Week   : {fmt(ws)}  →  {fmt(we)}")
    print(f"  Today  : {fmt(date.today())}")
    print("=" * 62)

    # ── Step 1: Get the CSV ────────────────────────────────────────────────────
    if args.upload_only:
        candidates = sorted(DOWNLOADS.glob("data*.csv"), key=lambda p: p.stat().st_mtime, reverse=True)
        if not candidates:
            print("\nNo data.csv in ~/Downloads. Export from Power BI first.")
            sys.exit(1)
        csv_path = candidates[0]
        print(f"\nUsing: {csv_path.name}  ({csv_path.stat().st_size // 1024} KB)")
    else:
        print(f"\n[1/3] Automating Chrome export ({fmt_s(ws)} → {fmt_s(we)})...")
        automate_chrome_export(config["powerbi"]["report_url"], we)
        csv_path = wait_for_csv(timeout_minutes=10)
        if csv_path is None:
            print("\nExport not detected. Try running with --upload-only after exporting manually.")
            sys.exit(1)

    # ── Step 2: Archive ───────────────────────────────────────────────────────
    print("\n[2/3] Archiving CSV …")
    archived = archive_csv(csv_path, ws, we)
    print(f"  Saved → downloads/{archived.name}")

    # ── Step 3: Upload with dedup ─────────────────────────────────────────────
    print("\n[3/3] Uploading to Google Sheets …")
    new_rows = upload_to_sheets(archived, config, ws, we, dry_run=args.dry_run)

    tag = "[DRY RUN] " if args.dry_run else ""
    print(f"\n{tag}✓ Done — {new_rows:,} new rows added.")
    if new_rows and not args.dry_run:
        sid = config["google_sheets"]["spreadsheet_id"]
        print(f"  https://docs.google.com/spreadsheets/d/{sid}")


if __name__ == "__main__":
    main()

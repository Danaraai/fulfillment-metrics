"""
Power BI → Google Sheets Exporter
Main entry point.

Usage:
  python3 main.py              # Run the export now
  python3 main.py --discover   # List all visuals on the report page
  python3 main.py --schedule   # Install weekly launchd job (macOS)
  python3 main.py --unschedule # Remove the weekly launchd job
"""

import asyncio
import argparse
import json
import os
import sys
import subprocess
from pathlib import Path
from datetime import datetime

CONFIG_FILE = Path(__file__).parent / "config.json"
PLIST_LABEL = "com.powerbi.jack-archer-export"
PLIST_PATH = Path.home() / "Library" / "LaunchAgents" / f"{PLIST_LABEL}.plist"


def load_config() -> dict:
    if not CONFIG_FILE.exists():
        print(f"ERROR: config.json not found at {CONFIG_FILE}")
        print("Run setup.sh first, or copy config.example.json to config.json")
        sys.exit(1)
    with open(CONFIG_FILE) as f:
        return json.load(f)


async def run_export(config: dict, discover: bool = False):
    """Run the full export pipeline: Power BI → CSV → Google Sheets."""
    from powerbi_exporter import export_visual_as_csv, _print_available_visuals
    from gsheet_helper import append_csv_to_sheet
    from playwright.async_api import async_playwright

    print("\n" + "=" * 60)
    print(f"  Power BI → Google Sheets Export")
    print(f"  {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 60)

    if discover:
        # Open browser, load report, list visuals then exit
        print("DISCOVER MODE: listing all visuals on the report page...")
        from playwright.async_api import async_playwright
        from powerbi_exporter import (
            SESSION_DIR, _handle_login_if_needed,
            _wait_for_report, _print_available_visuals
        )
        os.makedirs(SESSION_DIR, exist_ok=True)
        async with async_playwright() as p:
            browser = await p.chromium.launch_persistent_context(
                SESSION_DIR,
                headless=False,
                viewport={"width": 1920, "height": 1080},
                accept_downloads=True,
                ignore_https_errors=True,  # Bypass corporate SSL inspection CA
            )
            page = browser.pages[0] if browser.pages else await browser.new_page()
            try:
                await page.goto(config["powerbi"]["report_url"], wait_until="domcontentloaded", timeout=60_000)
            except Exception:
                pass
            await _handle_login_if_needed(page)
            await _wait_for_report(page)
            await _print_available_visuals(page)
            # Keep browser open 30s so user can see the results
            print("\nBrowser will close in 30 seconds...")
            await page.wait_for_timeout(30_000)
            await browser.close()
        return

    # Step 1: Export from Power BI
    print("\n[1/2] Exporting from Power BI...")
    csv_path = await export_visual_as_csv(config)
    print(f"  CSV saved: {csv_path}")

    # Step 2: Append to Google Sheets
    sheet_id = config["google_sheets"]["spreadsheet_id"]
    tab_name = config["google_sheets"]["tab_name"]

    if sheet_id == "YOUR_GOOGLE_SHEET_ID":
        print("\n⚠  Google Sheets not configured yet.")
        print("  1. Open config.json")
        print("  2. Set 'spreadsheet_id' to your sheet ID")
        print("  3. Set 'tab_name' to your tab name")
        print(f"\n  CSV is available at: {csv_path}")
        return

    print(f"\n[2/2] Uploading to Google Sheets tab '{tab_name}'...")
    rows = append_csv_to_sheet(csv_path, config)
    print(f"  Done. {rows} rows appended.")

    print("\n✓ Export complete!")


def install_weekly_schedule(config: dict):
    """Install a launchd plist to run the export every Monday at 8am."""
    script_path = Path(__file__).resolve()
    python_path = sys.executable
    log_dir = Path(__file__).parent / "logs"
    log_dir.mkdir(exist_ok=True)

    # Get the day of week from config (default: Monday = 2 in launchd)
    weekday = config.get("schedule", {}).get("weekday", 2)  # 1=Sun, 2=Mon, ... 7=Sat
    hour = config.get("schedule", {}).get("hour", 8)
    minute = config.get("schedule", {}).get("minute", 0)

    plist_content = f"""<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN"
    "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>Label</key>
    <string>{PLIST_LABEL}</string>

    <key>ProgramArguments</key>
    <array>
        <string>{python_path}</string>
        <string>{script_path}</string>
    </array>

    <key>StartCalendarInterval</key>
    <dict>
        <key>Weekday</key>
        <integer>{weekday}</integer>
        <key>Hour</key>
        <integer>{hour}</integer>
        <key>Minute</key>
        <integer>{minute}</integer>
    </dict>

    <key>WorkingDirectory</key>
    <string>{script_path.parent}</string>

    <key>StandardOutPath</key>
    <string>{log_dir}/export.log</string>

    <key>StandardErrorPath</key>
    <string>{log_dir}/export_error.log</string>

    <key>RunAtLoad</key>
    <false/>
</dict>
</plist>
"""

    PLIST_PATH.parent.mkdir(parents=True, exist_ok=True)
    PLIST_PATH.write_text(plist_content)
    print(f"Wrote plist: {PLIST_PATH}")

    # Load the agent
    result = subprocess.run(
        ["launchctl", "load", str(PLIST_PATH)],
        capture_output=True, text=True
    )
    if result.returncode == 0:
        day_names = {1: "Sunday", 2: "Monday", 3: "Tuesday", 4: "Wednesday",
                     5: "Thursday", 6: "Friday", 7: "Saturday"}
        print(f"\n✓ Scheduled! Will run every {day_names.get(weekday, 'Monday')} at {hour:02d}:{minute:02d}")
        print(f"  Logs: {log_dir}/export.log")
        print(f"  To remove: python3 main.py --unschedule")
    else:
        print(f"ERROR loading plist: {result.stderr}")
        print(f"Try manually: launchctl load {PLIST_PATH}")


def uninstall_schedule():
    """Remove the launchd job."""
    if PLIST_PATH.exists():
        subprocess.run(["launchctl", "unload", str(PLIST_PATH)], capture_output=True)
        PLIST_PATH.unlink()
        print(f"✓ Schedule removed.")
    else:
        print("No schedule found.")


def main():
    parser = argparse.ArgumentParser(description="Power BI → Google Sheets automated export")
    parser.add_argument(
        "--discover", action="store_true",
        help="Open report and list all visual titles (use to find your visual name)"
    )
    parser.add_argument(
        "--schedule", action="store_true",
        help="Install weekly macOS launchd schedule"
    )
    parser.add_argument(
        "--unschedule", action="store_true",
        help="Remove the weekly schedule"
    )
    args = parser.parse_args()

    config = load_config()

    if args.unschedule:
        uninstall_schedule()
        return

    if args.schedule:
        install_weekly_schedule(config)
        return

    # Run the export (with or without discover mode)
    asyncio.run(run_export(config, discover=args.discover))


if __name__ == "__main__":
    main()

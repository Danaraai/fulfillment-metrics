"""
Power BI Chrome Extension Exporter
Uses Claude-in-Chrome MCP to automate the export via your real Chrome browser.
No test browser, no Playwright, no corporate SSL issues.
"""

import asyncio
import json
import os
import shutil
import glob
from datetime import datetime
from pathlib import Path

DOWNLOADS_DIR = str(Path(__file__).parent / "downloads")


async def export_via_chrome(config: dict) -> str:
    """
    Uses Claude-in-Chrome MCP tools to:
    1. Open the Power BI report in real Chrome
    2. Click through the export dialog (More Options → Export data → CSV)
    3. Wait for download
    4. Return path to the downloaded CSV

    Note: This function is called from the MCP-enabled context (Claude Code).
    When running as a scheduled script, use run_export_headless() instead.
    """
    raise NotImplementedError(
        "chrome_exporter.py is used directly by Claude Code's Chrome MCP tools.\n"
        "To run the full pipeline manually, use:\n"
        "  python3 run_now.py"
    )


def find_latest_download(prefix: str = "data", max_age_seconds: int = 120) -> str | None:
    """
    Find the most recently downloaded CSV from Power BI in ~/Downloads.
    Returns the file path or None if not found within the time window.
    """
    downloads = Path.home() / "Downloads"
    candidates = []

    for pattern in [f"{prefix}.csv", "*.csv"]:
        for f in downloads.glob(pattern):
            age = (datetime.now().timestamp() - f.stat().st_mtime)
            if age <= max_age_seconds:
                candidates.append((f.stat().st_mtime, str(f)))

    if not candidates:
        return None

    candidates.sort(reverse=True)
    return candidates[0][1]


def copy_to_archive(src_path: str, merchant: str = "jack_archer") -> str:
    """Copy downloaded CSV to the local archive with a dated filename."""
    os.makedirs(DOWNLOADS_DIR, exist_ok=True)
    dated_name = f"{merchant}_{datetime.now().strftime('%Y%m%d_%H%M')}.csv"
    dest = os.path.join(DOWNLOADS_DIR, dated_name)
    shutil.copy(src_path, dest)
    return dest


def upload_csv_to_sheets(csv_path: str, config: dict) -> int:
    """Upload CSV to Google Sheets and return row count."""
    from gsheet_helper import append_csv_to_sheet
    return append_csv_to_sheet(csv_path, config, add_timestamp=True)

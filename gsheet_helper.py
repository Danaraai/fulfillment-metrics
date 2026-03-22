"""
Google Sheets Helper
Appends rows from a CSV file to a specified Google Sheet tab.
Supports both Service Account auth (recommended for automation)
and OAuth2 user credentials.
"""

import os
import json
import csv
from pathlib import Path
from datetime import datetime

from google.oauth2.service_account import Credentials as SACredentials
from google.oauth2.credentials import Credentials as UserCredentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import pickle
import pandas as pd

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
TOKEN_CACHE = str(Path(__file__).parent / "google_token.pickle")
OAUTH_CREDS_FILE = str(Path(__file__).parent / "google_oauth_credentials.json")
SERVICE_ACCOUNT_FILE = str(Path(__file__).parent / "google_service_account.json")


def get_google_credentials(config: dict):
    """
    Returns Google API credentials.
    Priority: service account > cached OAuth token > new OAuth flow.
    """
    # Option 1: Service Account (best for automation)
    if os.path.exists(SERVICE_ACCOUNT_FILE):
        print("Using Google Service Account credentials")
        creds = SACredentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
        return creds

    # Option 2: Cached OAuth token from previous run
    creds = None
    if os.path.exists(TOKEN_CACHE):
        with open(TOKEN_CACHE, "rb") as f:
            creds = pickle.load(f)

    if creds and creds.valid:
        print("Using cached Google OAuth token")
        return creds

    # Refresh if expired
    if creds and creds.expired and creds.refresh_token:
        print("Refreshing Google OAuth token...")
        creds.refresh(Request())
        with open(TOKEN_CACHE, "wb") as f:
            pickle.dump(creds, f)
        return creds

    # Option 3: New OAuth flow (opens browser once)
    if not os.path.exists(OAUTH_CREDS_FILE):
        raise FileNotFoundError(
            "No Google credentials found. Please follow setup instructions:\n"
            "  • Service Account: save credentials as 'google_service_account.json'\n"
            "  • OAuth: save credentials as 'google_oauth_credentials.json'\n"
            "See SETUP.md for step-by-step instructions."
        )

    print("Starting Google OAuth flow (browser will open once)...")
    flow = InstalledAppFlow.from_client_secrets_file(OAUTH_CREDS_FILE, SCOPES)
    creds = flow.run_local_server(port=0)
    with open(TOKEN_CACHE, "wb") as f:
        pickle.dump(creds, f)
    print("Google credentials saved. Won't need to authenticate again.")
    return creds


def append_csv_to_sheet(csv_path: str, config: dict, add_timestamp: bool = True) -> int:
    """
    Reads a CSV file and appends its rows to the configured Google Sheet tab.

    Args:
        csv_path: Path to the downloaded CSV file
        config: dict loaded from config.json
        add_timestamp: Whether to add a 'Exported At' column

    Returns:
        Number of rows appended
    """
    sheet_id = config["google_sheets"]["spreadsheet_id"]
    tab_name = config["google_sheets"]["tab_name"]
    skip_header_if_sheet_has_data = config["google_sheets"].get("skip_header_if_data_exists", True)
    clear_before_append = config["google_sheets"].get("clear_tab_before_append", False)

    print(f"Reading CSV: {csv_path}")
    df = pd.read_csv(csv_path, encoding="utf-8-sig")  # utf-8-sig handles BOM from Power BI

    if add_timestamp:
        df.insert(0, "Exported At", datetime.now().strftime("%Y-%m-%d %H:%M"))

    print(f"  Loaded {len(df)} rows x {len(df.columns)} columns")

    creds = get_google_credentials(config)
    service = build("sheets", "v4", credentials=creds)
    sheets = service.spreadsheets()

    range_name = f"{tab_name}!A1"

    if clear_before_append:
        # Clear the tab first, then write everything including header
        print(f"Clearing tab '{tab_name}'...")
        sheets.values().clear(spreadsheetId=sheet_id, range=tab_name).execute()
        rows_to_write = [df.columns.tolist()] + df.values.tolist()
        _write_rows(sheets, sheet_id, range_name, rows_to_write)
        print(f"Wrote {len(df)} rows (with header) to '{tab_name}'")

    else:
        # Check if sheet already has data (to decide on header)
        existing = sheets.values().get(
            spreadsheetId=sheet_id,
            range=f"{tab_name}!A1:A2"
        ).execute()
        existing_values = existing.get("values", [])
        sheet_is_empty = len(existing_values) == 0

        rows_to_write = []
        if sheet_is_empty or not skip_header_if_sheet_has_data:
            rows_to_write.append(df.columns.tolist())  # Add header

        rows_to_write += _df_to_values(df)

        # Append after existing data
        append_range = f"{tab_name}!A1"
        result = sheets.values().append(
            spreadsheetId=sheet_id,
            range=append_range,
            valueInputOption="USER_ENTERED",
            insertDataOption="INSERT_ROWS",
            body={"values": rows_to_write},
        ).execute()

        rows_appended = result.get("updates", {}).get("updatedRows", len(rows_to_write))
        print(f"Appended {rows_appended} rows to '{tab_name}' in sheet '{sheet_id}'")

    return len(df)


def _df_to_values(df: pd.DataFrame) -> list:
    """Convert DataFrame rows to list of lists, handling NaN and types."""
    rows = []
    for _, row in df.iterrows():
        clean_row = []
        for val in row:
            if pd.isna(val) if not isinstance(val, str) else False:
                clean_row.append("")
            elif isinstance(val, float) and val == int(val):
                clean_row.append(int(val))
            else:
                clean_row.append(str(val) if not isinstance(val, (int, float, bool)) else val)
        rows.append(clean_row)
    return rows


def _write_rows(sheets, sheet_id: str, range_name: str, rows: list):
    """Write rows to a sheet range."""
    sheets.values().update(
        spreadsheetId=sheet_id,
        range=range_name,
        valueInputOption="USER_ENTERED",
        body={"values": rows},
    ).execute()

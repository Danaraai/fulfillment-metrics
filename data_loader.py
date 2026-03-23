"""
data_loader.py  —  Reads data from Google Sheets for the Streamlit dashboard.

Auth:
  - Locally: uses google_token.pickle (OAuth)
  - Streamlit Cloud: uses st.secrets["gcp_service_account"] (Service Account JSON)
"""

import json
import os
import pickle
import warnings
from datetime import datetime

import pandas as pd
import streamlit as st

warnings.filterwarnings("ignore")

SHEET_ID = "1tuo7knxTvOR3snd_u1AnnW1iiN9l1-TaOLi-YaVqLM0"
SCOPES   = ["https://www.googleapis.com/auth/spreadsheets.readonly"]


# ── Auth ──────────────────────────────────────────────────────────────────────

def _get_creds():
    """Return Google API credentials (service account on cloud, OAuth locally)."""
    # Streamlit Cloud: service account stored in secrets
    if hasattr(st, "secrets") and "gcp_service_account" in st.secrets:
        from google.oauth2.service_account import Credentials
        info = json.loads(st.secrets["gcp_service_account"])
        return Credentials.from_service_account_info(info, scopes=SCOPES)

    # Local dev: cached OAuth token
    token_path = os.path.join(os.path.dirname(__file__), "google_token.pickle")
    if os.path.exists(token_path):
        with open(token_path, "rb") as f:
            creds = pickle.load(f)
        if creds and creds.valid:
            return creds
        if creds and creds.expired and creds.refresh_token:
            from google.auth.transport.requests import Request
            creds.refresh(Request())
            return creds

    raise RuntimeError(
        "No Google credentials found.\n"
        "  • On Streamlit Cloud: add 'gcp_service_account' to secrets.\n"
        "  • Locally: run the OAuth flow once to create google_token.pickle."
    )


def _service(creds):
    from googleapiclient.discovery import build
    return build("sheets", "v4", credentials=creds)


# ── Sheet readers ─────────────────────────────────────────────────────────────

def _read_tab(svc, tab_name: str) -> pd.DataFrame:
    """Read any tab into a DataFrame, parsing dates and numerics."""
    try:
        result = svc.spreadsheets().values().get(
            spreadsheetId=SHEET_ID,
            range=f"{tab_name}!A1:AO"   # wide enough for all columns
        ).execute()
    except Exception:
        return pd.DataFrame()

    values = result.get("values", [])
    if not values:
        return pd.DataFrame()

    headers = values[0]
    rows    = values[1:]
    rows    = [r + [""] * (len(headers) - len(r)) for r in rows]
    df      = pd.DataFrame(rows, columns=headers)

    # Parse Transaction Date
    if "Transaction Date" in df.columns:
        df["Transaction Date"] = pd.to_datetime(
            df["Transaction Date"], errors="coerce", infer_datetime_format=True
        )
        df = df.dropna(subset=["Transaction Date"])

    # Parse numeric cost columns
    for col in ["Original Invoice", "Fulfillment without Surcharge",
                "Surcharge Applied", "WMS Fuel Surcharge",
                "Delivery Area Surcharge", "Residential Area Surcharge",
                "Address Correction Fee", "Other Order Fee", "Insurance Amount",
                "Actual Weight (Oz)", "Dim Weight(Oz)", "Billable Weight(Oz)",
                "Transit Time (Days)", "Length", "Width", "Height"]:
        if col in df.columns:
            # Strip currency symbols / commas before parsing (handles "$5.43" or "1,234.56")
            df[col] = (
                df[col].astype(str)
                       .str.replace(r"[\$,]", "", regex=True)
                       .pipe(pd.to_numeric, errors="coerce")
                       .fillna(0)
            )

    return df


@st.cache_data(ttl=3600, show_spinner="Loading shipping data…")
def load_export() -> pd.DataFrame:
    """Load Export + Old Shipments tabs, concatenated and deduplicated."""
    creds = _get_creds()
    svc   = _service(creds)

    new_df = _read_tab(svc, "Export")
    old_df = _read_tab(svc, "Old Shipments")

    if old_df.empty and new_df.empty:
        return pd.DataFrame()
    if old_df.empty:
        return new_df
    if new_df.empty:
        return old_df

    combined = pd.concat([old_df, new_df], ignore_index=True)

    # Deduplicate on OrderID + Transaction Date (keep last = newest export wins)
    if "OrderID" in combined.columns and "Transaction Date" in combined.columns:
        combined = (
            combined
            .sort_values("Transaction Date")
            .drop_duplicates(subset=["OrderID"], keep="last")
            .reset_index(drop=True)
        )

    return combined


@st.cache_data(ttl=3600, show_spinner="Loading labor hours…")
def load_labor_hours() -> pd.DataFrame:
    """
    Load labor hours from 'Labor Hours 2025' and 'Labor Hours 2026' tabs.
    Columns: Date | Total Hours | Emp Hours | Temp Hours |
             Outbound Hours | Inbound Hours | Headcount
    """
    creds = _get_creds()
    svc   = _service(creds)

    frames = []
    for tab in ("Labor Hours 2025", "Labor Hours 2026"):
        try:
            result = svc.spreadsheets().values().get(
                spreadsheetId=SHEET_ID,
                range=f"{tab}!A1:G"
            ).execute()
        except Exception:
            continue

        values = result.get("values", [])
        if len(values) < 2:
            continue

        headers = values[0]
        rows    = values[1:]
        rows    = [r + [""] * (len(headers) - len(r)) for r in rows]
        frame   = pd.DataFrame(rows, columns=headers)

        frame["Date"] = pd.to_datetime(frame["Date"], errors="coerce",
                                       infer_datetime_format=True)
        frame = frame.dropna(subset=["Date"])

        for col in ["Total Hours", "Emp Hours", "Temp Hours",
                    "Outbound Hours", "Inbound Hours", "Headcount"]:
            if col in frame.columns:
                frame[col] = pd.to_numeric(frame[col], errors="coerce").fillna(0)

        frames.append(frame)

    if not frames:
        return pd.DataFrame()

    combined = pd.concat(frames, ignore_index=True)
    combined = (
        combined
        .sort_values("Date")
        .drop_duplicates(subset=["Date"], keep="last")
        .reset_index(drop=True)
    )
    return combined


@st.cache_data(ttl=3600, show_spinner="Loading daily metrics…")
def load_daily_metrics() -> pd.DataFrame:
    """
    Load the pre-calculated Daily Metrics tab.
    Columns: Week | Daily orders | Labor Hours Outbound | Labor hours total |
             Labor cost per hour ($/hr) | Week | OPLH |
             Total Labor Cost Per Order ($/order) |
             Outbound Labor cost per order ($/order) |
             Packaging cost per order | Month
    """
    creds = _get_creds()
    svc   = _service(creds)
    try:
        result = svc.spreadsheets().values().get(
            spreadsheetId=SHEET_ID,
            range="Daily Metrics!A1:K"
        ).execute()
    except Exception:
        return pd.DataFrame()

    values = result.get("values", [])
    if len(values) < 2:
        return pd.DataFrame()

    headers = values[0]
    rows    = values[1:]
    rows    = [r + [""] * (len(headers) - len(r)) for r in rows]
    df      = pd.DataFrame(rows, columns=headers)

    # Parse date column (column A = "Week" = daily date)
    date_col = df.columns[0]
    df[date_col] = pd.to_datetime(df[date_col], errors="coerce", infer_datetime_format=True)
    df = df.dropna(subset=[date_col])
    df = df.rename(columns={date_col: "Date"})

    # Strip $ and parse all numeric/currency columns
    for col in df.columns:
        if col == "Date":
            continue
        df[col] = (
            df[col].astype(str)
                   .str.replace(r"[\$,]", "", regex=True)
                   .pipe(pd.to_numeric, errors="coerce")
        )

    return df


@st.cache_data(ttl=3600)
def load_all():
    """Return (export_df, labor_df) together so both cache together."""
    return load_export(), load_labor_hours()

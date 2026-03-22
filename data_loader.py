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

@st.cache_data(ttl=3600, show_spinner="Loading shipping data…")
def load_export() -> pd.DataFrame:
    """Load the Export tab (Power BI shipping data)."""
    creds = _get_creds()
    svc   = _service(creds)
    result = svc.spreadsheets().values().get(
        spreadsheetId=SHEET_ID,
        range="Export!A1:Z"
    ).execute()
    values = result.get("values", [])
    if not values:
        return pd.DataFrame()

    headers = values[0]
    rows    = values[1:]
    # Pad short rows
    rows = [r + [""] * (len(headers) - len(r)) for r in rows]
    df   = pd.DataFrame(rows, columns=headers)

    # Parse date
    df["Transaction Date"] = pd.to_datetime(
        df["Transaction Date"], errors="coerce", infer_datetime_format=True
    )
    # Parse numeric cost columns
    for col in ["Original Invoice", "Fulfillment without Surcharge",
                "Surcharge Applied", "WMS Fuel Surcharge",
                "Delivery Area Surcharge", "Residential Area Surcharge",
                "Address Correction Fee", "Other Order Fee"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    df = df.dropna(subset=["Transaction Date"])
    return df


@st.cache_data(ttl=3600, show_spinner="Loading labor hours…")
def load_daily_metrics() -> pd.DataFrame:
    """Load the Daily Metrics tab (formula-driven KPIs)."""
    creds = _get_creds()
    svc   = _service(creds)
    try:
        result = svc.spreadsheets().values().get(
            spreadsheetId=SHEET_ID,
            range="Daily Metrics!A1:K"
        ).execute()
    except Exception:
        return pd.DataFrame()  # Tab doesn't exist yet — graceful fallback

    values = result.get("values", [])
    if len(values) < 2:
        return pd.DataFrame()

    headers = values[0]
    rows    = values[1:]
    rows    = [r + [""] * (len(headers) - len(r)) for r in rows]
    df      = pd.DataFrame(rows, columns=headers)

    df["Date"] = pd.to_datetime(df["Date"], errors="coerce", infer_datetime_format=True)
    df = df.dropna(subset=["Date"])

    numeric_cols = ["Daily Orders", "Labor Hours Outbound", "Labor Hours Total",
                    "Labor Cost/hr", "OPLH", "Total Labor Cost Per Order",
                    "Outbound Labor Cost Per Order", "Packaging Cost Per Order"]
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    return df


@st.cache_data(ttl=3600)
def load_all():
    """Return (export_df, metrics_df) together so both cache together."""
    return load_export(), load_daily_metrics()

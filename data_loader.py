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

SHEET_ID            = "1tuo7knxTvOR3snd_u1AnnW1iiN9l1-TaOLi-YaVqLM0"
COMPARISON_SHEET_ID = "1DyOwWoFSiAmtf5pDuaOZKCe9wPXwwkBZnXX5BsxMacg"
SCOPES              = ["https://www.googleapis.com/auth/spreadsheets.readonly"]


# ── Auth ──────────────────────────────────────────────────────────────────────

def _get_creds():
    """Return Google API credentials (service account on cloud, OAuth locally)."""
    # Streamlit Cloud: service account stored in secrets
    # Wrap in try/except — accessing st.secrets raises if secrets.toml is missing locally
    try:
        if "gcp_service_account" in st.secrets:
            from google.oauth2.service_account import Credentials
            info = json.loads(st.secrets["gcp_service_account"])
            return Credentials.from_service_account_info(info, scopes=SCOPES)
    except Exception:
        pass  # No secrets.toml locally — fall through to OAuth token

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
    Load labor hours from the 'Labor Hours 2025' and 'Labor Hours 2026' tabs
    in the Export PowerBI spreadsheet (kept current by sync_labor_hours.py).
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

        # Column A is always the date regardless of header name ("Date", "Week", etc.)
        date_col = frame.columns[0]
        frame[date_col] = pd.to_datetime(frame[date_col], errors="coerce",
                                         infer_datetime_format=True)
        frame = frame.dropna(subset=[date_col])
        frame = frame.rename(columns={date_col: "Date"})
        frame["Date"] = pd.to_datetime(frame["Date"])

        for col in ["Total Hours", "Emp Hours", "Temp Hours",
                    "Outbound Hours", "Inbound Hours", "Headcount"]:
            if col in frame.columns:
                frame[col] = pd.to_numeric(frame[col], errors="coerce").fillna(0)

        frames.append(frame)

    if not frames:
        return pd.DataFrame()

    combined = (
        pd.concat(frames, ignore_index=True)
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

    # Parse date column (column A = "Week" = daily date).
    # Google Sheets API may return either "2/20/2026" (formatted) or a serial
    # number like "46432" (days since 1899-12-30).  Handle both.
    date_col = df.columns[0]

    def _parse_date_cell(v):
        s = str(v).strip()
        if not s or s.lower() in ("nan", "none", ""):
            return pd.NaT
        # Try standard datetime parsing first
        parsed = pd.to_datetime(s, errors="coerce", infer_datetime_format=True)
        if pd.notna(parsed):
            return parsed
        # Fallback: Google Sheets serial number (days since 1899-12-30)
        try:
            serial = float(s)
            return pd.Timestamp("1899-12-30") + pd.to_timedelta(serial, unit="D")
        except (ValueError, TypeError):
            return pd.NaT

    # Rename column A to "Date" by position (not by name) to avoid renaming
    # column F which is also called "Week" in the Daily Metrics tab.
    df.iloc[:, 0] = df.iloc[:, 0].apply(_parse_date_cell)
    df = df[df.iloc[:, 0].notna()].copy()
    new_cols = list(df.columns)
    new_cols[0] = "Date"
    df.columns = new_cols
    # Explicitly cast to datetime64 — iloc assignment via apply() returns object dtype
    df["Date"] = pd.to_datetime(df["Date"])

    # Strip $ and parse all numeric/currency columns (skip Date and the
    # weekly-rollup "Week" column in position F which has date strings)
    for i, col in enumerate(df.columns):
        if i == 0:   # Date column — already parsed
            continue
        df[col] = (
            df[col].astype(str)
                   .str.replace(r"[\$,]", "", regex=True)
                   .pipe(pd.to_numeric, errors="coerce")
        )

    return df


@st.cache_data(ttl=3600, show_spinner="Loading negotiation comparison data…")
def load_comparison() -> pd.DataFrame:
    """
    Load Enriched_Data tab from the Shipping Pre vs Post Negotiation sheet.
    Columns: OrderID, Transaction Date, Carrier, Zone Used, Billable Weight(Oz),
             Pre_Neg_Total_Rate, Current_Total, Savings_Per_Shipment, Rate_Lookup_Status, …
    """
    creds = _get_creds()
    svc   = _service(creds)

    try:
        result = svc.spreadsheets().values().get(
            spreadsheetId=COMPARISON_SHEET_ID,
            range="Enriched_Data!A1:AH"
        ).execute()
    except Exception as e:
        import warnings
        warnings.warn(f"Could not load comparison sheet: {e}")
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
            df["Transaction Date"], errors="coerce"
        )
        df = df.dropna(subset=["Transaction Date"])

    # Parse all numeric columns
    numeric_cols = [
        "Pre_Neg_Base_Rate", "Pre_Neg_Fuel_Surcharge", "Pre_Neg_Residential_Sur",
        "Pre_Neg_DAS_Surcharge", "Pre_Neg_Total_Rate", "Current_Total",
        "Savings_Per_Shipment", "Original Invoice", "Fulfillment without Surcharge",
        "Surcharge Applied", "WMS Fuel Surcharge", "Delivery Area Surcharge",
        "Residential Area Surcharge", "Billable Weight(Oz)", "Actual Weight (Oz)",
        "Dim Weight(Oz)",
    ]
    for col in numeric_cols:
        if col in df.columns:
            df[col] = (
                df[col].astype(str)
                       .str.replace(r"[\$,]", "", regex=True)
                       .pipe(pd.to_numeric, errors="coerce")
                       .fillna(0)
            )

    return df


@st.cache_data(ttl=3600)
def load_all():
    """Return (export_df, labor_df) together so both cache together."""
    return load_export(), load_labor_hours()

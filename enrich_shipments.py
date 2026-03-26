#!/usr/bin/env python3
"""
enrich_shipments.py — Shipping Pre vs Post Negotiation Comparison

Reads shipment data from the Export FollowUI Google Sheet,
looks up pre-negotiation rates from the Shipbob rate table,
computes savings per shipment, and writes results to a dedicated
comparison spreadsheet (created automatically on first run).

Usage:
    python3 enrich_shipments.py           # Run and write to Google Sheets
    python3 enrich_shipments.py --dry-run # Compute only, print summary, no writes

Pre-negotiation surcharges applied on top of base rate:
    Fuel Surcharge:                  20.5%  of base transportation charge
    Residential Surcharge:           $1.08  flat  (if Residential Area Surcharge > 0)
    Delivery Area Surcharge (Res):   $0.10  flat  (if Delivery Area Surcharge > 0)
"""

import argparse
import io
import json
import pickle
import sys
from pathlib import Path

import pandas as pd
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from google.oauth2.service_account import Credentials as SACredentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

# Reuse write helpers from this project
sys.path.insert(0, str(Path(__file__).parent))
from gsheet_helper import _df_to_values, _write_rows

# ── Constants ──────────────────────────────────────────────────────────────────

CONFIG_PATH     = Path(__file__).parent / "config.json"
EXPORT_SHEET_ID = "1tuo7knxTvOR3snd_u1AnnW1iiN9l1-TaOLi-YaVqLM0"

# This script needs both Sheets + Drive scopes (Drive is needed to download xlsx rate table)
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.readonly",
]
# Separate token cache so we don't invalidate the existing spreadsheets-only token
TOKEN_CACHE     = Path(__file__).parent / "google_token_enrich.pickle"
OAUTH_CREDS     = Path(__file__).parent / "google_oauth_credentials.json"
SERVICE_ACCOUNT = Path(__file__).parent / "google_service_account.json"

FUEL_SURCHARGE_PCT   = 0.205   # 20.5% applied to base transportation charge
RESIDENTIAL_SUR      = 1.08    # Residential surcharge (flat)
DAS_RESIDENTIAL_SUR  = 0.10    # Delivery Area Surcharge - Residential (flat)

# Exact columns to output — source columns first, then computed columns.
# Any column not present in the data is silently skipped.
OUTPUT_COLUMNS = [
    # Identity
    "OrderID", "Transaction Date", "TrackingId",
    # Current invoice breakdown
    "Fulfillment without Surcharge", "Surcharge Applied", "Original Invoice",
    "WMS Fuel Surcharge", "Delivery Area Surcharge", "Residential Area Surcharge",
    "Address Correction Fee", "Other Order Fee", "Insurance Amount",
    # Shipping details
    "Ship Option ID", "Carrier", "Carrier Service", "Zone Used",
    "Actual Weight (Oz)", "Dim Weight(Oz)", "Billable Weight(Oz)",
    # Destination
    "Zip Code", "City", "Destination Country",
    # Category
    "Order Category",
    # Pre-negotiation computed columns
    "Pre_Neg_Base_Rate", "Pre_Neg_Fuel_Surcharge", "Pre_Neg_Residential_Sur",
    "Pre_Neg_DAS_Surcharge", "Pre_Neg_Total_Rate",
    "Current_Total", "Savings_Per_Shipment", "Rate_Lookup_Status",
]


def select_output_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Return df with only the OUTPUT_COLUMNS that actually exist in df."""
    cols = [c for c in OUTPUT_COLUMNS if c in df.columns]
    missing = [c for c in OUTPUT_COLUMNS if c not in df.columns]
    if missing:
        print(f"  Note: {len(missing)} columns not found in data and will be skipped: {missing}")
    return df[cols]


# ── Config helpers ─────────────────────────────────────────────────────────────

def load_config() -> dict:
    with open(CONFIG_PATH) as f:
        return json.load(f)


def save_config(config: dict):
    with open(CONFIG_PATH, "w") as f:
        json.dump(config, f, indent=2)


def get_credentials():
    """
    Return Google credentials with Sheets + Drive scopes.
    Uses a dedicated token cache (google_token_enrich.pickle) so the existing
    spreadsheets-only token for the main export pipeline is not affected.
    First run will open a browser for OAuth consent.
    """
    if SERVICE_ACCOUNT.exists():
        print("  Using service account credentials")
        return SACredentials.from_service_account_file(str(SERVICE_ACCOUNT), scopes=SCOPES)

    creds = None
    if TOKEN_CACHE.exists():
        with open(TOKEN_CACHE, "rb") as f:
            creds = pickle.load(f)

    if creds and creds.valid:
        print("  Using cached OAuth token (google_token_enrich.pickle)")
        return creds

    if creds and creds.expired and creds.refresh_token:
        print("  Refreshing OAuth token...")
        creds.refresh(Request())
        with open(TOKEN_CACHE, "wb") as f:
            pickle.dump(creds, f)
        return creds

    if not OAUTH_CREDS.exists():
        raise FileNotFoundError(
            "No Google credentials found. Expected 'google_oauth_credentials.json' "
            "in the project directory."
        )

    print("  Starting OAuth flow (browser will open once to grant Sheets + Drive access)...")
    flow = InstalledAppFlow.from_client_secrets_file(str(OAUTH_CREDS), SCOPES)
    creds = flow.run_local_server(port=0)
    with open(TOKEN_CACHE, "wb") as f:
        pickle.dump(creds, f)
    print(f"  Credentials saved to {TOKEN_CACHE.name}")
    return creds


# ── Google Sheets helpers ──────────────────────────────────────────────────────

def read_tab(service, sheet_id: str, tab_name: str, col_range: str = "A1:AZ") -> pd.DataFrame:
    """Read a sheet tab into a DataFrame. First row = headers."""
    try:
        result = service.spreadsheets().values().get(
            spreadsheetId=sheet_id,
            range=f"'{tab_name}'!{col_range}"
        ).execute()
    except Exception as e:
        print(f"  Warning: Could not read tab '{tab_name}': {e}")
        return pd.DataFrame()

    values = result.get("values", [])
    if not values:
        return pd.DataFrame()

    headers = values[0]
    rows    = values[1:]
    rows    = [r + [""] * (len(headers) - len(r)) for r in rows]
    df      = pd.DataFrame(rows, columns=headers)
    return df


def clear_and_write_tab(service, sheet_id: str, tab_name: str, df: pd.DataFrame):
    """Clear a tab then write DataFrame as static values (header + data rows)."""
    sheets = service.spreadsheets()
    sheets.values().clear(spreadsheetId=sheet_id, range=tab_name).execute()
    rows = [df.columns.tolist()] + _df_to_values(df)
    _write_rows(sheets, sheet_id, f"{tab_name}!A1", rows)
    print(f"  Wrote {len(df):,} rows to '{tab_name}'")


def create_or_get_comparison_sheet(service, config: dict) -> str:
    """
    Return the comparison spreadsheet ID.
    Creates it on first run and persists the ID to config.json.
    On subsequent runs reuses the existing sheet (clears + rewrites tabs).
    """
    sheet_id = config.get("comparison_sheet_id", "").strip()
    if sheet_id:
        try:
            service.spreadsheets().get(spreadsheetId=sheet_id).execute()
            print(f"  Using existing sheet: https://docs.google.com/spreadsheets/d/{sheet_id}")
            return sheet_id
        except Exception:
            print("  Stored comparison sheet not found — creating a new one...")

    print("  Creating new 'Shipping Pre vs Post Negotiation' spreadsheet...")
    new_sheet = service.spreadsheets().create(body={
        "properties": {"title": "Shipping Pre vs Post Negotiation"},
        "sheets": [
            {"properties": {"title": "Enriched_Data", "index": 0}},
            {"properties": {"title": "Summary",        "index": 1}},
        ]
    }).execute()

    sheet_id = new_sheet["spreadsheetId"]
    config["comparison_sheet_id"] = sheet_id
    save_config(config)
    print(f"  Created: https://docs.google.com/spreadsheets/d/{sheet_id}")
    return sheet_id


# ── Rate Table ─────────────────────────────────────────────────────────────────

def _parse_rate_raw_rows(raw_rows: list) -> pd.DataFrame:
    """
    Parse raw row data (list of lists) from the Shipbob rate table into a tidy DataFrame.

    Expected structure:
      Row 0  (ShipbobZone): zone number merged headers  — SKIP
      Row 1  (ZoneName):    zone numbers 1-10           — SKIP
      Row 2  (col headers): WeightOuncesStart | WeightOuncesEnd | Currency |
                            First | FlatR | First | FlatR | ... (one pair per zone 1-10)
                            — SKIP (we reconstruct column names from position)
      Row 3+: data rows with actual rates

    Column layout (0-indexed):
      0  = WeightOuncesStart
      1  = WeightOuncesEnd
      2  = Currency
      3  = Zone 1 First rate  ← "First" is the standard rate we use
      4  = Zone 1 FlatR
      5  = Zone 2 First rate
      ...  3 + (zone-1)*2 = Zone N First rate

    Returns tidy DataFrame: WeightOuncesStart | WeightOuncesEnd | Zone | Pre_Neg_Base_Rate
    """
    if len(raw_rows) < 4:
        raise ValueError(
            f"Rate table has unexpected structure: {len(raw_rows)} rows (need ≥ 4)."
        )

    # Build zone map dynamically from ShipbobZone header row (row 0).
    # Each column that contains a numeric zone number maps col_index → zone.
    # Structure: col 0=WeightStart, col 1=WeightEnd, col 2=Currency, col 3+=zone rates
    # (one column per zone, not two — e.g. zone 1 @ col 3, zone 2 @ col 4, etc.)
    zone_col_map = {}   # col_index → zone_number (int)
    for col_idx, val in enumerate(raw_rows[0]):
        val_str = str(val).strip()
        if val_str.isdigit():
            zone_col_map[col_idx] = int(val_str)

    if not zone_col_map:
        raise ValueError(
            "Could not find zone numbers in ShipbobZone header row. "
            f"Row 0 contents: {raw_rows[0]}"
        )

    data_rows = raw_rows[3:]   # skip ShipbobZone, ZoneName, col-header rows

    records = []
    for row in data_rows:
        if not row or not str(row[0]).strip():
            continue
        try:
            w_start = float(str(row[0]).replace(",", "").strip())
            w_end   = float(str(row[1]).replace(",", "").strip()) if len(row) > 1 else w_start
        except (ValueError, IndexError):
            continue

        for col_idx, zone in zone_col_map.items():
            if col_idx >= len(row):
                continue
            raw_val = str(row[col_idx]).replace("$", "").replace(",", "").strip()
            try:
                rate = float(raw_val)
            except ValueError:
                continue
            records.append({
                "WeightOuncesStart": w_start,
                "WeightOuncesEnd":   w_end,
                "Zone":              zone,
                "Pre_Neg_Base_Rate": rate,
            })

    if not records:
        raise ValueError("Could not parse any rate records from the rate table.")

    return (
        pd.DataFrame(records)
        .sort_values(["Zone", "WeightOuncesStart"])
        .reset_index(drop=True)
    )


def load_rate_table(creds, sheet_id: str) -> pd.DataFrame:
    """
    Load and parse the Shipbob pre-negotiated rate table.

    Tries the Google Sheets API first (works for native Google Sheets).
    If the file is an xlsx document (common error: "This operation is not
    supported for this document"), falls back to downloading via the Drive API
    and parsing with pandas/openpyxl.

    Returns tidy DataFrame: WeightOuncesStart | WeightOuncesEnd | Zone | Pre_Neg_Base_Rate
    """
    sheets_service = build("sheets", "v4", credentials=creds)

    # ── Attempt 1: Sheets API (native Google Sheets) ───────────────────────────
    print("  Trying Sheets API...")
    try:
        result = sheets_service.spreadsheets().values().get(
            spreadsheetId=sheet_id,
            range="A1:X"
        ).execute()
        raw = result.get("values", [])
        if raw and len(raw) >= 4:
            rate_df = _parse_rate_raw_rows(raw)
            print(f"  Rate table loaded via Sheets API: "
                  f"{rate_df['WeightOuncesStart'].nunique()} weight brackets × "
                  f"{rate_df['Zone'].nunique()} zones → "
                  f"{len(rate_df):,} combinations")
            return rate_df
    except Exception as e:
        err_str = str(e)
        if "not supported for this document" in err_str or "400" in err_str:
            print("  Sheets API unavailable for this file (xlsx format). "
                  "Falling back to Drive API download...")
        else:
            raise

    # ── Attempt 2: Drive API download (xlsx file) ──────────────────────────────
    print("  Downloading xlsx via Drive API...")
    try:
        drive_service = build("drive", "v3", credentials=creds)
        request = drive_service.files().get_media(fileId=sheet_id)
        buf = io.BytesIO()
        downloader = MediaIoBaseDownload(buf, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()
        buf.seek(0)
    except Exception as e:
        raise RuntimeError(
            f"Could not download rate table via Drive API: {e}\n"
            "  If the file is an xlsx, ensure the Drive scope was granted "
            "and re-run to trigger the OAuth flow."
        ) from e

    print("  Parsing xlsx...")
    try:
        # Read all rows as strings to preserve formatting (no header parsing yet)
        xl = pd.read_excel(buf, sheet_name=0, header=None, dtype=str)
    except Exception as e:
        raise RuntimeError(
            f"Could not parse xlsx: {e}\n"
            "  Ensure openpyxl is installed: pip install openpyxl"
        ) from e

    # Convert to list-of-lists to reuse the same parser
    raw_rows = xl.fillna("").values.tolist()
    rate_df  = _parse_rate_raw_rows(raw_rows)
    print(f"  Rate table loaded via Drive API (xlsx): "
          f"{rate_df['WeightOuncesStart'].nunique()} weight brackets × "
          f"{rate_df['Zone'].nunique()} zones → "
          f"{len(rate_df):,} combinations")
    return rate_df


# ── Shipment Data ──────────────────────────────────────────────────────────────

def load_shipments(service) -> pd.DataFrame:
    """Load and clean Export + Old Shipments tabs, concat and dedup on OrderID."""
    print("Loading shipment data...")

    new_df = read_tab(service, EXPORT_SHEET_ID, "Export")
    old_df = read_tab(service, EXPORT_SHEET_ID, "Old Shipments")

    print(f"  Export tab:        {len(new_df):,} rows")
    print(f"  Old Shipments tab: {len(old_df):,} rows")

    if new_df.empty and old_df.empty:
        raise RuntimeError("Both shipment tabs are empty — nothing to process.")

    combined = pd.concat(
        [df for df in [old_df, new_df] if not df.empty],
        ignore_index=True
    )

    # Deduplicate on OrderID — keep last (newest export wins)
    if "OrderID" in combined.columns and "Transaction Date" in combined.columns:
        combined["Transaction Date"] = pd.to_datetime(
            combined["Transaction Date"], errors="coerce", infer_datetime_format=True
        )
        combined = (
            combined
            .sort_values("Transaction Date")
            .drop_duplicates(subset=["OrderID"], keep="last")
            .reset_index(drop=True)
        )

    # Parse numeric columns (same pattern as data_loader.py)
    numeric_cols = [
        "Original Invoice", "Fulfillment without Surcharge", "Surcharge Applied",
        "WMS Fuel Surcharge", "Delivery Area Surcharge", "Residential Area Surcharge",
        "Address Correction Fee", "Other Order Fee", "Insurance Amount",
        "Actual Weight (Oz)", "Dim Weight(Oz)", "Billable Weight(Oz)",
        "Transit Time (Days)", "Length", "Width", "Height",
    ]
    for col in numeric_cols:
        if col in combined.columns:
            combined[col] = (
                combined[col].astype(str)
                             .str.replace(r"[\$,]", "", regex=True)
                             .pipe(pd.to_numeric, errors="coerce")
                             .fillna(0)
            )

    # Parse Zone Used as integer
    if "Zone Used" in combined.columns:
        combined["Zone Used"] = pd.to_numeric(
            combined["Zone Used"].astype(str).str.strip(), errors="coerce"
        )

    print(f"  Combined (deduped): {len(combined):,} rows")
    return combined


def filter_shipments(df: pd.DataFrame) -> pd.DataFrame:
    """
    Keep rows suitable for rate comparison.
    Excludes rows with no zone or no billable weight (can't do a rate lookup).
    Prints a breakdown of transaction types and statuses found.
    """
    if "Transaction Type" in df.columns:
        type_counts = df["Transaction Type"].value_counts().to_dict()
        print(f"  Transaction types: {type_counts}")
    if "Transaction Status" in df.columns:
        status_counts = df["Transaction Status"].value_counts().to_dict()
        print(f"  Transaction statuses: {status_counts}")

    before = len(df)
    # Only exclude rows where we literally cannot compute a rate lookup
    if "Zone Used" in df.columns:
        df = df[df["Zone Used"].notna() & (df["Zone Used"].astype(str).str.strip() != "")].copy()
    if "Billable Weight(Oz)" in df.columns:
        df = df[df["Billable Weight(Oz)"].notna() & (df["Billable Weight(Oz)"] > 0)].copy()

    removed = before - len(df)
    if removed:
        print(f"  Excluded {removed:,} rows with missing zone or weight")
    print(f"  Rows ready for rate lookup: {len(df):,}")
    return df


# ── Rate Lookup ────────────────────────────────────────────────────────────────

def lookup_rates(shipments: pd.DataFrame, rate_df: pd.DataFrame) -> pd.DataFrame:
    """
    Vectorised rate lookup using pandas merge_asof, processed per zone.

    For each shipment finds the rate table row where:
      WeightOuncesStart <= Billable Weight(Oz) <= WeightOuncesEnd
      AND Zone == Zone Used

    Rows with no match (unknown zone or weight above max bracket) get
      Pre_Neg_Base_Rate = NaN  and  Rate_Lookup_Status = "Weight out of range"
    """
    df = shipments.copy()
    df["_orig_index"] = range(len(df))

    result_frames = []
    matched_orig_indices = set()

    for zone in sorted(rate_df["Zone"].unique()):
        zone_rates     = rate_df[rate_df["Zone"] == zone].sort_values("WeightOuncesStart")
        zone_shipments = df[df["Zone Used"] == zone].sort_values("Billable Weight(Oz)")

        if zone_shipments.empty:
            continue

        merged = pd.merge_asof(
            zone_shipments,
            zone_rates[["WeightOuncesStart", "WeightOuncesEnd", "Pre_Neg_Base_Rate"]],
            left_on="Billable Weight(Oz)",
            right_on="WeightOuncesStart",
            direction="backward"
        )

        # Invalidate lookups where weight exceeds the bracket's upper bound
        out_of_range = (
            merged["Billable Weight(Oz)"] > merged["WeightOuncesEnd"]
        ) | merged["WeightOuncesEnd"].isna()
        merged.loc[out_of_range, "Pre_Neg_Base_Rate"] = float("nan")

        result_frames.append(merged)
        matched_orig_indices.update(merged["_orig_index"].tolist())

    # Rows with no zone match (zone is NaN or not in rate table)
    unmatched = df[~df["_orig_index"].isin(matched_orig_indices)].copy()
    if not unmatched.empty:
        unmatched["Pre_Neg_Base_Rate"]   = float("nan")
        unmatched["WeightOuncesStart"]   = float("nan")
        unmatched["WeightOuncesEnd"]     = float("nan")
        result_frames.append(unmatched)

    if not result_frames:
        df["Pre_Neg_Base_Rate"]  = float("nan")
        df["WeightOuncesStart"]  = float("nan")
        df["WeightOuncesEnd"]    = float("nan")
        return df

    enriched = (
        pd.concat(result_frames, ignore_index=True)
        .sort_values("_orig_index")
        .reset_index(drop=True)
        .drop(columns=["_orig_index", "WeightOuncesStart", "WeightOuncesEnd"],
              errors="ignore")
    )
    return enriched


# ── Enrichment Computation ─────────────────────────────────────────────────────

def compute_enriched_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Add 8 new columns to each shipment row:

    Pre_Neg_Base_Rate           — raw rate table lookup value
    Pre_Neg_Fuel_Surcharge      — base × 20.5%
    Pre_Neg_Residential_Sur     — $1.08 if Residential Area Surcharge > 0
    Pre_Neg_DAS_Surcharge       — $0.10 if Delivery Area Surcharge > 0
    Pre_Neg_Total_Rate          — sum of all 4 above
    Current_Total               — Original Invoice (alias for clarity)
    Savings_Per_Shipment        — Pre_Neg_Total_Rate − Current_Total
    Rate_Lookup_Status          — "OK" or "Weight out of range"
    """
    base = df["Pre_Neg_Base_Rate"]

    df["Pre_Neg_Fuel_Surcharge"] = (base * FUEL_SURCHARGE_PCT).round(2)

    df["Pre_Neg_Residential_Sur"] = df["Residential Area Surcharge"].apply(
        lambda v: RESIDENTIAL_SUR if pd.notna(v) and v > 0 else 0.0
    )

    df["Pre_Neg_DAS_Surcharge"] = df["Delivery Area Surcharge"].apply(
        lambda v: DAS_RESIDENTIAL_SUR if pd.notna(v) and v > 0 else 0.0
    )

    df["Pre_Neg_Total_Rate"] = (
        base.fillna(0)
        + df["Pre_Neg_Fuel_Surcharge"].fillna(0)
        + df["Pre_Neg_Residential_Sur"]
        + df["Pre_Neg_DAS_Surcharge"]
    ).round(2)

    df["Current_Total"]        = df["Original Invoice"].round(2)
    df["Savings_Per_Shipment"] = (df["Pre_Neg_Total_Rate"] - df["Current_Total"]).round(2)

    df["Rate_Lookup_Status"] = base.apply(
        lambda v: "OK" if pd.notna(v) and v > 0 else "Weight out of range"
    )

    df["Pre_Neg_Base_Rate"] = base.round(4)
    return df


# ── Summary Tab ────────────────────────────────────────────────────────────────

def build_summary(df: pd.DataFrame) -> pd.DataFrame:
    """
    Build aggregated summary grouped by:  Carrier | Zone | Merchant Name | Week
    Only rows with Rate_Lookup_Status == "OK" are included.
    """
    ok = df[df["Rate_Lookup_Status"] == "OK"].copy()
    if ok.empty:
        return pd.DataFrame()

    def agg_group(col: str) -> pd.DataFrame:
        grp = (
            ok.groupby(col, dropna=False)
            .agg(
                Shipment_Count     =(col if col != "OrderID" else "OrderID", "count"),
                Total_Pre_Neg_Rate =("Pre_Neg_Total_Rate",  "sum"),
                Total_Current      =("Current_Total",        "sum"),
                Total_Savings      =("Savings_Per_Shipment", "sum"),
            )
            .reset_index()
        )
        grp["Avg_Savings_Per_Shipment"] = (
            grp["Total_Savings"] / grp["Shipment_Count"]
        ).round(2)
        grp["Group_By"]    = col
        grp["Group_Value"] = grp[col].astype(str)
        for money_col in ["Total_Pre_Neg_Rate", "Total_Current", "Total_Savings"]:
            grp[money_col] = grp[money_col].round(2)
        return grp[["Group_By", "Group_Value", "Shipment_Count",
                    "Total_Pre_Neg_Rate", "Total_Current",
                    "Total_Savings", "Avg_Savings_Per_Shipment"]]

    frames = []
    for group_col in ["Carrier", "Zone Used", "Merchant Name", "Week"]:
        if group_col in ok.columns:
            frames.append(agg_group(group_col))

    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()


# ── Main ───────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="Shipping Pre vs Post Negotiation Enrichment",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__
    )
    parser.add_argument(
        "--dry-run", action="store_true",
        help="Compute only — print summary, do not write to Google Sheets"
    )
    args = parser.parse_args()

    print("=" * 65)
    print("  Shipping Negotiation Comparison Enrichment")
    print(f"  Mode: {'DRY RUN (no writes)' if args.dry_run else 'LIVE — writes to Google Sheets'}")
    print("=" * 65)

    config = load_config()
    if "pre_neg_rates_sheet_id" not in config or not config["pre_neg_rates_sheet_id"]:
        print("\nERROR: 'pre_neg_rates_sheet_id' missing from config.json")
        print('  Add:  "pre_neg_rates_sheet_id": "1sLyx152ZBC9OZN8zeMo0SP8xuupnk79k"')
        sys.exit(1)

    # ── Auth ──────────────────────────────────────────────────────────────────
    print("\nAuthenticating with Google (Sheets + Drive scopes)...")
    creds   = get_credentials()
    service = build("sheets", "v4", credentials=creds)

    # ── Rate table ────────────────────────────────────────────────────────────
    print("\nLoading pre-negotiated rate table...")
    rate_df = load_rate_table(creds, config["pre_neg_rates_sheet_id"])

    # ── Shipment data ─────────────────────────────────────────────────────────
    print()
    shipments = load_shipments(service)
    shipments = filter_shipments(shipments)

    if shipments.empty:
        print("\nNo shipment rows found after filtering. Nothing to process.")
        sys.exit(0)

    # ── Rate lookup + enrichment ───────────────────────────────────────────────
    print("\nLooking up pre-negotiation rates...")
    enriched = lookup_rates(shipments, rate_df)
    enriched = compute_enriched_columns(enriched)

    # ── Stats printout ────────────────────────────────────────────────────────
    ok_mask   = enriched["Rate_Lookup_Status"] == "OK"
    ok_count  = ok_mask.sum()
    oor_count = (~ok_mask).sum()
    ok_rows   = enriched[ok_mask]

    total_pre = ok_rows["Pre_Neg_Total_Rate"].sum()
    total_cur = ok_rows["Current_Total"].sum()
    total_sav = ok_rows["Savings_Per_Shipment"].sum()
    avg_sav   = total_sav / ok_count if ok_count else 0

    print(f"\n{'─' * 65}")
    print(f"  Total shipments processed:    {len(enriched):>10,}")
    print(f"  Rate lookup OK:               {ok_count:>10,}")
    print(f"  Weight out of range:          {oor_count:>10,}")
    print(f"  Total pre-negotiation cost:   ${total_pre:>12,.2f}")
    print(f"  Total current cost:           ${total_cur:>12,.2f}")
    print(f"  Total savings:                ${total_sav:>12,.2f}")
    print(f"  Avg savings per shipment:     ${avg_sav:>12.2f}")
    print(f"{'─' * 65}")

    if "Carrier" in enriched.columns and ok_count > 0:
        print("\n  Savings by Carrier:")
        carrier_grp = (
            ok_rows.groupby("Carrier")["Savings_Per_Shipment"]
            .agg(count="count", total="sum", avg="mean")
            .sort_values("total", ascending=False)
        )
        for carrier, row in carrier_grp.iterrows():
            print(f"    {str(carrier):<28} {int(row['count']):>6} shipments  "
                  f"total: ${row['total']:>10,.2f}   avg: ${row['avg']:.2f}")

    if args.dry_run:
        print("\n[DRY RUN] Computation complete. No data written to Google Sheets.")
        return

    # ── Create / get comparison sheet ─────────────────────────────────────────
    print("\nPreparing comparison spreadsheet...")
    comp_sheet_id = create_or_get_comparison_sheet(service, config)

    # ── Write output ──────────────────────────────────────────────────────────
    print("\nWriting to Google Sheets...")
    output_df = select_output_columns(enriched)
    clear_and_write_tab(service, comp_sheet_id, "Enriched_Data", output_df)

    summary_df = build_summary(enriched)
    if not summary_df.empty:
        clear_and_write_tab(service, comp_sheet_id, "Summary", summary_df)

    sheet_url = f"https://docs.google.com/spreadsheets/d/{comp_sheet_id}"
    print(f"\n{'=' * 65}")
    print(f"  Done! Open your comparison sheet:")
    print(f"  {sheet_url}")
    print(f"\n  Enriched_Data tab: {len(enriched):,} shipment rows + 8 new columns")
    if not summary_df.empty:
        print(f"  Summary tab:       {len(summary_df):,} aggregated rows")
    print(f"{'=' * 65}\n")


if __name__ == "__main__":
    main()

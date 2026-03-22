"""
setup_daily_metrics.py  —  One-time script to create the Daily Metrics tab
in the destination Google Sheet with formula-driven columns.

Run once:
    python3 setup_daily_metrics.py

Uses the same OAuth credentials as the rest of the project (google_token.pickle).
Safe to re-run: clears and rewrites the tab each time.
"""

import os
import pickle

SHEET_ID = "1tuo7knxTvOR3snd_u1AnnW1iiN9l1-TaOLi-YaVqLM0"
TAB_NAME  = "Daily Metrics"
SCOPES    = ["https://www.googleapis.com/auth/spreadsheets"]


def get_creds():
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
    raise RuntimeError("google_token.pickle not found. Run the OAuth flow first.")


def get_or_create_tab(service, tab_name):
    meta = service.spreadsheets().get(spreadsheetId=SHEET_ID).execute()
    for s in meta["sheets"]:
        if s["properties"]["title"] == tab_name:
            return s["properties"]["sheetId"]
    resp = service.spreadsheets().batchUpdate(
        spreadsheetId=SHEET_ID,
        body={"requests": [{"addSheet": {"properties": {"title": tab_name}}}]},
    ).execute()
    return resp["replies"][0]["addSheet"]["properties"]["sheetId"]


def main():
    from googleapiclient.discovery import build

    creds   = get_creds()
    service = build("sheets", "v4", credentials=creds)

    print(f"Setting up '{TAB_NAME}' tab…")
    sheet_id = get_or_create_tab(service, TAB_NAME)

    # Clear the tab first
    service.spreadsheets().values().clear(
        spreadsheetId=SHEET_ID, range=f"'{TAB_NAME}'"
    ).execute()

    # ── Layout ─────────────────────────────────────────────────────────────────
    # Cols A–K: data + formulas
    # Cols M–N: config labels + values (N1=18.25, N2=0.28, N3=0.31)
    # The formula row is row 2; config occupies N1:N3 independently

    data_headers = [
        "Date", "Daily Orders", "Labor Hours Outbound", "Labor Hours Total",
        "Labor Cost/hr", "Week (Mon)", "OPLH",
        "Total Labor Cost Per Order", "Outbound Labor Cost Per Order",
        "Packaging Cost Per Order", "Month",
    ]

    data_formulas = [
        # A2 — unique sorted dates from Export
        "=SORT(UNIQUE(FILTER('Export'!$J$2:$J,'Export'!$J$2:$J<>\"\")),1,TRUE)",
        # B2 — daily order count
        "=ARRAYFORMULA(IF(A2:A=\"\",\"\",COUNTIF('Export'!$J:$J,A2:A)))",
        # C2 — outbound labor hours (2025 first, then 2026)
        (
            "=ARRAYFORMULA(IF(A2:A=\"\",\"\","
            "IFERROR(XLOOKUP(A2:A,'Labor Hours 2025'!$A:$A,'Labor Hours 2025'!$E:$E,\"\",0),"
            "IFERROR(XLOOKUP(A2:A,'Labor Hours 2026'!$A:$A,'Labor Hours 2026'!$E:$E,\"\",0),\"\"))))"
        ),
        # D2 — total labor hours
        (
            "=ARRAYFORMULA(IF(A2:A=\"\",\"\","
            "IFERROR(XLOOKUP(A2:A,'Labor Hours 2025'!$A:$A,'Labor Hours 2025'!$B:$B,\"\",0),"
            "IFERROR(XLOOKUP(A2:A,'Labor Hours 2026'!$A:$A,'Labor Hours 2026'!$B:$B,\"\",0),\"\"))))"
        ),
        # E2 — labor cost/hr from config
        "=ARRAYFORMULA(IF(A2:A=\"\",\"\",$N$1))",
        # F2 — week start (Monday)
        "=ARRAYFORMULA(IF(A2:A=\"\",\"\",A2:A-WEEKDAY(A2:A,2)+1))",
        # G2 — OPLH (orders per outbound labor hour)
        "=ARRAYFORMULA(IF((A2:A=\"\")+(B2:B=0)+(C2:C=\"\")+(C2:C=0)>0,\"\",ROUND(B2:B/C2:C,1)))",
        # H2 — total labor cost per order
        "=ARRAYFORMULA(IF((A2:A=\"\")+(B2:B=0)+(D2:D=\"\")+(D2:D=0)>0,\"\",ROUND(D2:D*$N$1/B2:B,2)))",
        # I2 — outbound labor cost per order
        "=ARRAYFORMULA(IF((A2:A=\"\")+(B2:B=0)+(C2:C=\"\")+(C2:C=0)>0,\"\",ROUND(C2:C*$N$1/B2:B,2)))",
        # J2 — packaging cost per order (rate changed Oct 16, 2025)
        "=ARRAYFORMULA(IF(A2:A=\"\",\"\",IF(A2:A>=DATE(2025,10,16),$N$3,$N$2)))",
        # K2 — month start date
        "=ARRAYFORMULA(IF(A2:A=\"\",\"\",DATE(YEAR(A2:A),MONTH(A2:A),1)))",
    ]

    service.spreadsheets().values().batchUpdate(
        spreadsheetId=SHEET_ID,
        body={
            "valueInputOption": "USER_ENTERED",
            "data": [
                # A1:K1 — column headers
                {"range": f"'{TAB_NAME}'!A1:K1", "values": [data_headers]},
                # A2:K2 — formula row (spills downward automatically)
                {"range": f"'{TAB_NAME}'!A2:K2", "values": [data_formulas]},
                # M1:N3 — config area (labels + values)
                {
                    "range": f"'{TAB_NAME}'!M1:N3",
                    "values": [
                        ["Labor Rate ($/hr)",        18.25],
                        ["Pkg Cost < Oct 16 2025",   0.28],
                        ["Pkg Cost >= Oct 16 2025",  0.31],
                    ],
                },
            ],
        },
    ).execute()

    # ── Format: bold header, freeze row 1 ──────────────────────────────────────
    service.spreadsheets().batchUpdate(
        spreadsheetId=SHEET_ID,
        body={
            "requests": [
                {
                    "repeatCell": {
                        "range": {"sheetId": sheet_id, "startRowIndex": 0, "endRowIndex": 1},
                        "cell": {
                            "userEnteredFormat": {
                                "textFormat": {"bold": True},
                                "backgroundColor": {"red": 0.18, "green": 0.20, "blue": 0.26},
                            }
                        },
                        "fields": "userEnteredFormat(textFormat,backgroundColor)",
                    }
                },
                {
                    "updateSheetProperties": {
                        "properties": {
                            "sheetId": sheet_id,
                            "gridProperties": {"frozenRowCount": 1},
                        },
                        "fields": "gridProperties.frozenRowCount",
                    }
                },
            ]
        },
    ).execute()

    print(f"✓ '{TAB_NAME}' tab ready.")
    print("  Columns A–K: formula-driven, auto-populate from Export + Labor Hours tabs.")
    print("  Columns M–N: config (labor rate, packaging cost thresholds).")


if __name__ == "__main__":
    main()

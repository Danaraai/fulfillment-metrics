# Google Sheet Data Append Procedure

**CRITICAL: This is the ONLY way to update the Export sheet. Follow exactly.**

## Sheet Details
- **Sheet ID:** `1tuo7knxTvOR3snd_u1AnnW1iiN9l1-TaOLi-YaVqLM0`
- **Target Tab:** `Export`
- **Header Row:** Row 1 (DO NOT OVERWRITE)
- **Existing Data:** Always present from previous imports

## ⚠️ KEY RULE: ALWAYS APPEND, NEVER OVERWRITE

When updating with new data:
1. **DO NOT** upload starting at `Export!A2` (this overwrites existing data)
2. **ALWAYS** read the current sheet first to find the last row with data
3. **THEN** upload new data starting at the row AFTER the last existing row

## Correct Procedure

### Step 1: Read Existing Data Count
```python
result = service.spreadsheets().values().get(
    spreadsheetId=SHEET_ID,
    range="Export!A:A"
).execute()
existing_rows = len(result.get('values', []))
next_row = existing_rows + 1  # Row number to start appending
```

### Step 2: Prepare New Data
- Load Excel from `~/Downloads/data.xlsx` (sheet name: "Export")
- **Sort by Transaction Date** (ascending: earliest to latest)
- **Calculate Week column** with date ranges: `"MMM DD - MMM DD"` format
  - Example: "Apr 13 - Apr 19" for the week containing April 13-19
  - For missing dates: leave Week column empty
- **Format all dates** as strings: `"YYYY-MM-DD HH:MM:SS"`
- **Do NOT use formulas** - only raw string values

### Step 3: Single Upload Operation
```python
# IMPORTANT: Always upload as RAW, never USER_ENTERED
service.spreadsheets().values().update(
    spreadsheetId=SHEET_ID,
    range=f"Export!A{next_row}",  # Use calculated row number
    valueInputOption="RAW",        # CRITICAL: not USER_ENTERED
    body={'values': rows_to_upload}
).execute()
```

### Step 4: Verify
- Check that dates display correctly (not corrupted)
- Spot-check Week column shows proper date ranges
- Confirm no formulas were uploaded (only raw values)

## What NOT to Do

❌ Start upload at `Export!A2` without checking existing data  
❌ Use Google Sheets API sortRange after uploading  
❌ Add formulas via batchUpdate after uploading  
❌ Use `USER_ENTERED` as valueInputOption (causes date formatting issues)  
❌ Try to sort via API (use Google Sheets UI manually if needed)  
❌ Upload multiple times in sequence (do all processing locally, then one upload)

## Lessons from Data Corruption Incident

**What went wrong:**
- Appended data, then tried to add WEEKNUM formulas via API → corrupted dates
- Used API sortRange after upload → further corrupted data
- Multiple sequential operations instead of one clean upload → dates showed as "1900-02-20"

**The fix that works:**
- Load → Sort locally → Calculate weeks locally → Format as strings locally → Single clean upload with `valueInputOption="RAW"`

## Column Mapping

| Google Sheet Column | Source |
|-------------------|--------|
| Column A | Week (calculated date range) |
| Column B | Transaction Date (from Excel Column A) |
| Columns C-AL | Excel Columns B-AK |

## Date Format Examples

- Excel input: `2026-04-29 14:30:15`
- Week output: `"Apr 27 - Apr 29"`
- Transaction Date output: `"2026-04-29 14:30:15"`

All dates must be uploaded as TEXT strings, not Excel date numbers.

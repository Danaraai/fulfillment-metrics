#!/bin/bash
# One-time setup script for Power BI → Google Sheets exporter

set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

echo "============================================================"
echo "  Power BI → Google Sheets Exporter — Setup"
echo "============================================================"
echo ""

# 1. Install Python packages
echo "[1/3] Installing Python packages..."
pip3 install --user -r requirements.txt
echo "  ✓ Python packages installed"

# 2. Install Playwright browser
echo ""
echo "[2/3] Installing Playwright Chromium browser..."
python3 -m playwright install chromium
echo "  ✓ Chromium installed"

# 3. Create needed directories
mkdir -p session downloads logs

echo ""
echo "[3/3] Setup complete!"
echo ""
echo "============================================================"
echo "  NEXT STEPS"
echo "============================================================"
echo ""
echo "STEP 1 — Configure Google Sheets"
echo "  a. Go to: https://console.cloud.google.com/apis/credentials"
echo "  b. Create a project (or select existing)"
echo "  c. Enable 'Google Sheets API'"
echo "  d. Create credentials:"
echo "     Option A (easiest): OAuth 2.0 Client ID → Desktop App"
echo "       → Download JSON → save as: google_oauth_credentials.json"
echo "     Option B (best for automation): Service Account"
echo "       → Download JSON → save as: google_service_account.json"
echo "       → Share your Google Sheet with the service account email"
echo ""
echo "STEP 2 — Get your Google Sheet ID"
echo "  Your sheet URL looks like:"
echo "  https://docs.google.com/spreadsheets/d/SHEET_ID_HERE/edit"
echo "  Copy the SHEET_ID_HERE and paste it in config.json"
echo ""
echo "STEP 3 — Find your visual name"
echo "  Run this to see all visual titles on the Power BI report:"
echo "    python3 main.py --discover"
echo "  Then update 'visual_title' in config.json"
echo ""
echo "STEP 4 — Test the export"
echo "    python3 main.py"
echo ""
echo "STEP 5 — Schedule weekly (every Monday 8am)"
echo "    python3 main.py --schedule"
echo ""

/**
 * Labor Hours Sync — Google Apps Script
 *
 * Reads daily tabs from the Labor Hours source sheets and writes
 * consolidated rows to the destination sheet (one row per day).
 *
 * Columns written: Date | Total Hours | Emp Hours | Temp Hours | Outbound Hours | Inbound Hours | Headcount
 *
 * Rules:
 *  - Always stops at YESTERDAY (ignores pre-created future tabs)
 *  - Skips dates already in the destination (safe to re-run)
 *  - Run once manually for 2025; set a weekly trigger for 2026
 *
 * HOW TO INSTALL:
 *  1. Open your destination sheet: https://docs.google.com/spreadsheets/d/1tuo7knxTvOR3snd_u1AnnW1iiN9l1-TaOLi-YaVqLM0
 *  2. Extensions → Apps Script
 *  3. Delete the default code and paste this entire file
 *  4. Click Save (floppy icon)
 *  5. Run syncLaborHours2025() once (click ▶ Run)
 *  6. To schedule 2026: Triggers (clock icon) → Add Trigger → syncLaborHours2026 → Weekly → Saturday
 */

// ─── Configuration ────────────────────────────────────────────────────────────

var CONFIG = {
  DEST_SHEET_ID: '1tuo7knxTvOR3snd_u1AnnW1iiN9l1-TaOLi-YaVqLM0',

  SOURCE_2025: '1z3SFQrxnSOVIzPizHupDMmmzl705tGn4bpsn9rEpd9w',
  SOURCE_2026: '1QHSkxmnuUaQtVsiTQL5YvCqOKZqDnMr1WqJODKKAqh0',

  TAB_2025:   'Labor Hours 2025',
  TAB_2026:   'Labor Hours 2026',

  SKIP_TABS:  ['EOM Report', 'In/Out Hours', 'Tracker'],

  HEADER: ['Date', 'Total Hours', 'Emp Hours', 'Temp Hours',
           'Outbound Hours', 'Inbound Hours', 'Headcount'],
};

// ─── Entry points ─────────────────────────────────────────────────────────────

/** Run both years in sequence */
function syncAll() {
  syncLaborHours2025();
  syncLaborHours2026();
}

/** Run once to backfill all of 2025 */
function syncLaborHours2025() {
  _sync(CONFIG.SOURCE_2025, CONFIG.TAB_2025, 2025);
}

/** Run weekly (Saturday) for 2026 — only appends new days */
function syncLaborHours2026() {
  _sync(CONFIG.SOURCE_2026, CONFIG.TAB_2026, 2026);
}

// ─── Core sync logic ──────────────────────────────────────────────────────────

function _sync(sourceId, destTabName, year) {
  // Cutoff = yesterday at end of day (never include today or future tabs)
  var yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);
  yesterday.setHours(23, 59, 59, 0);

  var sourceSS = SpreadsheetApp.openById(sourceId);
  var destSS   = SpreadsheetApp.openById(CONFIG.DEST_SHEET_ID);

  // Get or create destination tab
  var destSheet = destSS.getSheetByName(destTabName);
  if (!destSheet) {
    destSheet = destSS.insertSheet(destTabName);
  }

  // Write header row if sheet is empty
  if (destSheet.getLastRow() === 0) {
    destSheet.appendRow(CONFIG.HEADER);
    SpreadsheetApp.flush();
  }

  // Collect dates already written — compare as "yyyy-MM-dd" strings to avoid
  // timezone issues when comparing Date objects
  var existingDates = {};
  var lastRow = destSheet.getLastRow();
  if (lastRow > 1) {
    var existing = destSheet.getRange(2, 1, lastRow - 1, 1).getValues();
    existing.forEach(function(row) {
      if (row[0]) existingDates[_fmt(row[0])] = true;
    });
  }

  var sourceSheets = sourceSS.getSheets();
  var newRows = [];

  sourceSheets.forEach(function(sheet) {
    var tabName = sheet.getName();
    if (CONFIG.SKIP_TABS.indexOf(tabName) !== -1) return;

    var tabDate = _parseTabDate(tabName, year);
    if (!tabDate) return;
    if (tabDate > yesterday) return;  // skip pre-created future tabs

    var dateStr = _fmt(tabDate);
    if (existingDates[dateStr]) return;  // already synced

    var data = _extractDay(sheet);
    if (!data) return;

    newRows.push([
      tabDate,           // ← Write actual Date object (not string) so XLOOKUP matches Export dates
      data.totalHours,
      data.empHours,
      data.tempHours,
      data.outboundHours,
      data.inboundHours,
      data.headcount,
    ]);
  });

  // Sort chronologically before appending
  newRows.sort(function(a, b) { return a[0] - b[0]; });

  if (newRows.length > 0) {
    var startRow = destSheet.getLastRow() + 1;
    var range = destSheet.getRange(startRow, 1, newRows.length, CONFIG.HEADER.length);
    range.setValues(newRows);
    // Format the Date column as M/D/YYYY to match the Export tab format
    destSheet.getRange(startRow, 1, newRows.length, 1).setNumberFormat('M/D/YYYY');
    SpreadsheetApp.flush();
  }

  Logger.log('Synced ' + newRows.length + ' new days → ' + destTabName +
             ' (skipped ' + Object.keys(existingDates).length + ' already present)');
}

// ─── Data extraction ──────────────────────────────────────────────────────────

/**
 * Reads up to 25 rows of a daily tab and picks out the summary values.
 * Labels float to different rows depending on headcount, so we scan all cells.
 */
function _extractDay(sheet) {
  var numRows = Math.min(sheet.getLastRow(), 25);
  if (numRows < 1) return null;

  var rows = sheet.getRange(1, 1, numRows, 19).getValues();

  var empHours      = 0;
  var tempHours     = 0;
  var totalHours    = 0;
  var outboundHours = 0;
  var inboundHours  = 0;
  var headcount     = 0;

  var STATUSES = ['present', 'late', 'sick', 'call out', 'absent'];

  rows.forEach(function(row) {
    // Count headcount: rows where col C is a known status and col A has a name
    var status = (row[2] || '').toString().trim().toLowerCase();
    var name   = (row[0] || '').toString().trim();
    if (name && STATUSES.indexOf(status) !== -1) headcount++;

    // Scan all cells for summary labels
    for (var i = 0; i < row.length; i++) {
      var cell = (row[i] || '').toString().trim().toLowerCase();

      if (cell.indexOf('total emp hours') !== -1 && cell.indexOf('temp') === -1) {
        empHours = _nextHours(row, i);
      } else if (cell.indexOf('total temp hours') !== -1) {
        tempHours = _nextHours(row, i);
      } else if (cell.indexOf('total emp/temp') !== -1 || cell.indexOf('total emp / temp') !== -1) {
        totalHours = _nextHours(row, i);
      } else if ((cell === 'total hours' || cell.indexOf('total labor hours') !== -1) && totalHours === 0) {
        // 2026 sheets sometimes use this label for the combined total
        var v = _nextHours(row, i);
        if (v > 0) totalHours = v;
      } else if (cell.indexOf('total outbound') !== -1) {
        outboundHours = _nextHours(row, i);
      } else if (cell.indexOf('total inbound') !== -1) {
        inboundHours = _nextHours(row, i);
      }
    }
  });

  // Fallback: if combined total missing, sum emp + temp
  if (totalHours === 0 && (empHours > 0 || tempHours > 0)) {
    totalHours = Math.round((empHours + tempHours) * 100) / 100;
  }

  return {
    empHours:      empHours,
    tempHours:     tempHours,
    totalHours:    totalHours,
    outboundHours: outboundHours,
    inboundHours:  inboundHours,
    headcount:     headcount,
  };
}

// ─── Helpers ──────────────────────────────────────────────────────────────────

/** Parse "May 1", "June 23", "Jan 2, 2026" → Date object */
function _parseTabDate(tabName, defaultYear) {
  var MONTHS = {
    'jan':1,'feb':2,'mar':3,'apr':4,'may':5,'jun':6,
    'jul':7,'aug':8,'sep':9,'oct':10,'nov':11,'dec':12,
    'january':1,'february':2,'march':3,'april':4,'june':6,
    'july':7,'august':8,'september':9,'october':10,'november':11,'december':12
  };

  // Match "May 1" or "May 1, 2026"
  var m = tabName.match(/^(\w+)\s+(\d+)(?:,\s*(\d{4}))?$/);
  if (!m) return null;

  var mon = MONTHS[m[1].toLowerCase()];
  if (!mon) return null;

  var day  = parseInt(m[2], 10);
  var year = m[3] ? parseInt(m[3], 10) : defaultYear;

  var d = new Date(year, mon - 1, day);
  // Sanity check: day must match (guards against invalid dates like Feb 31)
  if (d.getDate() !== day) return null;
  return d;
}

/** Return decimal hours from the first H:MM[:SS] value after labelIdx in a row */
function _nextHours(row, labelIdx) {
  for (var i = labelIdx + 1; i < row.length; i++) {
    var val = (row[i] || '').toString().trim();
    if (/^\d+:\d+/.test(val)) return _parseHours(val);
  }
  return 0;
}

/** "23:10:00" or "6:45" → decimal hours (rounded to 2dp) */
function _parseHours(s) {
  if (!s) return 0;
  var parts = s.trim().split(':');
  var h = parseInt(parts[0], 10) || 0;
  var m = parts.length > 1 ? parseInt(parts[1], 10) || 0 : 0;
  return Math.round((h + m / 60) * 100) / 100;
}

/** Format a Date as "yyyy-MM-dd" in the spreadsheet's timezone */
function _fmt(d) {
  return Utilities.formatDate(
    new Date(d),
    Session.getScriptTimeZone(),
    'yyyy-MM-dd'
  );
}

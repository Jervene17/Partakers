/**
 * Optional companion to the bot's /refresh_dashboard and /sync_dashboard
 * commands. Paste this into the spreadsheet's Extensions > Apps Script
 * editor, save, and it will push any edit made in a "{Service} Dashboard"
 * tab's grid straight into the matching service tab immediately — no need
 * to run /sync_dashboard by hand.
 *
 * This does NOT replace /sync_dashboard; it's a convenience for admins who
 * are actively editing in the Sheet and want changes to take effect right
 * away. /sync_dashboard remains useful for a bulk pass (e.g. after editing
 * many rows offline) and works identically whether or not this script is
 * installed.
 *
 * Setup:
 *   1. Open the Google Sheet -> Extensions -> Apps Script.
 *   2. Delete any placeholder code, paste this file's contents in.
 *   3. Save (the default trigger name "onEdit" is installed automatically
 *      — no manual trigger setup needed for simple, single-cell edits).
 *   4. Edit a Partaker cell in any "{Service} Dashboard" tab and confirm
 *      it updates the corresponding service tab within a second or two.
 *
 * Each "{Service} Dashboard" tab is laid out in stacked monthly blocks,
 * built by the bot's /refresh_dashboard:
 *   "September 2026"
 *   Date          | Role A | Role B | ...
 *   2026-09-06    |        |        |
 *   ...
 *   (blank separator row)
 *   "October 2026"
 *   ...
 * This script finds the nearest column-header row above the edited cell
 * (scanning upward, skipping blank rows) to know which role that column
 * is, and reads the edited row's own Date column for the date.
 * Every service's real data tab keeps the flat shape:
 *   A: Date       B: Role   C: Partaker   D: Status
 */

function onEdit(e) {
  var sheet = e.range.getSheet();
  var sheetName = sheet.getName();
  if (sheetName.slice(-10) !== " Dashboard") return; // only care about "{Service} Dashboard" tabs
  var service = sheetName.slice(0, -10);

  var row = e.range.getRow();
  var col = e.range.getColumn();
  if (col === 1) return; // column A is the Date column, not a role column

  var partaker = e.range.getValue().toString().trim();
  if (!partaker) return; // blank = no change intended, never erase target data

  var date = sheet.getRange(row, 1).getValue().toString().trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) return; // edited row isn't a data row (e.g. a month label)

  var role = findColumnHeader_(sheet, row, col);
  if (!role) return; // no header row found above — layout looks off, skip silently

  var ss = e.source;
  var targetSheet = ss.getSheetByName(service);
  if (!targetSheet) {
    SpreadsheetApp.getActive().toast(
      "Dashboard sync: no sheet named '" + service + "' — edit not applied.",
      "Sync error"
    );
    return;
  }

  var data = targetSheet.getDataRange().getValues(); // includes header row
  var matchRow = -1;
  for (var i = 1; i < data.length; i++) {
    var rDate = formatDate_(data[i][0]);
    var rRole = data[i][1].toString().trim();
    if (rDate === date && rRole === role) {
      matchRow = i + 1; // 1-indexed sheet row
      break;
    }
  }

  if (matchRow > 0) {
    targetSheet.getRange(matchRow, 3).setValue(partaker); // column C = Partaker
  } else {
    targetSheet.appendRow([date, role, partaker, "scheduled"]);
  }
}

/** Scans upward from `row` for the nearest "Date | Role1 | Role2 | ..."
 * header row and returns the role name in column `col`, or null if none
 * is found before a second blank row / the top of the sheet. */
function findColumnHeader_(sheet, row, col) {
  for (var r = row - 1; r >= 1; r--) {
    var firstCell = sheet.getRange(r, 1).getValue().toString().trim();
    if (firstCell === "Date") {
      return sheet.getRange(r, col).getValue().toString().trim();
    }
    // a month-label row (e.g. "September 2026") means we've gone past this
    // block's header without finding it — bail out
    if (firstCell && isMonthLabel_(firstCell)) return null;
  }
  return null;
}

function isMonthLabel_(text) {
  return /^[A-Z][a-z]+ \d{4}$/.test(text);
}

function formatDate_(value) {
  if (Object.prototype.toString.call(value) === "[object Date]") {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  return value.toString().trim();
}

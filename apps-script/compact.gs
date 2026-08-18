/**
 * compact.gs — removes "ghost" blank rows from the data tabs.
 *
 * THE BUG IT FIXES
 *   At some point the Verified column of a Matches tab had checkboxes inserted
 *   down the whole sheet, which writes a literal FALSE into every one of those
 *   cells. A FALSE is *content*, so getLastRow() reported 1000 even though the
 *   tab held no matches. handleSubmission() positions each new match at
 *   getLastRow() + 1, so real matches started landing at row 1001 — under ~999
 *   visually blank rows. Vets Tier 1 and Vets Tier 2 both hit this (found
 *   2026-08-18); Tshwane and 3-Man are padded the same way and would hit it on
 *   their first submission.
 *
 * WHY THIS RUNS SERVER-SIDE
 *   The Sheets bridge can only read/write/clear values. Clearing leaves the
 *   1000 empty rows (and their checkbox validation) in place, and rewriting the
 *   tab through the bridge turns the Submitted/Date datetimes into TEXT —
 *   which would break the Submitted-timestamp key that joins Matches to
 *   Frames/Players. Deleting the rows in Apps Script keeps the cell types.
 *
 * WHAT IT TOUCHES
 *   Matches / Frames / Players only. It deletes a row ONLY when every cell in
 *   it is blank once a boolean checkbox value is ignored — so a row carrying
 *   any date, team, name or score is never at risk. It then strips the leftover
 *   data validation below the last real row. Scores, dates and Verified flags
 *   on surviving rows are untouched. Idempotent: a second run reports 0.
 *
 * USE:  Master sheet → menu GNPA League → "Compact blank rows (all leagues)".
 */

var COMPACT_TABS = ['Matches', 'Frames', 'Players'];

/** True when the row holds nothing but blanks and/or checkbox booleans. */
function compactRowIsGhost_(row) {
  for (var i = 0; i < row.length; i++) {
    var v = row[i];
    if (v === '' || v === null || v === undefined) continue;
    if (typeof v === 'boolean') continue;          // an untouched checkbox
    if (typeof v === 'string' && v.trim() === '') continue;
    return false;                                  // real data — keep the row
  }
  return true;
}

/** Collapse ascending row numbers into [start, count] blocks. */
function compactBlocks_(rows) {
  var blocks = [], i = 0;
  while (i < rows.length) {
    var start = rows[i], n = 1;
    while (i + n < rows.length && rows[i + n] === start + n) n++;
    blocks.push([start, n]);
    i += n;
  }
  return blocks;
}

/** Compact one tab. Returns the number of rows deleted. */
function compactTab_(ss, tabName) {
  var sheet = ss.getSheetByName(tabName);
  if (!sheet || sheet.getLastRow() < 2) return 0;

  var lastRow = sheet.getLastRow();
  var lastCol = Math.max(sheet.getLastColumn(), 1);
  var values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  var ghosts = [];
  for (var i = 0; i < values.length; i++) {
    if (compactRowIsGhost_(values[i])) ghosts.push(i + 2); // sheet row number
  }
  if (!ghosts.length) return 0;

  // Delete bottom-up so earlier row numbers stay valid.
  var blocks = compactBlocks_(ghosts);
  for (var b = blocks.length - 1; b >= 0; b--) {
    sheet.deleteRows(blocks[b][0], blocks[b][1]);
  }

  // Strip leftover checkbox/data validation below the surviving data, and give
  // the tab a clean tail so getLastRow() can never over-report again.
  var keep = sheet.getLastRow();
  var maxRows = sheet.getMaxRows();
  if (maxRows > keep) {
    sheet.getRange(keep + 1, 1, maxRows - keep, sheet.getMaxColumns())
      .clearDataValidations().clearContent();
  }
  return ghosts.length;
}

/** Compact every league in the Master Config tab. */
function compactAllLeagues() {
  var leagues = configLeaguesFromMaster_();
  var lines = [], touched = [];

  for (var i = 0; i < leagues.length; i++) {
    var ss;
    try { ss = SpreadsheetApp.openById(leagues[i].sheetId); }
    catch (e) { lines.push(leagues[i].league + ': cannot open — ' + e.message); continue; }

    var parts = [], total = 0;
    for (var t = 0; t < COMPACT_TABS.length; t++) {
      var n = compactTab_(ss, COMPACT_TABS[t]);
      total += n;
      if (n) parts.push(COMPACT_TABS[t] + ' -' + n);
    }
    lines.push(leagues[i].league + ': ' + (total ? parts.join(', ') : 'clean'));
    if (total) touched.push(leagues[i].league);
  }

  // Standings key off the Submitted timestamp, not row position, so numbers do
  // not change — but rebuild the touched leagues anyway to refresh the cache.
  for (var r = 0; r < touched.length; r++) {
    try { rebuildStandingsForLeague(touched[r]); }
    catch (e) { lines.push(touched[r] + ': rebuild failed — ' + e.message); }
  }

  reportSetup_('Compact blank rows\n\n' + lines.join('\n'));
}

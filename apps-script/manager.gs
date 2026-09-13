/**
 * GNPA — Manager tools tab: add & rename/merge players from the phone
 * (added 2026-07-20)
 *
 * WHY: league managers already open the sheets, but the two roster jobs they
 * need are either fiddly or dangerous to do by hand:
 *   - ADD a new/reserve player  → they don't know which tab/columns.
 *   - RENAME / fix a misspelt player → changing only the roster leaves every
 *     game under the OLD name, so standings split them into two people. This
 *     actually happened (Suanét → Sugnét Esterhuizen, Ladies, 2026-07-20).
 *
 * WHAT: a "Manager" tab (2nd tab, next to Refresh) with two blocks — Add and
 * Rename/Merge — driven by dropdowns + a tick box, so it works inside the
 * Sheets MOBILE app (custom menus don't render there; an installable onEdit
 * trigger fired by a tap is the only thing that does). RENAME rewrites the
 * name across Roster + Players + Frames and rebuilds standings in one action,
 * so histories never split. Renaming a player to a name that already exists
 * MERGES them (one player, all games kept) — that's Jade's "merge = rename".
 *
 * MECHANISM: this reuses the SAME installable onEdit trigger the Refresh box
 * uses (onLeagueRefreshEdit in refresh.gs dispatches Manager-tab edits here).
 * No extra triggers — stays well under the Apps Script trigger cap.
 *
 * SETUP (one-time, from the Master sheet's script editor):
 *   1. Add this file + the 2-line dispatch edit in refresh.gs. Save.
 *   2. Run setupManagerTabs()  (or Master menu → GNPA League →
 *      "Set up manager tools tabs"). Authorize if prompted.
 *   No redeploy needed — installable triggers always run the latest code.
 *   Idempotent: re-run any time (also after adding a NEW league to Config).
 *
 * SCHEMA NOTES (verified across all 9 leagues, 2026-07-20):
 *   - Players/Frames name+team columns are IDENTICAL in every league, but this
 *     resolves them by HEADER NAME (colByHeader_) so a future layout change
 *     can't silently rewrite the wrong column.
 *   - Only Ladies ships a Roster tab; the others are created on first Add
 *     (getRosterForLeague in roster.gs unions the Roster tab when present).
 *   - Team list comes from the Team Codes tab (present in every league).
 *
 * TEAM DROPDOWN IS A LIVE RANGE (changed 2026-08-19 — this was a real bug)
 *   The Team dropdowns used to be built from a SNAPSHOT of the Team Codes tab
 *   (requireValueInList with the names copied in at setup time). If that read
 *   came back empty for any reason, Apps Script happily built a dropdown with
 *   ZERO items and setup still reported success — the manager then sees an
 *   empty picker and "Invalid: Input must be an item on the specified list"
 *   when they type. That is exactly what shipped: every league sheet ended up
 *   with an empty list rule on B5/B12 (found on Vets Tier 1, 2026-08-19).
 *   They now point at Team Codes!A2:A as a RANGE, so the picker always shows
 *   whatever that tab currently holds — new teams appear without re-running
 *   setup, and a league whose Team Codes tab is still empty is REPORTED as
 *   such by setupManagerTabs() instead of failing silently.
 */

var MANAGER_TAB = 'Manager';

// Add-a-player block
var MGR_ADD_TEAM   = 'B5';
var MGR_ADD_NAME   = 'B6';
var MGR_ADD_CHECK  = 'B7';
var MGR_ADD_STATUS = 'B8';

// Rename / fix / merge block
var MGR_REN_TEAM   = 'B12';
var MGR_REN_OLD    = 'B13';
var MGR_REN_NEW    = 'B14';
var MGR_REN_CHECK  = 'B15';
var MGR_REN_STATUS = 'B16';

// ============================================================
//  One-time setup: build the Manager tab on every league sheet
// ============================================================
function setupManagerTabs() {
  var leagues = configLeaguesFromMaster_();
  if (!leagues.length) {
    reportSetup_('No leagues configured in the Master Config tab.');
    return;
  }
  var report = [];
  leagues.forEach(function (entry) {
    try {
      var ss = SpreadsheetApp.openById(entry.sheetId);
      var nTeams = ensureManagerTab_(ss);
      ensureLeagueEditTrigger_(entry.sheetId); // shares the Refresh handler
      // Say how many teams the picker will show. A zero here is the failure
      // mode that shipped an empty dropdown to every league — never hide it.
      report.push((nTeams ? '✓ ' : '⚠ ') + entry.league + ' — ' +
        (nTeams ? nTeams + ' team(s) in the picker'
                : 'Team Codes tab is EMPTY — dropdown will show nothing until teams are added'));
    } catch (e) {
      report.push('✗ ' + entry.league + ' — ' + e.message);
    }
  });
  reportSetup_('Manager tools tabs:\n\n' + report.join('\n'));
}

// Reuse the single per-league onEdit trigger (handler lives in refresh.gs).
// Creating the trigger here means Manager tools work even on a league that
// never had the Refresh checkbox set up.
function ensureLeagueEditTrigger_(sheetId) {
  var has = ScriptApp.getProjectTriggers().some(function (t) {
    return t.getHandlerFunction() === REFRESH_HANDLER &&
           t.getTriggerSourceId() === sheetId;
  });
  if (!has) {
    ScriptApp.newTrigger(REFRESH_HANDLER).forSpreadsheet(sheetId).onEdit().create();
  }
}

function ensureManagerTab_(ss) {
  var sheet = ss.getSheetByName(MANAGER_TAB);
  if (!sheet) {
    sheet = ss.insertSheet(MANAGER_TAB, 1); // second tab (right after Refresh)
  } else {
    ss.setActiveSheet(sheet);
    ss.moveActiveSheet(2);
    sheet.clear();
    sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns())
         .clearDataValidations().removeCheckboxes();
  }

  sheet.setColumnWidth(1, 160);
  sheet.setColumnWidth(2, 320);
  sheet.setHiddenGridlines(true);

  // The picker points at the Team Codes tab itself, not a copy of its values:
  // teams added/renamed there show up immediately, and there is no way to end
  // up with a silently empty list. Invalid entries are REJECTED — this is a
  // pick-from-list field, and a half-typed team name ("Old Boys Unt") used to
  // sail through and then match nothing.
  var teamRange = managerTeamRange_(ss);
  var teamRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(teamRange, true).setAllowInvalid(false)
    .setHelpText('Pick a team from the list').build();

  // ---- Title ----
  band_(sheet, 'A1:B1', '⚙️  MANAGER TOOLS', '#1f6fc4', '#ffffff', 12);
  note_(sheet, 'A2:B2', 'Add or fix players from your phone. Changes go live and rebuild standings automatically.');

  // ---- Add a player ----
  band_(sheet, 'A4:B4', '➕  ADD A PLAYER', '#e8f0fe', '#1a1a1a', 11);
  label_(sheet, 'A5', 'Team');
  sheet.getRange(MGR_ADD_TEAM).setDataValidation(teamRule);
  label_(sheet, 'A6', 'New player name');
  input_(sheet, MGR_ADD_NAME);
  label_(sheet, 'A7', 'Tick to add  →');
  sheet.getRange(MGR_ADD_CHECK).insertCheckboxes().setValue(false);
  label_(sheet, 'A8', 'Result');
  status_(sheet, MGR_ADD_STATUS, 'Waiting');

  // ---- Rename / fix / merge ----
  band_(sheet, 'A10:B10', '✏️  RENAME / FIX / MERGE A PLAYER', '#e8f0fe', '#1a1a1a', 11);
  note_(sheet, 'A11:B11', 'Fixes the name on the roster AND every game, then rebuilds. Renaming to a name that already exists merges the two into one.');
  label_(sheet, 'A12', 'Team');
  sheet.getRange(MGR_REN_TEAM).setDataValidation(teamRule);
  label_(sheet, 'A13', 'Current name');
  input_(sheet, MGR_REN_OLD); // dropdown filled once a team is chosen
  label_(sheet, 'A14', 'Correct name');
  input_(sheet, MGR_REN_NEW);
  label_(sheet, 'A15', 'Tick to rename  →');
  sheet.getRange(MGR_REN_CHECK).insertCheckboxes().setValue(false);
  label_(sheet, 'A16', 'Result');
  status_(sheet, MGR_REN_STATUS, 'Waiting');

  note_(sheet, 'A18:B18', 'Tip: pick the Team first — the "Current name" list fills with that team’s players (including any misspelt ones sitting in the results).');

  return managerTeamList_(ss).length; // reported by setupManagerTabs()
}

// Small layout helpers (keep the tab looking intentional, not like a raw grid)
function band_(sheet, a1, text, bg, fg, size) {
  sheet.getRange(a1).merge();
  sheet.getRange(a1.split(':')[0]).setValue(text)
    .setFontWeight('bold').setFontSize(size || 11)
    .setBackground(bg).setFontColor(fg)
    .setVerticalAlignment('middle');
  sheet.setRowHeight(Number(a1.match(/\d+/)[0]), 28);
}
function note_(sheet, a1, text) {
  sheet.getRange(a1).merge();
  sheet.getRange(a1.split(':')[0]).setValue(text)
    .setFontStyle('italic').setFontColor('#666666').setWrap(true)
    .setVerticalAlignment('middle');
}
function label_(sheet, a1, text) {
  sheet.getRange(a1).setValue(text).setFontWeight('bold').setFontColor('#333333');
}
function input_(sheet, a1) {
  sheet.getRange(a1).setBackground('#fffef2').setBorder(true, true, true, true, false, false, '#cccccc', SpreadsheetApp.BorderStyle.SOLID);
}
function status_(sheet, a1, text) {
  sheet.getRange(a1).setValue(text).setFontStyle('italic').setFontColor('#888888');
}

// The RANGE the team dropdowns read (Team Codes, the Team column, row 2 down).
// A range keeps the picker live; blank cells in it are ignored by Sheets.
function managerTeamRange_(ss) {
  var sh = ss.getSheetByName(TEAM_CODES_TAB);
  if (!sh) {
    throw new Error('no "' + TEAM_CODES_TAB + '" tab — cannot build the team dropdown');
  }
  var col = colByHeader_(sh, 'Team') || 1;
  var rows = Math.max(sh.getMaxRows() - 1, 1);
  return sh.getRange(2, col, rows, 1);
}

// Teams for the dropdowns — from the Team Codes tab (every league has it).
function managerTeamList_(ss) {
  var sh = ss.getSheetByName(TEAM_CODES_TAB);
  if (!sh || sh.getLastRow() < 2) return [];
  var teamCol = colByHeader_(sh, 'Team') || 1;
  var vals = sh.getRange(2, teamCol, sh.getLastRow() - 1, 1).getValues();
  var seen = {}, out = [];
  vals.forEach(function (r) {
    var t = String(r[0] || '').trim();
    if (t && !seen[t.toLowerCase()]) { seen[t.toLowerCase()] = 1; out.push(t); }
  });
  out.sort(function (a, b) { return a.toLowerCase() < b.toLowerCase() ? -1 : 1; });
  return out;
}

// Is this actually a team on this sheet? The dropdown now rejects free text,
// but a value typed before this fix (or pasted in) can still be sitting in the
// cell — and acting on it renames nothing while reporting success.
function managerIsKnownTeam_(ss, team) {
  var t = String(team || '').trim().toLowerCase();
  if (!t) return false;
  return managerTeamList_(ss).some(function (n) { return n.toLowerCase() === t; });
}

// ============================================================
//  Edit handler — dispatched from onLeagueRefreshEdit (refresh.gs)
//  when the edited sheet is the Manager tab.
// ============================================================
function onManagerEdit_(e, sheet) {
  try {
    var a1 = e.range.getA1Notation();
    if (a1 === MGR_REN_TEAM) { refreshRenameNameList_(sheet, e.source); return; }
    if (a1 === MGR_ADD_CHECK && e.range.getValue() === true) { managerDoAdd_(e, sheet); return; }
    if (a1 === MGR_REN_CHECK && e.range.getValue() === true) { managerDoRename_(e, sheet); return; }
  } catch (err) {
    // Never throw out of a trigger — surface failures in the status cells.
    try { sheet.getRange(MGR_REN_STATUS).setValue('✗ ' + err.message); } catch (ignore) {}
  }
}

// When a team is chosen for renaming, fill "Current name" with that team's
// players — union of the Roster tab and everyone who has actually played
// (so a misspelling that only exists in the results still shows up).
// `quiet` suppresses the status line — set when called right after a rename,
// where the status cell already holds the result the manager needs to read.
function refreshRenameNameList_(sheet, ss, quiet) {
  var team = String(sheet.getRange(MGR_REN_TEAM).getValue() || '').trim();
  var oldCell = sheet.getRange(MGR_REN_OLD);
  oldCell.clearContent();
  if (!team) { oldCell.clearDataValidations(); return; }
  var names = playerNamesForTeam_(ss, team);
  if (names.length) {
    oldCell.setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(names, true).setAllowInvalid(false)
      .setHelpText('Pick the name to fix').build());
    if (!quiet) sheet.getRange(MGR_REN_STATUS)
      .setValue(names.length + ' player(s) on ' + team + ' — pick one');
  } else {
    oldCell.clearDataValidations();
    if (!quiet) sheet.getRange(MGR_REN_STATUS)
      .setValue('✗ No players found for ' + team + ' yet');
  }
}

function playerNamesForTeam_(ss, team) {
  var seen = {}, out = [];
  function add(n) {
    n = String(n || '').trim();
    if (n && !seen[n.toLowerCase()]) { seen[n.toLowerCase()] = 1; out.push(n); }
  }
  var ros = ss.getSheetByName(ROSTER_TAB);
  if (ros && ros.getLastRow() > 1) {
    var tC = colByHeader_(ros, 'Team'), nC = colByHeader_(ros, 'Player');
    if (tC && nC) {
      var rv = ros.getRange(2, 1, ros.getLastRow() - 1, ros.getLastColumn()).getValues();
      rv.forEach(function (r) { if (String(r[tC - 1]).trim() === team) add(r[nC - 1]); });
    }
  }
  var pl = ss.getSheetByName(PLAYERS_TAB);
  if (pl && pl.getLastRow() > 1) {
    var ptC = colByHeader_(pl, 'Team'), pnC = colByHeader_(pl, 'Player');
    if (ptC && pnC) {
      var pv = pl.getRange(2, 1, pl.getLastRow() - 1, pl.getLastColumn()).getValues();
      pv.forEach(function (r) { if (String(r[ptC - 1]).trim() === team) add(r[pnC - 1]); });
    }
  }
  out.sort(function (a, b) { return a.toLowerCase() < b.toLowerCase() ? -1 : 1; });
  return out;
}

// ---- Add ----
function managerDoAdd_(e, sheet) {
  var ss = e.source;
  var status = sheet.getRange(MGR_ADD_STATUS);
  var team = String(sheet.getRange(MGR_ADD_TEAM).getValue() || '').trim();
  var name = String(sheet.getRange(MGR_ADD_NAME).getValue() || '').trim();
  e.range.setValue(false); // untick first (a script edit can't re-fire this)

  if (!team) { status.setValue('✗ Pick a team first'); return; }
  if (!managerIsKnownTeam_(ss, team)) {
    status.setValue('✗ “' + team + '” is not a team on this sheet — pick one from the dropdown');
    return;
  }
  if (!name) { status.setValue('✗ Type the player’s name first'); return; }

  var already = playerNamesForTeam_(ss, team).map(function (n) { return n.toLowerCase(); });
  if (already.indexOf(name.toLowerCase()) >= 0) {
    status.setValue('✓ ' + name + ' is already on ' + team);
    sheet.getRange(MGR_ADD_NAME).clearContent();
    return;
  }

  var ros = ensureRosterTab_(ss);
  ros.appendRow([team, name, false, '']);
  // "they'll show in the app now" is only true if the cached roster is dropped.
  try { invalidateLeagueCaches_(leagueNameForSheetId_(ss.getId())); } catch (ce) {}
  status.setValue('✓ Added ' + name + ' to ' + team + ' — they’ll show in the app now');
  sheet.getRange(MGR_ADD_NAME).clearContent();
  // No rebuild needed: a player with no games doesn't change any standings.
}

function ensureRosterTab_(ss) {
  var sh = ss.getSheetByName(ROSTER_TAB);
  if (!sh) {
    sh = ss.insertSheet(ROSTER_TAB);
    sh.getRange(1, 1, 1, ROSTER_HEADERS.length).setValues([ROSTER_HEADERS]).setFontWeight('bold');
  }
  return sh;
}

// ---- Rename / merge (the one that prevents split histories) ----
function managerDoRename_(e, sheet) {
  var ss = e.source;
  var status = sheet.getRange(MGR_REN_STATUS);
  var team = String(sheet.getRange(MGR_REN_TEAM).getValue() || '').trim();
  var oldName = String(sheet.getRange(MGR_REN_OLD).getValue() || '').trim();
  var newName = String(sheet.getRange(MGR_REN_NEW).getValue() || '').trim();
  e.range.setValue(false);

  if (!team) { status.setValue('✗ Pick a team first'); return; }
  if (!managerIsKnownTeam_(ss, team)) {
    status.setValue('✗ “' + team + '” is not a team on this sheet — pick one from the dropdown');
    return;
  }
  if (!oldName) { status.setValue('✗ Pick the current name'); return; }
  if (!newName) { status.setValue('✗ Type the correct name'); return; }
  if (oldName === newName) { status.setValue('✗ Old and new name are the same'); return; }

  var league = leagueNameForSheetId_(ss.getId());
  if (!league) { status.setValue('✗ This sheet is not in the Master Config'); return; }

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    status.setValue('⏳ A rebuild is already running — try again in a minute');
    return;
  }
  try {
    status.setValue('⏳ Renaming …'); SpreadsheetApp.flush();
    var n = 0;
    n += renameInTab_(ss, ROSTER_TAB,  [['Team', 'Player']], team, oldName, newName);
    n += renameInTab_(ss, PLAYERS_TAB, [['Team', 'Player']], team, oldName, newName);
    n += renameInTab_(ss, FRAMES_TAB,  [['Home Team', 'Home Player'], ['Away Team', 'Away Player']], team, oldName, newName);

    // Nothing matched — say so instead of reporting a ✓ for a no-op, and skip
    // the ~20s rebuild. This is what "updated 0 row(s), standings rebuilt"
    // looked like on Vets Tier 1 when the team cell held a half-typed name.
    if (!n) {
      status.setValue('✗ No “' + oldName + '” found in ' + team + ' — nothing changed');
      return;
    }
    dedupeRosterTeamName_(ss, team, newName);

    rebuildTeamStandings(ss, league);
    rebuildPlayerStandings(ss, league);
    try { invalidateLeagueCaches_(league); } catch (ce) {}

    var tz = ss.getSpreadsheetTimeZone();
    status.setValue('✓ ' + oldName + ' → ' + newName + ' — updated ' + n +
      ' row(s), standings rebuilt · ' + Utilities.formatDate(new Date(), tz, 'EEE d MMM HH:mm'));
    sheet.getRange(MGR_REN_OLD).clearContent();
    sheet.getRange(MGR_REN_NEW).clearContent();
    refreshRenameNameList_(sheet, ss, true); // keep the ✓ message on screen
  } catch (err) {
    status.setValue('✗ Failed: ' + err.message);
  } finally {
    try { lock.releaseLock(); } catch (le) {}
  }
}

// Rewrite a player's name in one tab, scoped to their team, resolving columns
// by header so it can't touch the wrong column. Writes back only the name
// column (minimal blast radius — dates/scores/other cells untouched).
function renameInTab_(ss, tabName, pairs, team, oldName, newName) {
  var sh = ss.getSheetByName(tabName);
  if (!sh || sh.getLastRow() < 2) return 0;
  var nRows = sh.getLastRow() - 1;
  var changed = 0;
  pairs.forEach(function (pair) {
    var teamCol = colByHeader_(sh, pair[0]);
    var nameCol = colByHeader_(sh, pair[1]);
    if (!teamCol || !nameCol) return;
    var teamVals = sh.getRange(2, teamCol, nRows, 1).getValues();
    var nameVals = sh.getRange(2, nameCol, nRows, 1).getValues();
    var dirty = false;
    for (var i = 0; i < nRows; i++) {
      if (String(teamVals[i][0]).trim() === team && String(nameVals[i][0]).trim() === oldName) {
        nameVals[i][0] = newName; changed++; dirty = true;
      }
    }
    if (dirty) sh.getRange(2, nameCol, nRows, 1).setValues(nameVals);
  });
  return changed;
}

// After a merge (rename onto an existing name) the roster may hold the same
// player twice for that team — keep one, drop the rest.
function dedupeRosterTeamName_(ss, team, newName) {
  var sh = ss.getSheetByName(ROSTER_TAB);
  if (!sh || sh.getLastRow() < 2) return;
  var teamCol = colByHeader_(sh, 'Team'), nameCol = colByHeader_(sh, 'Player');
  if (!teamCol || !nameCol) return;
  var kept = false;
  for (var r = sh.getLastRow(); r >= 2; r--) {
    var t = String(sh.getRange(r, teamCol).getValue()).trim();
    var nm = String(sh.getRange(r, nameCol).getValue()).trim();
    if (t === team && nm === newName) {
      if (kept) sh.deleteRow(r); else kept = true;
    }
  }
}

// Find a 1-based column index by header text (case-insensitive). 0 if absent.
function colByHeader_(sheet, headerName) {
  var lastCol = sheet.getLastColumn();
  if (lastCol < 1) return 0;
  var hdr = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var want = String(headerName).trim().toLowerCase();
  for (var i = 0; i < hdr.length; i++) {
    if (String(hdr[i]).trim().toLowerCase() === want) return i + 1;
  }
  return 0;
}

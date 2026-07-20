/**
 * GNPA — Mobile "Refresh standings" checkbox (added 2026-07-18)
 *
 * WHY: the Google Sheets mobile app does not render Apps Script custom menus,
 * so managers on their phones have no way to run "Rebuild standings". The one
 * thing that CAN fire a script from a tap inside the mobile app is an
 * installable onEdit trigger — so each league sheet gets a tiny "Refresh" tab
 * (first tab) with a checkbox. Tick it → that league's standings rebuild →
 * the box unticks itself and a status cell shows the result.
 *
 * SETUP (one-time, from the Master sheet's script editor):
 *   1. Add this file. Save.
 *   2. Run setupRebuildCheckboxes() once (or Master menu → GNPA League →
 *      "Set up mobile refresh checkboxes"). Authorize when prompted.
 *   No redeploy needed — installable triggers always run the latest saved code.
 *   Re-running setup is safe/idempotent (skips leagues that already have a
 *   trigger; rebuilds the Refresh tab layout if someone mangled it). Run it
 *   again whenever a NEW league is added to Config.
 *
 * DESIGN NOTES (each one matters — don't "simplify" them away):
 *   - The checkbox lives on its own "Refresh" tab because writeStandingsTab()
 *     does sheet.clear() on the standings tabs — anything placed there is
 *     wiped on every rebuild.
 *   - The handler must NOT use SpreadsheetApp.getActiveSpreadsheet() (directly
 *     or via lookupLeagueSheetId/getAllConfiguredLeagues): in a trigger fired
 *     on a league sheet, "active" is ambiguous and may not be the Master. It
 *     reads Config via MASTER_SPREADSHEET_ID and rebuilds via e.source.
 *   - Unticking the box with setValue() cannot loop: programmatic edits never
 *     fire onEdit triggers.
 *   - The trigger runs as the account that ran setup (Jade), so league
 *     managers can tick the box with plain edit access — no auth prompt.
 */

var MASTER_SPREADSHEET_ID = '1B-5USbnz5bxt7TBSGY0n9PQVvAMb4CttEXQWOHN8vb4';
var REFRESH_TAB = 'Refresh';
var REFRESH_CHECKBOX_CELL = 'B1';
var REFRESH_STATUS_CELL = 'B2';
var REFRESH_HANDLER = 'onLeagueRefreshEdit';

// ------------------------------------------------------------
// One-time setup: Refresh tab + installable trigger per league
// ------------------------------------------------------------
function setupRebuildCheckboxes() {
  var leagues = configLeaguesFromMaster_();
  if (!leagues.length) {
    reportSetup_('No leagues configured in the Master Config tab.');
    return;
  }

  // Which league sheets already have our trigger?
  var existing = {};
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === REFRESH_HANDLER) {
      existing[t.getTriggerSourceId()] = true;
    }
  });

  var report = [];
  leagues.forEach(function (entry) {
    try {
      var ss = SpreadsheetApp.openById(entry.sheetId);
      ensureRefreshTab_(ss);
      if (!existing[entry.sheetId]) {
        ScriptApp.newTrigger(REFRESH_HANDLER)
          .forSpreadsheet(entry.sheetId)
          .onEdit()
          .create();
        report.push('✓ ' + entry.league + ' — tab + trigger created');
      } else {
        report.push('✓ ' + entry.league + ' — already set up');
      }
    } catch (e) {
      report.push('✗ ' + entry.league + ' — ' + e.message);
    }
  });
  reportSetup_('Mobile refresh checkboxes:\n\n' + report.join('\n'));
}

function ensureRefreshTab_(ss) {
  var sheet = ss.getSheetByName(REFRESH_TAB);
  if (!sheet) {
    sheet = ss.insertSheet(REFRESH_TAB, 0); // first tab — what mobile opens to
  } else {
    ss.setActiveSheet(sheet);
    ss.moveActiveSheet(1);
  }
  sheet.getRange('A1').setValue('Tick the box to refresh standings →').setFontWeight('bold');
  var box = sheet.getRange(REFRESH_CHECKBOX_CELL);
  box.insertCheckboxes();
  if (box.getValue() !== true) box.setValue(false);
  sheet.getRange('A2').setValue('Status');
  if (!sheet.getRange(REFRESH_STATUS_CELL).getValue()) {
    sheet.getRange(REFRESH_STATUS_CELL).setValue('Never run from here yet');
  }
  sheet.getRange('A4').setValue(
    'Rebuild takes ~20 seconds. The box unticks itself when done — '
    + 'check Status, then look at the standings tabs.'
  ).setFontStyle('italic').setFontColor('#666666');
  sheet.setColumnWidth(1, 280);
  sheet.setColumnWidth(2, 260);
}

// ------------------------------------------------------------
// The trigger handler (runs on EVERY edit in a league sheet —
// bail out fast unless it's the checkbox being ticked)
// ------------------------------------------------------------
function onLeagueRefreshEdit(e) {
  try {
    if (!e || !e.range) return;
    var sheet = e.range.getSheet();
    // Manager-tools tab shares this same installable trigger (manager.gs).
    if (sheet.getName() === MANAGER_TAB) { onManagerEdit_(e, sheet); return; }
    if (sheet.getName() !== REFRESH_TAB) return;
    if (e.range.getA1Notation() !== REFRESH_CHECKBOX_CELL) return;
    if (e.range.getValue() !== true) return; // untick or non-checkbox edit

    var ss = e.source;
    var status = sheet.getRange(REFRESH_STATUS_CELL);
    var league = leagueNameForSheetId_(ss.getId());
    if (!league) {
      status.setValue('✗ This sheet is not in the Master Config — cannot rebuild');
      e.range.setValue(false);
      return;
    }

    // One rebuild at a time across the whole script (rebuilds share tabs/cache).
    var lock = LockService.getScriptLock();
    if (!lock.tryLock(5000)) {
      status.setValue('⏳ A rebuild is already running — try again in a minute');
      e.range.setValue(false);
      return;
    }
    try {
      status.setValue('⏳ Rebuilding ' + league + ' standings…');
      SpreadsheetApp.flush();
      // Same path as the Master menu rebuild — seed merge included — but on
      // the event's own spreadsheet (no getActiveSpreadsheet in trigger context).
      rebuildTeamStandings(ss, league);
      rebuildPlayerStandings(ss, league);
      try { CacheService.getScriptCache().remove('standings:' + league); } catch (ce) {}
      var tz = ss.getSpreadsheetTimeZone();
      status.setValue('✓ Rebuilt ' + Utilities.formatDate(new Date(), tz, 'EEE d MMM HH:mm'));
    } catch (err) {
      status.setValue('✗ Failed: ' + err.message);
    } finally {
      try { lock.releaseLock(); } catch (le) {}
      e.range.setValue(false); // safe: script edits never re-fire onEdit
    }
  } catch (outer) {
    // Never throw from a trigger — failures land in the status cell above.
  }
}

// ------------------------------------------------------------
// Helpers (trigger-safe: Master by ID, never "active" spreadsheet)
// ------------------------------------------------------------
function configLeaguesFromMaster_() {
  var master = SpreadsheetApp.openById(MASTER_SPREADSHEET_ID);
  var sheet = master.getSheetByName(CONFIG_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return [];
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  var out = [];
  for (var i = 0; i < data.length; i++) {
    var name = String(data[i][0] || '').trim();
    var id = String(data[i][1] || '').trim();
    if (name && id) out.push({ league: name, sheetId: id });
  }
  return out;
}

function leagueNameForSheetId_(sheetId) {
  var leagues = configLeaguesFromMaster_();
  for (var i = 0; i < leagues.length; i++) {
    if (leagues[i].sheetId === sheetId) return leagues[i].league;
  }
  return null;
}

function reportSetup_(msg) {
  // Works from the menu (UI alert) and from the bare script editor (log).
  try {
    SpreadsheetApp.getUi().alert(msg);
  } catch (e) {
    Logger.log(msg);
  }
}

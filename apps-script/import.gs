/**
 * GNPA — "Import Players" tab: bulk enrolment import, done by the manager
 * (added 2026-09-20)
 *
 * WHY: at the start of a season the league admin sends an enrolment
 * spreadsheet — one row per enrolled player, grouped into team blocks with a
 * blank row between them:
 *
 *     Club | Team | Player | ID Number | Phone | Gender | Race
 *
 * Someone then has to turn ~120 of those rows into a Roster tab and a set of
 * team codes. Doing it by hand is slow and the file is never clean. The real
 * 2026 Tshwane enrolment (119 rows) carried, all at once: trailing spaces on
 * nearly every club/team/name, two players entered twice, five blank rows
 * under a team name that differed from a real team only by a trailing space,
 * six ID numbers whose leading zero Excel had eaten, three IDs that fail their
 * checksum, one 14-digit ID, and eleven cells with a stray "*" pasted in.
 *
 * So this is NOT a "load the file" button. It is a two-tick tool:
 *
 *     Tick CHECK   → nothing is written. Every row gets a plain-English note
 *                    in the "What to fix" column saying what is wrong with it,
 *                    and B7 summarises. The manager fixes the rows and ticks
 *                    Check again until it says Ready.
 *     Tick IMPORT  → writes the Roster tab and tops up Team Codes. Refuses
 *                    while any ⛔ row remains, and says how many.
 *
 * THREE SEVERITIES (the distinction is the whole point — see impAnalyse_):
 *     ⛔ must fix   the app would end up with broken data. Blocks the import.
 *     ⚠ check      probably a typo, but harmless to the app. Import proceeds.
 *     ✎ cleaned    fixed automatically, reported so nothing happens silently.
 *
 * WHAT IT WRITES
 *     Roster      Team | Player | Captain? | Code | Club | ID Number | Phone
 *                 Columns 5-7 are EXTRA. readRosterTab() in roster.gs reads
 *                 only the first 4, so the app ignores them — and the public
 *                 ?action=roster endpoint returns names only, never codes,
 *                 IDs or phone numbers. ID/phone stay inside the sheet.
 *     Team Codes  topped up via setupTeamCodesForSheet() (roster.gs), which
 *                 never overwrites a code that already exists.
 *
 * MECHANISM: same installable onEdit trigger as Refresh and Manager
 * (onLeagueRefreshEdit in refresh.gs dispatches Import-tab edits here). No new
 * trigger, so this stays under the Apps Script trigger cap, and it works in
 * the Sheets MOBILE app where custom menus do not render.
 *
 * SETUP (one-time, from the Master sheet's script editor):
 *   1. Add this file + the 1-line dispatch in refresh.gs. Save.
 *   2. Run setupImportTabs()  (or Master menu → GNPA League →
 *      "Set up player import tabs"). Authorize if prompted.
 *   No redeploy needed — installable triggers always run the latest code.
 *   Idempotent: re-run any time, including after adding a league to Config.
 *
 * ACCENTS ARE SAFE HERE. Bridge/API writes double-encode non-ASCII (é → Ã©),
 * which is why roster edits from tooling have to go in as =UNICHAR(). This
 * runs inside Apps Script on values a human pasted into the sheet, so
 * Amandré / Séthembiso / Sugnét / Hélena survive untouched. Do not "helpfully"
 * route this through the bridge later.
 */

var IMPORT_TAB = 'Import Players';

var IMP_MODE    = 'B4';  // Replace / Add dropdown
var IMP_CHECK   = 'B5';  // checkbox — validate only, writes nothing
var IMP_IMPORT  = 'B6';  // checkbox — commit
var IMP_STATUS  = 'B7';  // summary line

var IMP_HEADER_ROW = 9;   // the paste area's own header
var IMP_FIRST_ROW  = 10;  // first pasted data row
var IMP_ISSUE_COL  = 8;   // column H — "What to fix"

var IMP_HEADERS = ['Club', 'Team', 'Player', 'ID Number', 'Phone', 'Gender', 'Race', 'What to fix'];

var IMP_MODE_REPLACE = 'Replace the whole roster (new season)';
var IMP_MODE_ADD     = 'Add to the existing roster (mid-season)';

// Squad size per league, mirrored from LEAGUES in index.html. Used ONLY to
// warn about a short squad — never to block. If a league is missing here the
// fallback is 5, which under-warns rather than nagging about a valid squad.
var IMP_SQUAD_MIN = {
  'vets tier 1': 6, 'vets tier 2': 6, 'super north': 6, 'super south': 6,
  'premier': 6, 'ladies': 6, 'tshwane': 5, '3-man': 3, 'juniors': 2
};

// ============================================================
//  One-time setup: build the Import tab on every league sheet
// ============================================================
function setupImportTabs() {
  var leagues = configLeaguesFromMaster_();
  if (!leagues.length) {
    reportSetup_('No leagues configured in the Master Config tab.');
    return;
  }
  var report = [];
  leagues.forEach(function (entry) {
    try {
      var ss = SpreadsheetApp.openById(entry.sheetId);
      ensureImportTab_(ss);
      ensureLeagueEditTrigger_(entry.sheetId); // shares the Refresh handler
      report.push('✓ ' + entry.league);
    } catch (e) {
      report.push('✗ ' + entry.league + ' — ' + e.message);
    }
  });
  reportSetup_(
    'Player import tabs:\n\n' + report.join('\n') +
    '\n\nOn each league sheet, open "' + IMPORT_TAB + '", paste the enrolment ' +
    'rows from row ' + IMP_FIRST_ROW + ' down, then tick Check.'
  );
}

function ensureImportTab_(ss) {
  var sheet = ss.getSheetByName(IMPORT_TAB);
  var keep = null;

  if (!sheet) {
    sheet = ss.insertSheet(IMPORT_TAB, 2); // third tab: Refresh, Manager, Import
  } else {
    // Re-running setup must not throw away rows a manager has already pasted.
    keep = impReadPastedBlock_(sheet);
    ss.setActiveSheet(sheet);
    ss.moveActiveSheet(3);
    sheet.clear();
    sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns())
         .clearDataValidations().removeCheckboxes();
  }

  sheet.setColumnWidth(1, 170);
  sheet.setColumnWidth(2, 300);
  for (var c = 3; c <= 7; c++) sheet.setColumnWidth(c, 150);
  sheet.setColumnWidth(IMP_ISSUE_COL, 460);
  sheet.setHiddenGridlines(true);

  band_(sheet, 'A1:D1', '📥  IMPORT ENROLLED PLAYERS', '#1f6fc4', '#ffffff', 12);
  note_(sheet, 'A2:H2',
    'Paste the enrolment sheet into row ' + IMP_FIRST_ROW + ' and below, in the column order shown on ' +
    'row ' + IMP_HEADER_ROW + '. Blank rows between teams are fine. Then tick Check — nothing is ' +
    'written until you tick Import.');

  label_(sheet, 'A4', 'Mode');
  sheet.getRange(IMP_MODE)
    .setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList([IMP_MODE_REPLACE, IMP_MODE_ADD], true)
      .setAllowInvalid(false).setHelpText('Replace wipes the current roster first').build())
    .setValue(IMP_MODE_ADD);

  label_(sheet, 'A5', '1. Tick to CHECK  →');
  sheet.getRange(IMP_CHECK).insertCheckboxes().setValue(false);
  label_(sheet, 'A6', '2. Tick to IMPORT  →');
  sheet.getRange(IMP_IMPORT).insertCheckboxes().setValue(false);
  label_(sheet, 'A7', 'Result');
  status_(sheet, IMP_STATUS, 'Paste the rows below, then tick Check');
  sheet.getRange(IMP_STATUS).setWrap(true);
  sheet.setRowHeight(7, 46);

  sheet.getRange(IMP_HEADER_ROW, 1, 1, IMP_HEADERS.length)
    .setValues([IMP_HEADERS]).setFontWeight('bold')
    .setBackground('#e8f0fe').setFontColor('#1a1a1a');
  sheet.setFrozenRows(IMP_HEADER_ROW);

  // ID numbers are text: a 13-digit number loses its leading zero and turns
  // into scientific notation the moment Sheets treats it as a number.
  var body = Math.max(sheet.getMaxRows() - IMP_HEADER_ROW, 1);
  sheet.getRange(IMP_FIRST_ROW, 4, body, 2).setNumberFormat('@');
  sheet.getRange(IMP_FIRST_ROW, IMP_ISSUE_COL, body, 1).setWrap(true);

  if (keep && keep.length) {
    sheet.getRange(IMP_FIRST_ROW, 1, keep.length, 7)
      .setValues(keep.map(function (r) { return r.slice(0, 7); }));
  }
  return sheet;
}

// The pasted block as raw values (columns A-G), trailing blank rows trimmed.
function impReadPastedBlock_(sheet) {
  var last = sheet.getLastRow();
  if (last < IMP_FIRST_ROW) return [];
  var vals = sheet.getRange(IMP_FIRST_ROW, 1, last - IMP_FIRST_ROW + 1, 7).getValues();
  while (vals.length && vals[vals.length - 1].every(function (v) { return String(v || '').trim() === ''; })) {
    vals.pop();
  }
  return vals;
}

// ============================================================
//  Edit handler — dispatched from onLeagueRefreshEdit (refresh.gs)
// ============================================================
function onImportEdit_(e, sheet) {
  try {
    var a1 = e.range.getA1Notation();
    if (a1 === IMP_CHECK  && e.range.getValue() === true) { impDoCheck_(e, sheet);  return; }
    if (a1 === IMP_IMPORT && e.range.getValue() === true) { impDoImport_(e, sheet); return; }
  } catch (err) {
    try {
      sheet.getRange(IMP_CHECK).setValue(false);
      sheet.getRange(IMP_IMPORT).setValue(false);
      sheet.getRange(IMP_STATUS).setValue('✗ ' + err.message);
    } catch (ignore) {}
  }
}

// ============================================================
//  Reading + cleaning
// ============================================================
function impClean_(v) {
  return String(v == null ? '' : v).replace(/\s+/g, ' ').trim();
}

// Digits only — strips the stray "*", spaces, dashes and brackets people paste.
function impDigits_(v) {
  return String(v == null ? '' : v).replace(/\D/g, '');
}

// SA ID numbers carry a Luhn check digit over all 13 digits.
function impLuhn_(s) {
  var total = 0;
  for (var i = 0; i < s.length; i++) {
    var d = parseInt(s.charAt(s.length - 1 - i), 10);
    if (i % 2 === 1) { d *= 2; if (d > 9) d -= 9; }
    total += d;
  }
  return total % 10 === 0;
}

// First 6 digits are YYMMDD. Catches transposed digits the checksum misses.
function impDobOk_(s) {
  var mm = parseInt(s.substr(2, 2), 10);
  var dd = parseInt(s.substr(4, 2), 10);
  return mm >= 1 && mm <= 12 && dd >= 1 && dd <= 31;
}

function impSquadMin_(league) {
  var n = IMP_SQUAD_MIN[String(league || '').trim().toLowerCase()];
  return n || 5;
}

// Turn the pasted block into records, dropping blank separator rows and a
// pasted header row. Row numbers are SHEET rows so the notes point somewhere.
function impParse_(sheet) {
  var raw = impReadPastedBlock_(sheet);
  var out = [];
  for (var i = 0; i < raw.length; i++) {
    var r = raw[i];
    var rec = {
      row:    IMP_FIRST_ROW + i,
      club:   impClean_(r[0]),
      team:   impClean_(r[1]),
      name:   impClean_(r[2]),
      idRaw:  impClean_(r[3]),
      telRaw: impClean_(r[4]),
      gender: impClean_(r[5]),
      race:   impClean_(r[6]),
      dirty:  false,   // something was auto-cleaned
      notes:  []
    };
    if (!rec.club && !rec.team && !rec.name && !rec.idRaw && !rec.telRaw) continue;

    // A pasted header row ("Club | Team | Player | ...") is not a player.
    var n = rec.name.toLowerCase();
    if (n === 'player' || n === 'name' || n === 'full name' || n === 'player name') continue;

    // Did cleaning change anything? Compare against the untouched cells.
    if (String(r[0] || '') !== rec.club || String(r[1] || '') !== rec.team ||
        String(r[2] || '') !== rec.name) {
      rec.dirty = true;
    }
    out.push(rec);
  }
  return out;
}

// ============================================================
//  The rules
// ============================================================
// Returns { records, teams, nTeams, nPlayers, blockers, warnings, fixes }.
// Each record gains .notes (strings, already severity-prefixed), .blocked,
// .skip (blank team stub — not imported, not an error) plus the cleaned
// .id / .tel it would be imported with.
function impAnalyse_(records, league) {
  var byTeamKey = {};   // lower-case team -> { display, recs }
  var i, rec;

  // ---- pass 1: per-row cleaning of ID + phone ----
  for (i = 0; i < records.length; i++) {
    rec = records[i];

    // ID number
    rec.id = '';
    var idRaw = rec.idRaw;
    var noId = /^(no\s*id|none|n\/?a|-)$/i.test(idRaw);
    if (!idRaw) {
      if (rec.name) rec.notes.push('⚠ No ID number');
    } else if (noId) {
      rec.notes.push('⚠ ID recorded as "' + idRaw + '" — get the real one before the season starts');
    } else {
      var digits = impDigits_(idRaw);
      if (digits !== idRaw) rec.dirty = true;

      if (digits.length === 12 && impLuhn_('0' + digits)) {
        // Excel ate the leading zero of an otherwise perfect ID. Very common.
        rec.id = '0' + digits;
        rec.notes.push('✎ ID had a missing leading zero — restored as ' + rec.id);
        rec.dirty = true;
      } else if (digits.length === 13) {
        rec.id = digits;
        if (!impLuhn_(digits)) {
          rec.notes.push('⚠ ID ' + digits + ' fails its check digit — one digit is probably wrong');
        } else if (!impDobOk_(digits)) {
          rec.notes.push('⚠ ID ' + digits + ' starts with an impossible date of birth');
        } else if (digits !== idRaw) {
          rec.notes.push('✎ Stray characters removed from the ID ("' + idRaw + '" → ' + digits + ')');
        }
      } else {
        rec.id = digits;
        rec.notes.push('⚠ ID "' + idRaw + '" has ' + digits.length +
                       ' digits — an SA ID has 13. Check it against their card');
      }
    }

    // Phone
    rec.tel = impDigits_(rec.telRaw);
    if (rec.tel.length === 11 && rec.tel.indexOf('27') === 0) rec.tel = '0' + rec.tel.substr(2);
    if (rec.telRaw && rec.tel !== rec.telRaw) {
      rec.notes.push('✎ Phone cleaned up ("' + rec.telRaw + '" → ' + rec.tel + ')');
      rec.dirty = true;
    }
    if (rec.name && !rec.tel) {
      rec.notes.push('⚠ No phone number');
    } else if (rec.tel && (rec.tel.length !== 10 || rec.tel.charAt(0) !== '0')) {
      rec.notes.push('⚠ Phone ' + rec.tel + ' is ' + rec.tel.length +
                     ' digits — expected 10, starting with 0');
    }

    // Structure
    if (!rec.team && rec.name) {
      rec.notes.push('⛔ No team on this row — which team is ' + rec.name + ' playing for?');
      rec.blocked = true;
    }
    if (rec.team && !rec.name) {
      // A team with the player column left empty. Imports nothing either way;
      // pass 3 decides which of the two notes explains it.
      rec.skip = true;
    }

    if (rec.team) {
      var key = rec.team.toLowerCase();
      if (!byTeamKey[key]) byTeamKey[key] = { display: rec.team, recs: [] };
      byTeamKey[key].recs.push(rec);
    }
  }

  // ---- pass 2: duplicates within a team, and across teams ----
  var byIdAll = {};    // cleaned ID -> [rec]
  var byNameAll = {};  // lower-case name -> [rec]

  // Same name twice in one team block: one person entered twice. Unambiguous,
  // so it is auto-dropped rather than blocking the manager. (One ID shared by
  // two different names is caught file-wide below, which reports it better.)
  Object.keys(byTeamKey).forEach(function (key) {
    var seenName = {};
    byTeamKey[key].recs.forEach(function (r) {
      if (!r.name) return;
      var nk = r.name.toLowerCase();
      if (seenName[nk]) {
        r.dropDuplicate = true;
        r.notes.push('✎ Same as row ' + seenName[nk] + ' — duplicate entry, only one will be imported');
      } else {
        seenName[nk] = r.row;
      }
    });
  });

  records.forEach(function (r) {
    if (!r.name || r.dropDuplicate) return;
    var nk = r.name.toLowerCase();
    (byNameAll[nk] = byNameAll[nk] || []).push(r);
    if (r.id && r.id.length === 13) (byIdAll[r.id] = byIdAll[r.id] || []).push(r);
  });

  // The same person enrolled for two teams — standings would count them for
  // both, and the app would offer them in two dropdowns. Always a blocker.
  Object.keys(byNameAll).forEach(function (nk) {
    var list = byNameAll[nk];
    var teams = {};
    list.forEach(function (r) { if (r.team) teams[r.team.toLowerCase()] = r.team; });
    if (Object.keys(teams).length > 1) {
      var names = Object.keys(teams).map(function (k) { return teams[k]; }).join('" and "');
      list.forEach(function (r) {
        r.notes.push('⛔ ' + r.name + ' is enrolled for both "' + names + '" — pick one');
        r.blocked = true;
      });
    }
  });
  Object.keys(byIdAll).forEach(function (id) {
    var list = byIdAll[id];
    var names = {};
    list.forEach(function (r) { names[r.name.toLowerCase()] = r.name; });
    if (Object.keys(names).length > 1) {
      var shown = Object.keys(names).map(function (k) { return names[k]; }).join('" / "');
      list.forEach(function (r) {
        r.notes.push('⛔ ID ' + id + ' is used by "' + shown + '" — same ID, different names');
        r.blocked = true;
      });
    }
  });

  // ---- pass 3: per-team checks ----
  var min = impSquadMin_(league);
  var teams = [];
  Object.keys(byTeamKey).forEach(function (key) {
    var t = byTeamKey[key];
    var players = t.recs.filter(function (r) { return r.name && !r.dropDuplicate && !r.blocked; });
    var named   = t.recs.filter(function (r) { return r.name; });

    if (!named.length) {
      // A team name with nothing under it at all.
      t.recs.forEach(function (r) {
        r.skip = true;
        r.notes.push('⚠ "' + t.display + '" has a team name but no players — this block will be skipped');
      });
    } else {
      // Blank rows inside a block that otherwise has players. The Tshwane file
      // had five of these under "Legends Westside Snipers " (trailing space),
      // which cleaning merges into the real block of the same name — so the
      // rows quietly do nothing. Say so; never drop a pasted row in silence.
      t.recs.forEach(function (r) {
        if (!r.name) {
          r.notes.push('⚠ No player name on this row — skipped (the rest of "' +
                       t.display + '" imports fine)');
        }
      });
    }
    // Short squad. Skipped when the team already has a ⛔ row: "0 player(s)"
    // stacked on top of the real error is noise, and the count is meaningless
    // until the blocker is resolved.
    var hasBlocked = t.recs.some(function (r) { return r.blocked; });
    if (named.length && !hasBlocked && players.length < min) {
      named[0].notes.push('⚠ ' + t.display + ' has only ' + players.length +
                          ' player(s) — this league plays ' + min + ' a side');
    }
    teams.push({ name: t.display, players: players });
  });

  teams.sort(function (a, b) { return a.name.toLowerCase() < b.name.toLowerCase() ? -1 : 1; });

  var blockers = 0, warnings = 0, fixes = 0, nPlayers = 0;
  records.forEach(function (r) {
    var hasBlock = false, hasWarn = false, hasFix = false;
    r.notes.forEach(function (n) {
      if (n.charAt(0) === '⛔') hasBlock = true;
      else if (n.charAt(0) === '⚠') hasWarn = true;
      else hasFix = true;
    });
    if (hasBlock) blockers++;
    if (hasWarn) warnings++;
    if (hasFix) fixes++;
  });
  teams.forEach(function (t) { nPlayers += t.players.length; });

  return {
    records: records,
    teams: teams.filter(function (t) { return t.players.length; }),
    nTeams: teams.filter(function (t) { return t.players.length; }).length,
    nPlayers: nPlayers,
    blockers: blockers,
    warnings: warnings,
    fixes: fixes
  };
}

// ============================================================
//  Check (writes notes only — never touches the roster)
// ============================================================
function impDoCheck_(e, sheet) {
  var status = sheet.getRange(IMP_STATUS);
  e.range.setValue(false);

  var league = leagueNameForSheetId_(e.source.getId());
  var records = impParse_(sheet);
  if (!records.length) {
    impClearNotes_(sheet);
    status.setValue('✗ Nothing pasted yet — put the enrolment rows in row ' +
                    IMP_FIRST_ROW + ' and below, then tick Check again');
    return;
  }

  var res = impAnalyse_(records, league);
  impWriteNotes_(sheet, res);

  var counts = res.nTeams + ' team(s), ' + res.nPlayers + ' player(s)';
  if (res.blockers) {
    status.setValue('⛔ ' + res.blockers + ' row(s) must be fixed before importing — see the ' +
      '"What to fix" column. Also: ' + res.warnings + ' to check, ' + res.fixes +
      ' cleaned automatically. Would import ' + counts + '.');
  } else {
    status.setValue('✓ Ready to import — ' + counts + '. ' + res.warnings +
      ' row(s) worth checking, ' + res.fixes + ' cleaned automatically. ' +
      'Read those, then tick Import.');
  }
}

function impClearNotes_(sheet) {
  var last = sheet.getLastRow();
  if (last < IMP_FIRST_ROW) return;
  var n = last - IMP_FIRST_ROW + 1;
  sheet.getRange(IMP_FIRST_ROW, IMP_ISSUE_COL, n, 1)
    .clearContent().setBackground(null).setFontColor(null);
}

function impWriteNotes_(sheet, res) {
  impClearNotes_(sheet);
  var last = sheet.getLastRow();
  if (last < IMP_FIRST_ROW) return;

  var n = last - IMP_FIRST_ROW + 1;
  var text = [], bg = [], fg = [];
  for (var i = 0; i < n; i++) { text.push(['']); bg.push([null]); fg.push([null]); }

  res.records.forEach(function (r) {
    var i = r.row - IMP_FIRST_ROW;
    if (i < 0 || i >= n || !r.notes.length) return;
    text[i][0] = r.notes.join('\n');
    if (r.blocked)                         { bg[i][0] = '#fce8e6'; fg[i][0] = '#a50e0e'; }
    else if (r.notes.join('').indexOf('⚠') >= 0) { bg[i][0] = '#fef7e0'; fg[i][0] = '#8a6100'; }
    else                                   { bg[i][0] = '#e6f4ea'; fg[i][0] = '#137333'; }
  });

  var range = sheet.getRange(IMP_FIRST_ROW, IMP_ISSUE_COL, n, 1);
  range.setValues(text);
  range.setBackgrounds(bg);
  range.setFontColors(fg);
}

// ============================================================
//  Import (the only path that writes)
// ============================================================
function impDoImport_(e, sheet) {
  var ss = e.source;
  var status = sheet.getRange(IMP_STATUS);
  e.range.setValue(false);

  var league = leagueNameForSheetId_(ss.getId());
  if (!league) { status.setValue('✗ This sheet is not in the Master Config'); return; }

  var records = impParse_(sheet);
  if (!records.length) {
    status.setValue('✗ Nothing pasted yet — put the enrolment rows in row ' +
                    IMP_FIRST_ROW + ' and below');
    return;
  }

  // Always re-run the rules: the manager may have edited rows since Check, and
  // importing against a stale verdict is how bad data gets in.
  var res = impAnalyse_(records, league);
  impWriteNotes_(sheet, res);

  if (res.blockers) {
    status.setValue('⛔ Not imported — ' + res.blockers + ' row(s) still need fixing. ' +
      'Look for the red rows in "What to fix", sort them out, then tick Check.');
    return;
  }
  if (!res.nPlayers) {
    status.setValue('✗ Not imported — no usable players found. Check the column order matches row ' +
                    IMP_HEADER_ROW + '.');
    return;
  }

  var mode = impClean_(sheet.getRange(IMP_MODE).getValue()) || IMP_MODE_ADD;
  var replace = (mode === IMP_MODE_REPLACE);

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    status.setValue('⏳ Something else is running on this sheet — try again in a minute');
    return;
  }
  try {
    status.setValue('⏳ Importing ' + res.nPlayers + ' player(s)…');
    SpreadsheetApp.flush();

    var ros = impEnsureRosterTab_(ss);
    var removed = 0;
    var existing = {};   // "team|player" lower-case -> true

    if (replace) {
      if (ros.getLastRow() > 1) {
        removed = ros.getLastRow() - 1;
        ros.getRange(2, 1, removed, ros.getMaxColumns()).clearContent();
      }
    } else {
      readRosterTab(ss).forEach(function (r) {
        existing[(r.team + '|' + r.player).toLowerCase()] = true;
      });
    }

    var rows = [];
    res.teams.forEach(function (t) {
      t.players.forEach(function (p) {
        if (!replace && existing[(t.name + '|' + p.name).toLowerCase()]) return;
        rows.push([t.name, p.name, false, '', p.club, p.id, p.tel]);
      });
    });

    if (rows.length) {
      var start = Math.max(ros.getLastRow() + 1, 2);
      ros.getRange(start, 4, rows.length, 1).setNumberFormat('@'); // Code keeps leading zeros
      ros.getRange(start, 6, rows.length, 2).setNumberFormat('@'); // ID + phone are text
      ros.getRange(start, 1, rows.length, 7).setValues(rows);
      ros.getRange(start, 3, rows.length, 1).insertCheckboxes();
    }

    // In Replace mode, drop code rows for teams that are gone AND never played
    // — otherwise last season's teams keep showing in the Manager tab's team
    // picker, which reads Team Codes!A2:A live. A team with matches on the
    // Players tab is never pruned, whatever the roster says.
    var pruned = replace ? impPruneTeamCodes_(ss, res.teams) : 0;

    // Issues a code to every team that does not have one. Existing codes are
    // never overwritten (roster.gs), so mid-season codes keep working.
    var codes = setupTeamCodesForSheet(ss);

    try { invalidateLeagueCaches_(league); } catch (ce) {}

    var tz = ss.getSpreadsheetTimeZone();
    status.setValue('✓ Imported ' + rows.length + ' player(s) across ' + res.nTeams + ' team(s)' +
      (replace ? ' (replaced ' + removed + ' old roster row(s)' + (pruned ? ', removed ' + pruned + ' unused team code(s)' : '') + ')'
               : ' (' + (res.nPlayers - rows.length) + ' already on the roster)') +
      '. ' + codes.added + ' new team code(s) issued — open the Team Codes tab and send each team its code. ' +
      'Standings are not affected. · ' + Utilities.formatDate(new Date(), tz, 'EEE d MMM HH:mm'));
  } catch (err) {
    status.setValue('✗ Import failed: ' + err.message + ' — nothing further was written');
  } finally {
    try { lock.releaseLock(); } catch (le) {}
  }
}

function impEnsureRosterTab_(ss) {
  var sh = ss.getSheetByName(ROSTER_TAB);
  var headers = ROSTER_HEADERS.concat(['Club', 'ID Number', 'Phone']);
  if (!sh) {
    sh = ss.insertSheet(ROSTER_TAB);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 240);
    sh.setColumnWidth(2, 200);
  } else if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
  } else if (sh.getLastColumn() < headers.length) {
    // Existing 4-column Roster: widen it without disturbing columns 1-4.
    sh.getRange(1, 5, 1, 3).setValues([['Club', 'ID Number', 'Phone']]).setFontWeight('bold');
  }
  return sh;
}

// Remove Team Codes rows for teams not in the new roster and with no games on
// the Players tab. Returns how many were removed.
function impPruneTeamCodes_(ss, teams) {
  var sh = ss.getSheetByName(TEAM_CODES_TAB);
  if (!sh || sh.getLastRow() < 2) return 0;
  var teamCol = colByHeader_(sh, 'Team') || 1;

  var keep = {};
  teams.forEach(function (t) { keep[t.name.toLowerCase()] = true; });

  var played = {};
  var pl = ss.getSheetByName(PLAYERS_TAB);
  if (pl && pl.getLastRow() > 1) {
    var tc = colByHeader_(pl, 'Team');
    if (tc) {
      pl.getRange(2, tc, pl.getLastRow() - 1, 1).getValues().forEach(function (r) {
        var t = impClean_(r[0]); if (t) played[t.toLowerCase()] = true;
      });
    }
  }

  var removed = 0;
  for (var r = sh.getLastRow(); r >= 2; r--) {
    var name = impClean_(sh.getRange(r, teamCol).getValue());
    if (!name) continue;
    var k = name.toLowerCase();
    if (!keep[k] && !played[k]) { sh.deleteRow(r); removed++; }
  }
  return removed;
}

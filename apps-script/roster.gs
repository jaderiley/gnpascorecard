/**
 * GNPA Pool League — Roster + Dual-Code Captain Gate
 *
 * This file is part of the MASTER sheet's Apps Script project. It lives
 * alongside Code.gs and seed.gs (Apps Script merges every .gs file into one
 * global scope, so the functions here are callable from Code.gs).
 *
 * WHAT IT ADDS
 *   1. A per-league "Roster" tab:  Team | Player | Captain? | Code
 *      - Every player on every team, loaded at season start.
 *      - "Captain?" is a checkbox marking that team's captain(s).
 *      - "Code" is a short, static, season-long password for each captain.
 *
 *   2. A dual-code submission gate (validateCaptainCodes): a match is only
 *      accepted if BOTH opposing captains enter their own code. One valid
 *      submission still produces exactly one pending row; the manager's
 *      Verified-checkbox approval is untouched.
 *
 *   3. Roster serving: the app's team/player dropdowns read this tab first.
 *
 *   4. Manager menu tools: set up Roster tabs, auto-generate codes, reset a code.
 *
 * APPLIES TO EVERY LEAGUE. Nothing here is league-specific. (The only
 * per-league branch in the whole system is the percentage formula in
 * Code.gs > rebuildPlayerStandings, which keys off the "Ladies" name.)
 */

// ============================================================
//  ROSTER TAB
// ============================================================
var ROSTER_TAB = 'Roster';
var ROSTER_HEADERS = ['Team', 'Player', 'Captain?', 'Code'];

// When false, leagues that have NOT set up captain codes yet keep submitting
// as before (graceful rollout — nothing breaks mid-season). As soon as a
// league has at least one captain with a code, the gate enforces for that
// league. Flip to true to require codes for EVERY league, no exceptions.
var ROSTER_GATE_STRICT = false;

// Auto-generated code settings. Ambiguous characters (0/O, 1/I/L) are omitted
// so codes are easy to read and type on a phone.
var CODE_ALPHABET = 'ABCDEFGHJKMNPQRSTUVWXYZ23456789';
var CODE_LENGTH = 4;

// ============================================================
//  ROSTER READING
// ============================================================

// Reads the Roster tab into [{ team, player, captain, code }]. Returns [] if
// the tab is missing or empty.
function readRosterTab(ss) {
  var sheet = ss.getSheetByName(ROSTER_TAB);
  if (!sheet || sheet.getLastRow() < 2) return [];
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, ROSTER_HEADERS.length).getValues();
  var out = [];
  for (var i = 0; i < data.length; i++) {
    var team = String(data[i][0] || '').trim();
    var player = String(data[i][1] || '').trim();
    if (!team && !player) continue;
    out.push({
      team: team,
      player: player,
      captain: data[i][2] === true || String(data[i][2]).toUpperCase() === 'TRUE',
      code: normCode(data[i][3])
    });
  }
  return out;
}

// Builds the roster payload the app uses for its dropdowns. Prefers the
// Roster tab; falls back to the legacy seed/Players derivation so leagues
// that haven't built a Roster tab yet keep working unchanged.
//   returns { teams: [..], rosters: { team: [players] }, captains: { team: [players] } }
function getRosterForLeague(ss) {
  var roster = readRosterTab(ss);

  if (roster.length) {
    var rosters = {};
    var captains = {};
    var teamSet = {};
    roster.forEach(function (r) {
      if (!r.team) return;
      teamSet[r.team] = true;
      if (r.player) {
        if (!rosters[r.team]) rosters[r.team] = [];
        rosters[r.team].push(r.player);
        if (r.captain) {
          if (!captains[r.team]) captains[r.team] = [];
          captains[r.team].push(r.player);
        }
      }
    });
    Object.keys(rosters).forEach(function (t) { rosters[t].sort(); });
    Object.keys(captains).forEach(function (t) { captains[t].sort(); });
    return { teams: Object.keys(teamSet).sort(), rosters: rosters, captains: captains };
  }

  // ---- Legacy fallback: derive from seed + submitted Players data ----
  var playerSeed = readPlayerSeed(ss);
  var teamSeed = readTeamSeed(ss);
  var rostersL = {};
  Object.keys(playerSeed).forEach(function (k) {
    var s = playerSeed[k];
    if (!rostersL[s.team]) rostersL[s.team] = [];
    rostersL[s.team].push(s.name);
  });
  Object.keys(rostersL).forEach(function (t) { rostersL[t].sort(); });
  var teamSetL = {};
  Object.keys(teamSeed).forEach(function (t) { teamSetL[t] = true; });
  Object.keys(rostersL).forEach(function (t) { teamSetL[t] = true; });
  return { teams: Object.keys(teamSetL).sort(), rosters: rostersL, captains: {} };
}

// ============================================================
//  DUAL-CODE GATE
// ============================================================

// Validates the two captain codes on an incoming submission.
//   returns { ok: true, homeCaptain, awayCaptain }  on success
//   returns { ok: false, message }                  on rejection
function validateCaptainCodes(ss, data) {
  var captains = readRosterTab(ss).filter(function (r) { return r.captain && r.code && r.team; });

  // No captain codes configured for this league yet.
  if (!captains.length) {
    if (ROSTER_GATE_STRICT) {
      return { ok: false, message: 'Captain codes are not set up for this league yet. Ask your league manager to set up the Roster.' };
    }
    return { ok: true }; // graceful: behave as before until codes exist
  }

  var home = String(data.homeTeam || '').trim();
  var away = String(data.awayTeam || '').trim();
  if (home && away && teamEq(home, away)) {
    return { ok: false, message: 'Home and away team are the same — pick two different teams.' };
  }

  var hc = normCode(data.homeCode);
  var ac = normCode(data.awayCode);
  if (!hc || !ac) {
    return { ok: false, message: 'Both captains must enter their code before submitting.' };
  }

  var hMatch = findCaptainByCode(captains, hc);
  var aMatch = findCaptainByCode(captains, ac);
  if (!hMatch) return { ok: false, message: 'Home captain code not recognised. Check it with your manager.' };
  if (!aMatch) return { ok: false, message: 'Away captain code not recognised. Check it with your manager.' };

  if (!teamEq(hMatch.team, home)) {
    return { ok: false, message: 'The home code belongs to "' + hMatch.team + '", not the home team "' + home + '".' };
  }
  if (!teamEq(aMatch.team, away)) {
    return { ok: false, message: 'The away code belongs to "' + aMatch.team + '", not the away team "' + away + '".' };
  }
  if (teamEq(hMatch.team, aMatch.team)) {
    return { ok: false, message: 'Both codes belong to the same team — each opposing captain must enter their own code.' };
  }

  // Authenticated. Use the roster's captain names as the source of truth.
  return { ok: true, homeCaptain: hMatch.player, awayCaptain: aMatch.player };
}

function normCode(v) {
  return String(v == null ? '' : v).trim().toUpperCase();
}

function teamEq(a, b) {
  return String(a || '').trim().toLowerCase() === String(b || '').trim().toLowerCase();
}

function findCaptainByCode(captains, code) {
  for (var i = 0; i < captains.length; i++) {
    if (captains[i].code === code) return captains[i];
  }
  return null;
}

// ============================================================
//  CODE GENERATION
// ============================================================

function makeCode(existingSet) {
  for (var attempt = 0; attempt < 200; attempt++) {
    var c = '';
    for (var i = 0; i < CODE_LENGTH; i++) {
      c += CODE_ALPHABET.charAt(Math.floor(Math.random() * CODE_ALPHABET.length));
    }
    if (!existingSet[c]) { existingSet[c] = true; return c; }
  }
  // Extremely unlikely fallback: lengthen until unique.
  var fallback = c + CODE_ALPHABET.charAt(Math.floor(Math.random() * CODE_ALPHABET.length));
  existingSet[fallback] = true;
  return fallback;
}

// Fills a blank Code cell for every Captain?=TRUE row in one league's Roster,
// keeping codes unique within that league. Existing codes are left untouched
// (so a manager-set or previously-issued code is never overwritten).
//   returns number of codes generated
function generateCodesForSheet(ss) {
  var sheet = ss.getSheetByName(ROSTER_TAB);
  if (!sheet || sheet.getLastRow() < 2) return 0;
  var range = sheet.getRange(2, 1, sheet.getLastRow() - 1, ROSTER_HEADERS.length);
  var values = range.getValues();

  var used = {};
  values.forEach(function (r) {
    var code = normCode(r[3]);
    if (code) used[code] = true;
  });

  var generated = 0;
  for (var i = 0; i < values.length; i++) {
    var isCaptain = values[i][2] === true || String(values[i][2]).toUpperCase() === 'TRUE';
    var hasCode = normCode(values[i][3]) !== '';
    if (isCaptain && !hasCode) {
      values[i][3] = makeCode(used);
      generated++;
    }
  }
  if (generated) range.setValues(values);
  return generated;
}

// ============================================================
//  ROSTER SET-UP / SEEDING
// ============================================================

// Creates (or tops up) the Roster tab for one league sheet, seeding Team +
// Player rows from whatever player data already exists (Roster tab itself,
// then seed, then submitted Players). Never deletes existing roster rows or
// codes; only appends players not already listed. "Captain?" starts unticked
// and "Code" blank — the manager ticks captains, then runs "Generate codes".
//   returns { created: bool, added: number }
function setupRosterForSheet(ss) {
  var sheet = ss.getSheetByName(ROSTER_TAB);
  var created = false;
  if (!sheet) {
    sheet = ss.insertSheet(ROSTER_TAB);
    sheet.getRange(1, 1, 1, ROSTER_HEADERS.length).setValues([ROSTER_HEADERS]).setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 220);
    sheet.setColumnWidth(2, 200);
    created = true;
  } else if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, ROSTER_HEADERS.length).setValues([ROSTER_HEADERS]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  // Existing (team|player) pairs already in the Roster, so we never duplicate.
  var existing = {};
  var current = readRosterTab(ss);
  current.forEach(function (r) { existing[(r.team + '|' + r.player).toLowerCase()] = true; });

  // Candidate players: from seed first, then submitted Players tab.
  var candidates = []; // [{team, player}]
  var seen = {};
  var addCand = function (team, player) {
    if (!team || !player) return;
    var key = (team + '|' + player).toLowerCase();
    if (seen[key]) return;
    seen[key] = true;
    candidates.push({ team: team, player: player });
  };

  var seed = readPlayerSeed(ss);
  Object.keys(seed).forEach(function (k) { addCand(seed[k].team, seed[k].name); });

  var playersSheet = ss.getSheetByName(PLAYERS_TAB);
  if (playersSheet && playersSheet.getLastRow() >= 2) {
    var pd = playersSheet.getRange(2, 1, playersSheet.getLastRow() - 1, PLAYERS_HEADERS.length).getValues();
    pd.forEach(function (row) { addCand(String(row[4] || '').trim(), String(row[6] || '').trim()); });
  }

  // Append the ones not already in the Roster, sorted by team then player.
  candidates.sort(function (a, b) {
    return a.team.localeCompare(b.team) || a.player.localeCompare(b.player);
  });
  var toAdd = candidates
    .filter(function (c) { return !existing[(c.team + '|' + c.player).toLowerCase()]; })
    .map(function (c) { return [c.team, c.player, false, '']; });

  if (toAdd.length) {
    var startRow = Math.max(sheet.getLastRow() + 1, 2);
    sheet.getRange(startRow, 1, toAdd.length, ROSTER_HEADERS.length).setValues(toAdd);
    // Make the Captain? column real checkboxes for the rows we just added.
    sheet.getRange(startRow, 3, toAdd.length, 1).insertCheckboxes();
  }

  return { created: created, added: toAdd.length };
}

// ============================================================
//  MENU HANDLERS  (wired up from Code.gs > onOpen)
// ============================================================

function setupAllRosters() {
  var leagues = getAllConfiguredLeagues();
  if (!leagues.length) {
    SpreadsheetApp.getUi().alert('No leagues configured. Fill in the Config tab first.');
    return;
  }
  var report = [];
  leagues.forEach(function (entry) {
    try {
      var ss = SpreadsheetApp.openById(entry.sheetId);
      var r = setupRosterForSheet(ss);
      report.push('✓ ' + entry.league + ' — ' + (r.created ? 'created, ' : '') + r.added + ' player(s) added');
    } catch (e) {
      report.push('✗ ' + entry.league + ' — ' + e.message);
    }
  });
  SpreadsheetApp.getUi().alert(
    'Roster set-up done:\n\n' + report.join('\n') +
    '\n\nNext: open each Roster tab, tick "Captain?" for each team\'s captain(s), ' +
    'then run "Generate captain codes".'
  );
}

function generateAllCodes() {
  var leagues = getAllConfiguredLeagues();
  if (!leagues.length) {
    SpreadsheetApp.getUi().alert('No leagues configured.');
    return;
  }
  var report = [];
  leagues.forEach(function (entry) {
    try {
      var ss = SpreadsheetApp.openById(entry.sheetId);
      var n = generateCodesForSheet(ss);
      report.push('✓ ' + entry.league + ' — ' + n + ' new code(s)');
    } catch (e) {
      report.push('✗ ' + entry.league + ' — ' + e.message);
    }
  });
  SpreadsheetApp.getUi().alert('Code generation done:\n\n' + report.join('\n'));
}

// Manager-friendly single-league actions (prompt for the league name).
function generateCodesPrompt() {
  var ui = SpreadsheetApp.getUi();
  var leagues = getAllConfiguredLeagues();
  if (!leagues.length) { ui.alert('No leagues configured.'); return; }
  var list = leagues.map(function (l, i) { return (i + 1) + '. ' + l.league; }).join('\n');
  var res = ui.prompt('Generate captain codes',
    'Type the league name exactly (or "ALL" for every league):\n\n' + list,
    ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  var name = res.getResponseText().trim();
  if (!name) return;
  if (name.toUpperCase() === 'ALL') { generateAllCodes(); return; }
  var sheetId = lookupLeagueSheetId(name);
  if (!sheetId) { ui.alert('Unknown league: ' + name); return; }
  try {
    var n = generateCodesForSheet(SpreadsheetApp.openById(sheetId));
    ui.alert('✓ ' + n + ' new code(s) generated for ' + name + '.');
  } catch (e) {
    ui.alert('Failed: ' + e.message);
  }
}

// Reset (re-issue) one captain's code, e.g. if it leaks.
function resetCaptainCodePrompt() {
  var ui = SpreadsheetApp.getUi();
  var leagues = getAllConfiguredLeagues();
  if (!leagues.length) { ui.alert('No leagues configured.'); return; }
  var lr = ui.prompt('Reset a captain code', 'League name:', ui.ButtonSet.OK_CANCEL);
  if (lr.getSelectedButton() !== ui.Button.OK) return;
  var league = lr.getResponseText().trim();
  var sheetId = lookupLeagueSheetId(league);
  if (!sheetId) { ui.alert('Unknown league: ' + league); return; }

  var pr = ui.prompt('Reset a captain code', 'Captain (player) name to re-issue a code for:', ui.ButtonSet.OK_CANCEL);
  if (pr.getSelectedButton() !== ui.Button.OK) return;
  var player = pr.getResponseText().trim();
  if (!player) return;

  var ss = SpreadsheetApp.openById(sheetId);
  var sheet = ss.getSheetByName(ROSTER_TAB);
  if (!sheet || sheet.getLastRow() < 2) { ui.alert('No Roster tab for ' + league + '.'); return; }
  var range = sheet.getRange(2, 1, sheet.getLastRow() - 1, ROSTER_HEADERS.length);
  var values = range.getValues();
  var used = {};
  values.forEach(function (r) { var c = normCode(r[3]); if (c) used[c] = true; });

  var done = false;
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][1] || '').trim().toLowerCase() === player.toLowerCase()) {
      var old = normCode(values[i][3]);
      if (old) delete used[old];
      values[i][2] = true; // ensure marked as captain
      values[i][3] = makeCode(used);
      done = true;
      range.setValues(values);
      ui.alert('✓ New code for ' + values[i][1] + ' (' + values[i][0] + '): ' + values[i][3]);
      break;
    }
  }
  if (!done) ui.alert('No roster row found for player "' + player + '" in ' + league + '.');
}

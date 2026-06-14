/**
 * GNPA Pool League — Team Roster + Dual Team-Code Submission Gate
 *
 * Part of the MASTER sheet's Apps Script project (lives alongside Code.gs and
 * seed.gs; Apps Script merges every .gs file into one global scope).
 *
 * WHAT IT ADDS
 *   1. A per-league "Team Codes" tab:  Team | Code
 *      - One short, static, season-long code per team (e.g. 0001, 0002, ...).
 *      - Anyone on the team can use it — no specific captain needs to be
 *        present.
 *
 *   2. A dual-code submission gate (validateTeamCodes): a match is only
 *      accepted if BOTH teams enter their own team code. One valid submission
 *      still produces exactly one pending row; the manager's Verified-checkbox
 *      approval is untouched.
 *
 *   3. An optional per-league "Roster" tab (Team | Player | Captain? | Code)
 *      that feeds the app's player/team dropdowns. This is just a convenience
 *      for data entry and is independent of the team-code gate.
 *
 * APPLIES TO EVERY LEAGUE. Nothing here is league-specific.
 */

// ============================================================
//  TEAM CODES TAB  (the submission gate)
// ============================================================
var TEAM_CODES_TAB = 'Team Codes';
var TEAM_CODES_HEADERS = ['Team', 'Code'];

// When false, leagues that have NOT set up team codes yet keep submitting as
// before (graceful rollout — nothing breaks mid-season). As soon as a league
// has team codes, the gate enforces for that league. Flip to true to require
// codes for EVERY league, no exceptions.
var ROSTER_GATE_STRICT = false;

// ============================================================
//  OPTIONAL PLAYER ROSTER TAB  (feeds app dropdowns only)
// ============================================================
var ROSTER_TAB = 'Roster';
var ROSTER_HEADERS = ['Team', 'Player', 'Captain?', 'Code'];

// ============================================================
//  TEAM CODE READING + VALIDATION
// ============================================================

// Reads the Team Codes tab into { teamLowerCase: normalisedCode }.
function readTeamCodes(ss) {
  var sheet = ss.getSheetByName(TEAM_CODES_TAB);
  if (!sheet || sheet.getLastRow() < 2) return {};
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, TEAM_CODES_HEADERS.length).getValues();
  var map = {};
  for (var i = 0; i < data.length; i++) {
    var team = String(data[i][0] || '').trim();
    var code = normCode(data[i][1]);
    if (team && code) map[team.toLowerCase()] = code;
  }
  return map;
}

// Validates the two team codes on an incoming submission.
//   returns { ok: true }              on success
//   returns { ok: false, message }    on rejection
function validateTeamCodes(ss, data) {
  var codes = readTeamCodes(ss);

  // No team codes configured for this league yet.
  if (!Object.keys(codes).length) {
    if (ROSTER_GATE_STRICT) {
      return { ok: false, message: 'Team codes are not set up for this league yet. Ask your league manager.' };
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
    return { ok: false, message: 'Both teams must enter their code before submitting.' };
  }

  var hExpected = codes[home.toLowerCase()];
  var aExpected = codes[away.toLowerCase()];
  if (hExpected == null) {
    return { ok: false, message: 'No team code is set for "' + home + '". Ask your manager.' };
  }
  if (aExpected == null) {
    return { ok: false, message: 'No team code is set for "' + away + '". Ask your manager.' };
  }
  if (hc !== hExpected) {
    return { ok: false, message: 'Wrong code for ' + home + '.' };
  }
  if (ac !== aExpected) {
    return { ok: false, message: 'Wrong code for ' + away + '.' };
  }
  if (hc === ac) {
    return { ok: false, message: 'Both teams entered the same code — each team must enter its own.' };
  }

  return { ok: true };
}

function normCode(v) {
  return String(v == null ? '' : v).trim().toUpperCase();
}

function teamEq(a, b) {
  return String(a || '').trim().toLowerCase() === String(b || '').trim().toLowerCase();
}

// Zero-pad a sequence number to at least 4 digits: 1 -> "0001", 23 -> "0023".
function padCode(n) {
  var s = String(n);
  while (s.length < 4) s = '0' + s;
  return s;
}

// ============================================================
//  TEAM DISCOVERY + CODE SET-UP
// ============================================================

// Gathers every distinct team name known for a league, from (in order) the
// existing Team Codes tab, the player Roster/seed, Team Standings, and the
// submitted Players tab.
function collectTeamsForLeague(ss) {
  var set = {};
  var add = function (t) { t = String(t || '').trim(); if (t) set[t] = true; };

  var existing = ss.getSheetByName(TEAM_CODES_TAB);
  if (existing && existing.getLastRow() >= 2) {
    existing.getRange(2, 1, existing.getLastRow() - 1, 1).getValues().forEach(function (r) { add(r[0]); });
  }

  // Roster tab + seed (via getRosterForLeague).
  try { getRosterForLeague(ss).teams.forEach(add); } catch (e) { /* ignore */ }

  var ts = ss.getSheetByName(TEAM_STANDINGS_TAB);
  if (ts && ts.getLastRow() >= 2) {
    ts.getRange(2, 2, ts.getLastRow() - 1, 1).getValues().forEach(function (r) { add(r[0]); });
  }

  var ps = ss.getSheetByName(PLAYERS_TAB);
  if (ps && ps.getLastRow() >= 2) {
    ps.getRange(2, 5, ps.getLastRow() - 1, 1).getValues().forEach(function (r) { add(r[0]); });
  }

  return Object.keys(set).sort();
}

// Creates/tops up the Team Codes tab for one league: every team without a code
// gets the next sequential code (0001, 0002, ...). Existing codes are never
// overwritten. Codes stay unique within the league.
//   returns { created: bool, added: number }
function setupTeamCodesForSheet(ss) {
  var sheet = ss.getSheetByName(TEAM_CODES_TAB);
  var created = false;
  if (!sheet) {
    sheet = ss.insertSheet(TEAM_CODES_TAB);
    sheet.getRange(1, 1, 1, TEAM_CODES_HEADERS.length).setValues([TEAM_CODES_HEADERS]).setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 240);
    created = true;
  } else if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, TEAM_CODES_HEADERS.length).setValues([TEAM_CODES_HEADERS]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  // Keep codes as text so leading zeros (0001) are preserved.
  sheet.getRange(2, 2, Math.max(sheet.getMaxRows() - 1, 1), 1).setNumberFormat('@');

  // Existing teams + codes already in the tab.
  var existingTeams = {};
  var maxNum = 0;
  if (sheet.getLastRow() >= 2) {
    var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, TEAM_CODES_HEADERS.length).getValues();
    rows.forEach(function (r) {
      var team = String(r[0] || '').trim();
      if (team) existingTeams[team.toLowerCase()] = true;
      var n = parseInt(normCode(r[1]), 10);
      if (!isNaN(n) && n > maxNum) maxNum = n;
    });
  }

  // Append teams that don't have a code row yet.
  var teams = collectTeamsForLeague(ss);
  var toAdd = [];
  teams.forEach(function (team) {
    if (existingTeams[team.toLowerCase()]) return;
    maxNum++;
    toAdd.push([team, padCode(maxNum)]);
  });

  if (toAdd.length) {
    var startRow = Math.max(sheet.getLastRow() + 1, 2);
    sheet.getRange(startRow, 1, toAdd.length, TEAM_CODES_HEADERS.length).setValues(toAdd);
  }

  return { created: created, added: toAdd.length };
}

// ============================================================
//  OPTIONAL PLAYER ROSTER (dropdowns only — not the gate)
// ============================================================

// Reads the Roster tab into [{ team, player, captain, code }]. [] if absent.
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
// Roster tab; falls back to the legacy seed/Players derivation.
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

// Creates/tops up the player Roster tab from existing data (optional helper).
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

  var existing = {};
  readRosterTab(ss).forEach(function (r) { existing[(r.team + '|' + r.player).toLowerCase()] = true; });

  var candidates = [];
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

  candidates.sort(function (a, b) {
    return a.team.localeCompare(b.team) || a.player.localeCompare(b.player);
  });
  var toAdd = candidates
    .filter(function (c) { return !existing[(c.team + '|' + c.player).toLowerCase()]; })
    .map(function (c) { return [c.team, c.player, false, '']; });

  if (toAdd.length) {
    var startRow = Math.max(sheet.getLastRow() + 1, 2);
    sheet.getRange(startRow, 1, toAdd.length, ROSTER_HEADERS.length).setValues(toAdd);
    sheet.getRange(startRow, 3, toAdd.length, 1).insertCheckboxes();
  }

  return { created: created, added: toAdd.length };
}

// ============================================================
//  MENU HANDLERS  (wired up from Code.gs > onOpen)
// ============================================================

function setupAllTeamCodes() {
  var leagues = getAllConfiguredLeagues();
  if (!leagues.length) {
    SpreadsheetApp.getUi().alert('No leagues configured. Fill in the Config tab first.');
    return;
  }
  var report = [];
  leagues.forEach(function (entry) {
    try {
      var ss = SpreadsheetApp.openById(entry.sheetId);
      var r = setupTeamCodesForSheet(ss);
      report.push('✓ ' + entry.league + ' — ' + (r.created ? 'created, ' : '') + r.added + ' team code(s) added');
    } catch (e) {
      report.push('✗ ' + entry.league + ' — ' + e.message);
    }
  });
  SpreadsheetApp.getUi().alert(
    'Team codes set up:\n\n' + report.join('\n') +
    '\n\nOpen each "Team Codes" tab to see/share each team\'s code. Leagues ' +
    'with no teams yet: add team names to the tab and run this again to fill ' +
    'their codes.'
  );
}

// Reset (re-issue) one team's code, e.g. if it leaks.
function resetTeamCodePrompt() {
  var ui = SpreadsheetApp.getUi();
  var lr = ui.prompt('Reset a team code', 'League name:', ui.ButtonSet.OK_CANCEL);
  if (lr.getSelectedButton() !== ui.Button.OK) return;
  var league = lr.getResponseText().trim();
  var sheetId = lookupLeagueSheetId(league);
  if (!sheetId) { ui.alert('Unknown league: ' + league); return; }

  var tr = ui.prompt('Reset a team code', 'Team name to re-issue a code for:', ui.ButtonSet.OK_CANCEL);
  if (tr.getSelectedButton() !== ui.Button.OK) return;
  var team = tr.getResponseText().trim();
  if (!team) return;

  var ss = SpreadsheetApp.openById(sheetId);
  var sheet = ss.getSheetByName(TEAM_CODES_TAB);
  if (!sheet || sheet.getLastRow() < 2) { ui.alert('No Team Codes tab for ' + league + '.'); return; }
  var range = sheet.getRange(2, 1, sheet.getLastRow() - 1, TEAM_CODES_HEADERS.length);
  var values = range.getValues();

  var maxNum = 0, used = {};
  values.forEach(function (r) {
    var n = parseInt(normCode(r[1]), 10);
    if (!isNaN(n)) { used[n] = true; if (n > maxNum) maxNum = n; }
  });

  for (var i = 0; i < values.length; i++) {
    if (teamEq(values[i][0], team)) {
      maxNum++;
      values[i][1] = padCode(maxNum);
      range.setValues(values);
      ui.alert('✓ New code for ' + values[i][0] + ': ' + values[i][1]);
      return;
    }
  }
  ui.alert('No row found for team "' + team + '" in ' + league + '.');
}

// Optional: build the player Roster tabs (for the app's dropdowns).
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
  SpreadsheetApp.getUi().alert('Player rosters set up:\n\n' + report.join('\n'));
}

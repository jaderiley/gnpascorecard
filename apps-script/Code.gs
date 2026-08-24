/**
 * GNPA Pool League — Master Apps Script (multi-sheet router)
 *
 * This script lives in the MASTER sheet. It reads a config table mapping
 * league names to per-league Spreadsheet IDs, then writes incoming match
 * submissions to the correct league sheet.
 *
 * ARCHITECTURE
 *   Master Sheet (where this script lives)
 *     └── Config tab — table of [League Name | Spreadsheet ID]
 *
 *   Per-league Sheets (one per league, owned by the manager)
 *     └── Team Codes (Team | Code — one season-long code per team; the gate)
 *     └── Roster   (Team | Player | Captain? | Code — optional, feeds dropdowns)
 *     └── Matches  (raw submissions, with Verified checkbox column)
 *     └── Frames
 *     └── Players
 *     └── Team Standings    (computed, regenerated on rebuild)
 *     └── Player Standings  (computed, regenerated on rebuild)
 *
 * APPROVAL WORKFLOW
 *   Every match row has a "Verified" checkbox. Standings only count
 *   matches where Verified = TRUE. Manager reviews submissions, edits any
 *   wrong cells in place, then ticks the box.
 *
 * SUBMISSION GATE
 *   Match submissions require BOTH teams' codes (see roster.gs >
 *   validateTeamCodes). One valid submission produces one pending row.
 *
 * SETUP
 *   1. Open the MASTER sheet → Extensions → Apps Script
 *   2. Paste this file (Code.gs), plus roster.gs and seed.gs. Save.
 *   3. Deploy → Manage deployments → pencil icon → New version → Deploy.
 *      (URL stays the same.)
 *   4. Reload the master sheet — you'll see a "GNPA League" menu appear.
 *   5. GNPA League → "Set up Config tab", then fill in Spreadsheet IDs.
 *   6. GNPA League → "Initialize all league sheets".
 *   7. GNPA League → "Set up team codes (all leagues)". Open each
 *      "Team Codes" tab to see/share each team's code.
 */

// ============================================================
//  CONFIG SHEET (lives in the master spreadsheet)
// ============================================================
var CONFIG_SHEET_NAME = 'Config';
var CONFIG_HEADERS = ['League', 'Spreadsheet ID', 'Notes'];

// Default leagues — used when initializing the Config tab.
// You can edit/extend the Config tab manually after setup.
var DEFAULT_LEAGUES = [
  'Vets Tier 1',
  'Vets Tier 2',
  'Super North',
  'Super South',
  'Premier',
  'Ladies',
  'Tshwane',
  '3-Man',
  'Juniors'
];

// ============================================================
//  TAB NAMES inside each per-league sheet
// ============================================================
var MATCHES_TAB  = 'Matches';
var FRAMES_TAB   = 'Frames';
var PLAYERS_TAB  = 'Players';
var TEAM_STANDINGS_TAB   = 'Team Standings';
var PLAYER_STANDINGS_TAB = 'Player Standings';

// The one and only per-league branch in the system. For this league, the
// player percentage is the standard own-frames win rate (frames won ÷ frames
// the player personally played). For every OTHER league it is frames won ÷
// the team's total frames for the season. See rebuildPlayerStandings.
var OWN_FRAMES_PCT_LEAGUE = 'Ladies';

// Player standings: a player must have played at least this fraction of the
// league's maximum framesPlayed to be ranked. Anyone below appears under a
// "LESS THAN ...% LEAGUE GAMES PLAYED" heading, unranked.
var PLAYER_QUALIFY_PCT = 0.60;
// Per-league overrides of the qualify threshold. Ladies requires 85% of the
// league's max games played to be ranked (Jade, 2026-07-18). Keep in sync with
// PLAYER_QUALIFY_PCT in gnpa_cli.py.
var PLAYER_QUALIFY_PCT_OVERRIDE = { 'Ladies': 0.85 };
function qualifyPctForLeague_(leagueName) {
  var key = String(leagueName).trim();
  return PLAYER_QUALIFY_PCT_OVERRIDE.hasOwnProperty(key)
    ? PLAYER_QUALIFY_PCT_OVERRIDE[key] : PLAYER_QUALIFY_PCT;
}
// ============================================================
//  SCHEMAS
// ============================================================
var MATCHES_HEADERS = [
  'Submitted', 'Date', 'League', 'Format', 'Home Team', 'Away Team',
  'Home Captain', 'Away Captain', 'Home Score', 'Away Score',
  'Frames Played', 'Result', 'Verified'
];

var FRAMES_HEADERS = [
  'Submitted', 'Date', 'League', 'Type', 'Home Team', 'Away Team',
  'Home Pos', 'Away Pos', 'Home Player', 'Away Player',
  'Home Points', 'Away Points', 'Winner'
];

var PLAYERS_HEADERS = [
  'Submitted', 'Date', 'League', 'Side', 'Team',
  'Position', 'Player', 'Frames Won', 'Frames Played'
];

// ============================================================
//  ENTRY POINTS (HTTP)
// ============================================================

function doPost(e) {
  try {
    var body = e.postData && e.postData.contents ? e.postData.contents : '{}';
    var data = JSON.parse(body);

    if (data.ping) {
      return jsonResponse({ ok: true, message: 'Endpoint is alive.' });
    }

    // Secret-gated Sheets bridge (bridge.gs) — isolated from match submission.
    if (data.action === 'bridge') {
      return jsonResponse(handleBridge(data));
    }

    var result = handleSubmission(data);
    return jsonResponse(result);
  } catch (err) {
    return jsonResponse({ ok: false, message: 'Error: ' + err.message });
  }
}

function doGet(e) {
  // ?action=knownNames&league=Ladies — used by app's typo-detection.
  // ?action=standings&league=Ladies — returns team + player standings as JSON.
  // ?action=leagues — returns the list of configured leagues.
  try {
    var params = (e && e.parameter) || {};
    if (e && e.parameter && e.parameter.action === 'roster') {
      return handleRosterRequest(e);
    }
    if (params.action === 'knownNames') {
      var league = params.league;
      if (!league) return jsonResponse({ ok: false, message: 'league parameter required' });
      var knCache = CacheService.getScriptCache();
      var knKey = 'knownNames:' + league + ':' + (params.team || '');
      var knHit = knCache.get(knKey);
      if (knHit) {
        return ContentService.createTextOutput(knHit).setMimeType(ContentService.MimeType.JSON);
      }
      var sheetId = lookupLeagueSheetId(league);
      if (!sheetId) return jsonResponse({ ok: false, message: 'Unknown league: ' + league });
      var result = getKnownNames(sheetId, params.team || null);
      var knPayload = JSON.stringify({ ok: true, names: result.names, teams: result.teams });
      try { knCache.put(knKey, knPayload, 1200); } catch (cacheErr) { /* ignore */ }
      return ContentService.createTextOutput(knPayload).setMimeType(ContentService.MimeType.JSON);
    }
    if (params.action === 'standings') {
      var league = params.league;
      if (!league) return jsonResponse({ ok: false, message: 'league parameter required' });
      // Try cache first
      var cache = CacheService.getScriptCache();
      var cacheKey = 'standings:' + league;
      var cached = cache.get(cacheKey);
      if (cached) {
        return ContentService.createTextOutput(cached).setMimeType(ContentService.MimeType.JSON);
      }
      var sheetId = lookupLeagueSheetId(league);
      if (!sheetId) return jsonResponse({ ok: false, message: 'Unknown league: ' + league });
      var standings = getStandings(sheetId);
      var payload = JSON.stringify({ ok: true, league: league, teams: standings.teams, players: standings.players });
      try { cache.put(cacheKey, payload, 1200); } catch (cacheErr) { /* ignore */ }
      return ContentService.createTextOutput(payload).setMimeType(ContentService.MimeType.JSON);
    }
    if (params.action === 'leagues') {
      var cache = CacheService.getScriptCache();
      var cached = cache.get('leagues');
      if (cached) {
        return ContentService.createTextOutput(cached).setMimeType(ContentService.MimeType.JSON);
      }
      var leagues = getAllConfiguredLeagues().map(function(l){ return l.league; });
      var payload = JSON.stringify({ ok: true, leagues: leagues });
      try { cache.put('leagues', payload, 1200); } catch (cacheErr) { /* ignore */ }
      return ContentService.createTextOutput(payload).setMimeType(ContentService.MimeType.JSON);
    }
  } catch (err) {
    return jsonResponse({ ok: false, message: 'Error: ' + err.message });
  }
  return jsonResponse({ ok: true, message: 'GNPA endpoint live. POST match data here.' });
}
// ============================================================
//  ROSTER HANDLER
// ============================================================
function handleRosterRequest(e) {
  var resp = function (obj) {
    return ContentService.createTextOutput(JSON.stringify(obj))
      .setMimeType(ContentService.MimeType.JSON);
  };
  var league = e && e.parameter && e.parameter.league;
  if (!league) return resp({ ok: false, message: 'Missing league' });

  // Cache first. Measured 2026-08-24: uncached this call cost 4-6s EVERY time
  // (it reopens the spreadsheet and rebuilds the Roster + seed + Players union
  // from scratch), against a ~2.0s floor for any Apps Script request. That
  // 4-6s window is also when a phone locking or dropping wifi kills the fetch,
  // which used to strand the app on "Loading teams..." forever.
  var cache = CacheService.getScriptCache();
  var cacheKey = 'roster:' + league;
  try {
    var hit = cache.get(cacheKey);
    if (hit) return ContentService.createTextOutput(hit).setMimeType(ContentService.MimeType.JSON);
  } catch (cacheErr) { /* fall through and compute it */ }

  var sheetId = lookupLeagueSheetId(league);
  if (!sheetId) return resp({ ok: false, message: 'Unknown league: ' + league });

  var ss = SpreadsheetApp.openById(sheetId);
  // Prefer the authoritative Roster tab (season-start roster + captains);
  // getRosterForLeague (roster.gs) falls back to seed/Players data if the
  // Roster tab hasn't been built yet.
  var roster = getRosterForLeague(ss);

  var payload = JSON.stringify({
    ok: true,
    teams: roster.teams,
    rosters: roster.rosters,
    captains: roster.captains
  });
  // Only ever cache a good answer — never cache an error or an empty roster,
  // or one bad read would be served to every captain for 20 minutes.
  if (roster.teams && roster.teams.length) {
    try { cache.put(cacheKey, payload, 1200); } catch (cacheErr) { /* ignore */ }
  }
  return ContentService.createTextOutput(payload).setMimeType(ContentService.MimeType.JSON);
}
// ============================================================
//  SUBMISSION HANDLER
// ============================================================

function handleSubmission(data) {
  if (!data.date || !data.homeTeam || !data.awayTeam) {
    return { ok: false, message: 'Missing required fields (date, homeTeam, awayTeam)' };
  }
  if (!data.league) {
    return { ok: false, message: 'League is required' };
  }

  var sheetId = lookupLeagueSheetId(data.league);
  if (!sheetId) {
    return {
      ok: false,
      message: 'League "' + data.league + '" is not configured. Add it to the Config tab in the master sheet.'
    };
  }

  var leagueSS;
  try {
    leagueSS = SpreadsheetApp.openById(sheetId);
  } catch (err) {
    return {
      ok: false,
      message: 'Could not open the sheet for "' + data.league + '". Check the Spreadsheet ID in Config.'
    };
  }

  // === Dual team-code gate (roster.gs). Both teams must enter their own
  //     team code. Rejected submissions never reach the sheet.
  var gate = validateTeamCodes(leagueSS, data);
  if (!gate.ok) {
    return { ok: false, message: gate.message };
  }

  // === Duplicate-submission guard (added 2026-07-10).
  // Real-world failure seen in the Ladies league on 2026-07-09: the client's
  // fetch can die AFTER the server has already written all the rows (flaky
  // mobile signal, Apps Script redirect quirks). The app then shows "Failed",
  // re-enables the Submit button, the captain taps again, and the entire
  // match is double-written (4 of 8 matches that night landed twice).
  // An identical resubmission (same date/teams/scores/framesPlayed) is always
  // a retry — swallow it and report success so the client stops retrying.
  // Extended 2026-07-18: a retap where the captain also CHANGED THE DATE (the
  // 2026-07-16 OBS Mafia vs Bitch Squad 7/14-vs-7/15 double-write) is caught
  // too, because such a resubmission lands within minutes of the original (see
  // RETAP_WINDOW_MS in findDuplicateMatch_). A resubmission with DIFFERENT
  // scores, or the same fixture genuinely replayed weeks later, still lands;
  // the manager verifies the right row.
  var lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch (lockErr) { /* proceed — the sheet scan below still guards */ }
  if (findDuplicateMatch_(leagueSS, data)) {
    try { lock.releaseLock(); } catch (e) {}
    return {
      ok: true,
      duplicate: true,
      message: 'This match is already logged in ' + data.league + ' — duplicate submission ignored.'
    };
  }

  var now = new Date();
  var timestamp = Utilities.formatDate(now, leagueSS.getSpreadsheetTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  var format = data.format || 'rotation';

  var result;
  if (data.homeScore > data.awayScore) result = 'Home win';
  else if (data.awayScore > data.homeScore) result = 'Away win';
  else result = 'Draw';

  // === Matches tab — append one row, with Verified=FALSE checkbox.
  // We use explicit row positioning (not appendRow) because pre-existing
  // data validation in the Verified column can confuse appendRow's
  // "find the first empty row" logic.
  var matchesSheet = getOrCreateSheet(leagueSS, MATCHES_TAB, MATCHES_HEADERS);
  var newRow = matchesSheet.getLastRow() + 1;
  matchesSheet.getRange(newRow, 1, 1, MATCHES_HEADERS.length).setValues([[
    timestamp, data.date, data.league, format, data.homeTeam, data.awayTeam,
    data.homeCaptain || '', data.awayCaptain || '',
    data.homeScore, data.awayScore, data.framesPlayed, result, false
  ]]);
  var verifiedCol = MATCHES_HEADERS.length;
  matchesSheet.getRange(newRow, verifiedCol).insertCheckboxes();

  // === Frames tab
  if (data.frames && data.frames.length > 0) {
    var framesSheet = getOrCreateSheet(leagueSS, FRAMES_TAB, FRAMES_HEADERS);
    var frameRows = data.frames.map(function (f) {
      return [
        timestamp, data.date, data.league,
        f.type || 'singles',
        data.homeTeam, data.awayTeam,
        f.homePos, f.awayPos,
        f.homePlayer || '', f.awayPlayer || '',
        f.homePoints != null ? f.homePoints : '',
        f.awayPoints != null ? f.awayPoints : '',
        f.winner === 'home' ? 'Home' : 'Away'
      ];
    });
    framesSheet.getRange(framesSheet.getLastRow() + 1, 1, frameRows.length, FRAMES_HEADERS.length).setValues(frameRows);
  }

  // === Players tab
  if (data.players && data.players.length > 0) {
    var playersSheet = getOrCreateSheet(leagueSS, PLAYERS_TAB, PLAYERS_HEADERS);
    var playerRows = data.players.map(function (p) {
      return [
        timestamp, data.date, data.league,
        p.side === 'home' ? 'Home' : 'Away', p.team,
        p.position, p.name, p.framesWon, p.framesPlayed
      ];
    });
    playersSheet.getRange(playersSheet.getLastRow() + 1, 1, playerRows.length, PLAYERS_HEADERS.length).setValues(playerRows);
  }

  try { lock.releaseLock(); } catch (e) {}
  // A player making their debut enters the roster union via the Players tab,
  // so this submission just changed what ?action=roster would return.
  invalidateLeagueCaches_(data.league);
  return {
    ok: true,
    message: 'Match logged in ' + data.league + ' (pending verification)'
  };
}

// A retap resubmission lands within seconds/minutes of the original write, so
// if the same fixture+scores was submitted this recently we treat it as a
// duplicate even when the captain changed the match Date. A genuine rematch of
// the same two teams is weeks apart — well outside this window — so it still
// goes through.
var RETAP_WINDOW_MS = 15 * 60 * 1000;

// Returns true if the Matches tab already holds a row for the same
// teams + scores + framesPlayed AND either (a) the same match date, or (b) it
// was submitted within RETAP_WINDOW_MS of now (a same-session retap that also
// changed the date). Serialized by the script lock in handleSubmission, so two
// concurrent retries can't both pass the check.
function findDuplicateMatch_(leagueSS, data) {
  var sheet = leagueSS.getSheetByName(MATCHES_TAB);
  if (!sheet || sheet.getLastRow() < 2) return false;
  var tz = leagueSS.getSpreadsheetTimeZone();
  var nowMs = Date.now();
  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, MATCHES_HEADERS.length).getValues();
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    var sameFixture =
        String(r[4]).trim() === String(data.homeTeam).trim() &&
        String(r[5]).trim() === String(data.awayTeam).trim() &&
        Number(r[8]) === Number(data.homeScore) &&
        Number(r[9]) === Number(data.awayScore) &&
        Number(r[10]) === Number(data.framesPlayed);
    if (!sameFixture) continue;

    // (a) same match date → definite resubmission.
    var rowDate = r[1] instanceof Date ? Utilities.formatDate(r[1], tz, 'yyyy-MM-dd') : String(r[1]).trim();
    if (rowDate === String(data.date).trim()) return true;

    // (b) different date but submitted moments ago → retap that changed the date.
    var submittedMs = parseSubmittedMs_(r[0]);
    if (submittedMs != null) {
      var age = nowMs - submittedMs;
      if (age >= 0 && age <= RETAP_WINDOW_MS) return true;
    }
  }
  return false;
}

// Parse a Matches "Submitted" cell (a Date object, or a 'yyyy-MM-dd HH:mm:ss'
// string in the sheet's timezone) to epoch ms. Returns null if unparseable —
// callers then fall back to the same-date check only (never a false positive).
function parseSubmittedMs_(v) {
  if (v instanceof Date) return v.getTime();
  var s = String(v == null ? '' : v).trim();
  if (!s) return null;
  var t = new Date(s.replace(' ', 'T')).getTime();
  return isNaN(t) ? null : t;
}

// ============================================================
//  CONFIG LOOKUP
// ============================================================

function lookupLeagueSheetId(leagueName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return null;

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  for (var i = 0; i < data.length; i++) {
    var name = String(data[i][0] || '').trim();
    var id   = String(data[i][1] || '').trim();
    if (name === leagueName && id) return id;
  }
  return null;
}

function getAllConfiguredLeagues() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return [];

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  var out = [];
  for (var i = 0; i < data.length; i++) {
    var name = String(data[i][0] || '').trim();
    var id   = String(data[i][1] || '').trim();
    if (name && id) out.push({ league: name, sheetId: id });
  }
  return out;
}

// ============================================================
//  KNOWN NAMES (for app's typo-detection)
// ============================================================

function getKnownNames(sheetId, teamFilter) {
  var ss;
  try { ss = SpreadsheetApp.openById(sheetId); } catch (e) { return { names: [], teams: [] }; }
  var seen = {};
  var out = [];
  var teamsSeen = {};
  var teamsOut = [];

  // Source 1: raw Players tab (live app-submitted matches)
  var sheet = ss.getSheetByName(PLAYERS_TAB);
  if (sheet && sheet.getLastRow() >= 2) {
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, PLAYERS_HEADERS.length).getValues();
    for (var i = 0; i < data.length; i++) {
      var team = String(data[i][4] || '');
      var name = String(data[i][6] || '').trim();
      if (!name) continue;
      if (teamFilter && team !== teamFilter) continue;
      var key = name.toLowerCase();
      if (!seen[key]) { seen[key] = true; out.push(name); }
      if (team && !teamsSeen[team.toLowerCase()]) { teamsSeen[team.toLowerCase()] = true; teamsOut.push(team); }
    }
  }

  // Source 2: Player Standings tab (includes seeded historical data)
  var psSheet = ss.getSheetByName(PLAYER_STANDINGS_TAB);
  if (psSheet && psSheet.getLastRow() >= 2) {
    var psData = psSheet.getRange(2, 1, psSheet.getLastRow() - 1, 6).getValues();
    // Cols: Pos, Team, Player, F/W, F/P, %
    for (var i = 0; i < psData.length; i++) {
      var team = String(psData[i][1] || '');
      var name = String(psData[i][2] || '').trim();
      if (!name) continue;
      if (teamFilter && team !== teamFilter) continue;
      var key = name.toLowerCase();
      if (!seen[key]) { seen[key] = true; out.push(name); }
      if (team && !teamsSeen[team.toLowerCase()]) { teamsSeen[team.toLowerCase()] = true; teamsOut.push(team); }
    }
  }

  // Source 3: Team Standings tab (for team name typo detection)
  var tsSheet = ss.getSheetByName(TEAM_STANDINGS_TAB);
  if (tsSheet && tsSheet.getLastRow() >= 2) {
    var tsData = tsSheet.getRange(2, 1, tsSheet.getLastRow() - 1, 2).getValues();
    for (var i = 0; i < tsData.length; i++) {
      var team = String(tsData[i][1] || '').trim();
      if (team && !teamsSeen[team.toLowerCase()]) { teamsSeen[team.toLowerCase()] = true; teamsOut.push(team); }
    }
  }

  return { names: out, teams: teamsOut };
}

// Read the pre-computed Team Standings and Player Standings tabs from a
// league sheet and return them as JSON-friendly arrays. Returns empty arrays
// if standings haven't been built yet.
function getStandings(sheetId) {
  var ss;
  try { ss = SpreadsheetApp.openById(sheetId); } catch (e) { return { teams: [], players: [] }; }

  var teams = [];
  var teamSheet = ss.getSheetByName(TEAM_STANDINGS_TAB);
  if (teamSheet && teamSheet.getLastRow() > 1) {
    var teamData = teamSheet.getRange(2, 1, teamSheet.getLastRow() - 1, 10).getValues();
    // Cols: Pos, Team, W, L, D, Played, Frames Played, Frames Won, % Won, Points
    teams = teamData.filter(function(r){ return r[1]; }).map(function(r){
      return {
        pos: Number(r[0]) || 0,
        team: String(r[1] || ''),
        w: Number(r[2]) || 0,
        l: Number(r[3]) || 0,
        d: Number(r[4]) || 0,
        played: Number(r[5]) || 0,
        framesPlayed: Number(r[6]) || 0,
        framesWon: Number(r[7]) || 0,
        pct: Number(r[8]) || 0,
        points: Number(r[9]) || 0
      };
    });
  }

  var players = [];
  var playerSheet = ss.getSheetByName(PLAYER_STANDINGS_TAB);
  if (playerSheet && playerSheet.getLastRow() > 1) {
    // Cols: Pos, Team, Player, F/W, F/P, %
    var playerData = playerSheet.getRange(2, 1, playerSheet.getLastRow() - 1, 6).getValues();
    players = playerData.filter(function(r){ return r[2]; }).map(function(r){
      return {
        pos: Number(r[0]) || 0,
        team: String(r[1] || ''),
        player: String(r[2] || ''),
        framesWon: Number(r[3]) || 0,
        framesPlayed: Number(r[4]) || 0,
        pct: Number(r[5]) || 0
      };
    });
  }

  return { teams: teams, players: players };
}

// ============================================================
//  STANDINGS REBUILD
// ============================================================

function rebuildStandingsForLeague(leagueName) {
  var sheetId = lookupLeagueSheetId(leagueName);
  if (!sheetId) throw new Error('League not configured: ' + leagueName);
  var ss = SpreadsheetApp.openById(sheetId);
  rebuildTeamStandings(ss, leagueName);
  rebuildPlayerStandings(ss, leagueName);
  invalidateLeagueCaches_(leagueName);
}

/**
 * Drop every cached GET response for one league.
 *
 * Call this from ANY code path that changes what the app would be told —
 * a rebuild, a submission (a debutant joins the roster union via the Players
 * tab), a manager Add/Rename, or the mobile Refresh checkbox. A stale team
 * picker is worse than a slow one, so when in doubt, call it.
 *
 * What it CANNOT catch: a manager editing the Roster or Team Codes tab by
 * hand. That is what the 20-minute TTL is for — do not raise it much beyond
 * that, and point managers at the Refresh checkbox as the manual escape hatch.
 */
function invalidateLeagueCaches_(leagueName) {
  if (!leagueName) return;
  try {
    var cache = CacheService.getScriptCache();
    cache.remove('standings:' + leagueName);
    cache.remove('roster:' + leagueName);
    // knownNames is keyed per-team as well; clear the league-wide entry and
    // let any per-team entries age out on their own (they are read-only
    // typo-detection hints, never used to gate a submission).
    cache.remove('knownNames:' + leagueName + ':');
  } catch (e) { /* cache is best-effort — never break a write over it */ }
}

function rebuildTeamStandings(ss, leagueName) {
  // 1. Read historical seed (if any)
  var seed = readTeamSeed(ss);

  // 2. Read live match data
  var matchesSheet = ss.getSheetByName(MATCHES_TAB);
  var teams = {};

  if (matchesSheet && matchesSheet.getLastRow() >= 2) {
    var data = matchesSheet.getRange(2, 1, matchesSheet.getLastRow() - 1, MATCHES_HEADERS.length).getValues();
    data.forEach(function (row) {
      if (row[12] !== true) return;
      var home = row[4], away = row[5];
      var hScore = Number(row[8]) || 0, aScore = Number(row[9]) || 0;
      var totalUnitsInMatch = hScore + aScore;

      [
        { name: home, my: hScore, opp: aScore },
        { name: away, my: aScore, opp: hScore }
      ].forEach(function (t) {
        if (!t.name) return;
        if (!teams[t.name]) {
          teams[t.name] = { won: 0, lost: 0, draw: 0, played: 0, unitsWon: 0, unitsTotal: 0 };
        }
        var team = teams[t.name];
        team.played++;
        team.unitsWon   += t.my;
        team.unitsTotal += totalUnitsInMatch;
        if (t.my > t.opp) team.won++;
        else if (t.my < t.opp) team.lost++;
        else team.draw++;
      });
    });
  }

  // 3. Merge seed into live data
  Object.keys(seed).forEach(function (name) {
    var s = seed[name];
    if (!teams[name]) {
      teams[name] = { won: 0, lost: 0, draw: 0, played: 0, unitsWon: 0, unitsTotal: 0 };
    }
    var t = teams[name];
    t.won      += s.won;
    t.lost     += s.lost;
    t.draw     += s.draw;
    t.played   += s.played;
    t.unitsWon += s.unitsWon;
    t.unitsTotal += s.unitsTotal;
  });

  // 4. Build rows, sort, write
  var rows = Object.keys(teams).map(function (name) {
    var t = teams[name];
    var pct = t.unitsTotal > 0 ? (t.unitsWon / t.unitsTotal * 100) : 0;
    var points = t.won * 2 + t.draw;
    return [name, t.won, t.lost, t.draw, t.played, t.unitsTotal, t.unitsWon, pct, points];
  });
  rows.sort(function (a, b) { return b[8] - a[8] || b[7] - a[7]; });
  rows = rows.map(function (r, i) { return [i + 1].concat(r); });

  writeStandingsTab(ss, TEAM_STANDINGS_TAB,
    ['Pos', 'Team', 'W', 'L', 'D', 'Played', 'Frames Played', 'Frames Won', '% Won', 'Points'],
    rows, /*pctCol=*/ 9);
}

function rebuildPlayerStandings(ss, leagueName) {
  // 1. Read historical seed (if any)
  var seed = readPlayerSeed(ss);

  // 2. Read live match data
  var matchesSheet = ss.getSheetByName(MATCHES_TAB);
  var playersSheet = ss.getSheetByName(PLAYERS_TAB);
  var players = {};

  // Per-team season frame total (the denominator for non-Ladies leagues).
  // For each team and each verified match, the "frames per position" is the
  // number of frames a single player slot plays that night (all present
  // players share it; a sub who played fewer doesn't lower it). Summing those
  // across the team's matches gives the season total a fully-attending player
  // would have played — e.g. 66 or 72 — and it grows as the season goes on.
  var liveTeamMatchFrames = {}; // team -> { matchKey -> framesPerPosition }

  if (matchesSheet && playersSheet && matchesSheet.getLastRow() >= 2 && playersSheet.getLastRow() >= 2) {
    var verifiedKeys = {};
    var matchesData = matchesSheet.getRange(2, 1, matchesSheet.getLastRow() - 1, MATCHES_HEADERS.length).getValues();
    matchesData.forEach(function (row) {
      if (row[12] === true) {
        verifiedKeys[formatTimestamp(row[0])] = true;
      }
    });

    var playersData = playersSheet.getRange(2, 1, playersSheet.getLastRow() - 1, PLAYERS_HEADERS.length).getValues();
    playersData.forEach(function (row) {
      var matchKey = formatTimestamp(row[0]);
      if (!verifiedKeys[matchKey]) return;
      var team = row[4], name = String(row[6] || '').trim();
      if (!name) return;
      var fw = Number(row[7]) || 0, fp = Number(row[8]) || 0;
      var k = team + '|' + name;
      if (!players[k]) {
        players[k] = { team: team, name: name, framesWon: 0, framesPlayed: 0 };
      }
      players[k].framesWon   += fw;
      players[k].framesPlayed += fp;

      // Track frames-per-position for this team in this match (max present).
      if (!liveTeamMatchFrames[team]) liveTeamMatchFrames[team] = {};
      if (fp > (liveTeamMatchFrames[team][matchKey] || 0)) {
        liveTeamMatchFrames[team][matchKey] = fp;
      }
    });
  }

  // 3. Merge seed into live data
  Object.keys(seed).forEach(function (k) {
    var s = seed[k];
    if (!players[k]) {
      players[k] = { team: s.team, name: s.name, framesWon: 0, framesPlayed: 0 };
    }
    players[k].framesWon    += s.framesWon;
    players[k].framesPlayed += s.framesPlayed;
  });

  // 3b. Team season frame totals = live (summed per match) + historical seed.
  //     The seed stores per-player totals, so the historical per-position
  //     total for a team is the max framesPlayed among its seeded players.
  var teamFrames = {};
  Object.keys(liveTeamMatchFrames).forEach(function (team) {
    var byMatch = liveTeamMatchFrames[team];
    var sum = 0;
    Object.keys(byMatch).forEach(function (mk) { sum += byMatch[mk]; });
    teamFrames[team] = (teamFrames[team] || 0) + sum;
  });
  var seedTeamFrames = {};
  Object.keys(seed).forEach(function (k) {
    var s = seed[k];
    if (s.framesPlayed > (seedTeamFrames[s.team] || 0)) seedTeamFrames[s.team] = s.framesPlayed;
  });
  Object.keys(seedTeamFrames).forEach(function (team) {
    teamFrames[team] = (teamFrames[team] || 0) + seedTeamFrames[team];
  });

  // 4. Work out the qualifying threshold from the league's max framesPlayed.
  //    Players >= threshold get a ranked spot; the rest go below a heading row.
  //    (This is about games PLAYED, independent of the percentage formula.)
  var maxFP = 0;
  Object.keys(players).forEach(function (k) {
    if (players[k].framesPlayed > maxFP) maxFP = players[k].framesPlayed;
  });
  var qualifyPct = qualifyPctForLeague_(leagueName);
  var threshold = maxFP * qualifyPct;

  // The single per-league branch: Ladies ranks on own frames played; every
  // other league ranks on the team's total frames for the season.
  var useOwnFrames = (String(leagueName).trim() === OWN_FRAMES_PCT_LEAGUE);

  var qualified = [];
  var unqualified = [];
  Object.keys(players).forEach(function (k) {
    var p = players[k];
    if (p.framesPlayed <= 0) return; // skip ghosts (seeded with 0/0)
    var denom = useOwnFrames ? p.framesPlayed : (teamFrames[p.team] || p.framesPlayed);
    var pct = denom > 0 ? (p.framesWon / denom) * 100 : 0;
    var row = [p.team, p.name, p.framesWon, p.framesPlayed, pct];
    if (p.framesPlayed >= threshold) qualified.push(row);
    else unqualified.push(row);
  });

  // Sort both groups by % desc
  qualified.sort(function (a, b) { return b[4] - a[4]; });
  unqualified.sort(function (a, b) { return b[4] - a[4]; });

  // 5. Stitch output: ranked qualified players, then optional heading + unranked rest.
  var rows = qualified.map(function (r, i) { return [i + 1].concat(r); });

  var separatorSheetRow = -1; // sheet-row index of the heading row, for formatting later
  if (unqualified.length) {
    separatorSheetRow = rows.length + 2; // +1 for header row, +1 because rows.length is current count before push
    rows.push(['', '', 'LESS THAN ' + qualifyPct * 100 +'% LEAGUE GAMES PLAYED', '', '', '']);
    unqualified.forEach(function (r) { rows.push([''].concat(r)); });
  }

  writeStandingsTab(ss, PLAYER_STANDINGS_TAB,
    ['Pos', 'Team', 'Player', 'F/W', 'F/P', '%'],
    rows, /*pctCol=*/ 6);

  // Style the separator row so it stands out.
  if (separatorSheetRow > 0) {
    var sheet = ss.getSheetByName(PLAYER_STANDINGS_TAB);
    sheet.getRange(separatorSheetRow, 1, 1, 6)
      .setFontWeight('bold')
      .setBackground('#f0f0f0')
      .setHorizontalAlignment('center');
  }
}


function writeStandingsTab(ss, tabName, headers, rows, pctCol) {
  var sheet = ss.getSheetByName(tabName);
  if (!sheet) sheet = ss.insertSheet(tabName);
  sheet.clear();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  sheet.setFrozenRows(1);
  if (rows.length) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    if (pctCol) {
      sheet.getRange(2, pctCol, rows.length, 1).setNumberFormat('0.00');
    }
  }
}

function formatTimestamp(ts) {
  if (ts instanceof Date) {
    return Utilities.formatDate(ts, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  }
  return String(ts);
}

// ============================================================
//  HELPERS
// ============================================================

function getOrCreateSheet(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  } else if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
//  MENU + ADMIN ACTIONS
// ============================================================

function onOpen() {
  SpreadsheetApp.getUi().createMenu('GNPA League')
    .addItem('Set up Config tab', 'setupConfigTab')
    .addItem('Initialize all league sheets', 'initializeAllLeagueSheets')
    .addSeparator()
    .addItem('Set up team codes (all leagues)', 'setupAllTeamCodes')
    .addItem('Reset one team code', 'resetTeamCodePrompt')
    .addItem('Set up player rosters (optional)', 'setupAllRosters')
    .addSeparator()
    .addItem('Rebuild standings — pick league', 'rebuildStandingsPrompt')
    .addItem('Rebuild standings — ALL leagues', 'rebuildAllStandings')
    .addSeparator()
    .addItem('Set up mobile refresh checkboxes', 'setupRebuildCheckboxes')
    .addItem('Set up manager tools tabs', 'setupManagerTabs')
    .addSeparator()
    .addItem('Compact blank rows (all leagues)', 'compactAllLeagues')
    .addToUi();
}

function setupConfigTab() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(CONFIG_SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, CONFIG_HEADERS.length).setValues([CONFIG_HEADERS]).setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(2, 380);

    var rows = DEFAULT_LEAGUES.map(function (name) { return [name, '', '']; });
    sheet.getRange(2, 1, rows.length, CONFIG_HEADERS.length).setValues(rows);
  }
  SpreadsheetApp.getUi().alert(
    'Config tab ready.\n\n' +
    'For each league, paste the Spreadsheet ID of its dedicated Google Sheet ' +
    'into column B. Get the ID from the URL: ' +
    'docs.google.com/spreadsheets/d/{ID}/edit'
  );
}

function initializeAllLeagueSheets() {
  var leagues = getAllConfiguredLeagues();
  if (!leagues.length) {
    SpreadsheetApp.getUi().alert('No leagues configured yet. Fill in the Config tab first.');
    return;
  }
  var report = [];
  leagues.forEach(function (entry) {
    try {
      var ss = SpreadsheetApp.openById(entry.sheetId);
      getOrCreateSheet(ss, MATCHES_TAB, MATCHES_HEADERS);
      getOrCreateSheet(ss, FRAMES_TAB, FRAMES_HEADERS);
      getOrCreateSheet(ss, PLAYERS_TAB, PLAYERS_HEADERS);
      report.push('✓ ' + entry.league);
    } catch (e) {
      report.push('✗ ' + entry.league + ' — ' + e.message);
    }
  });
  SpreadsheetApp.getUi().alert('Initialization done:\n\n' + report.join('\n'));
}

function rebuildStandingsPrompt() {
  var leagues = getAllConfiguredLeagues();
  if (!leagues.length) {
    SpreadsheetApp.getUi().alert('No leagues configured.');
    return;
  }
  var ui = SpreadsheetApp.getUi();
  var list = leagues.map(function (l, i) { return (i + 1) + '. ' + l.league; }).join('\n');
  var result = ui.prompt(
    'Rebuild Standings',
    'Type the league name exactly:\n\n' + list,
    ui.ButtonSet.OK_CANCEL
  );
  if (result.getSelectedButton() !== ui.Button.OK) return;
  var name = result.getResponseText().trim();
  if (!name) return;
  try {
    rebuildStandingsForLeague(name);
    ui.alert('✓ Standings rebuilt for ' + name);
  } catch (e) {
    ui.alert('Failed: ' + e.message);
  }
}

function rebuildAllStandings() {
  var leagues = getAllConfiguredLeagues();
  if (!leagues.length) {
    SpreadsheetApp.getUi().alert('No leagues configured.');
    return;
  }
  var report = [];
  leagues.forEach(function (entry) {
    try {
      rebuildStandingsForLeague(entry.league);
      report.push('✓ ' + entry.league);
    } catch (e) {
      report.push('✗ ' + entry.league + ' — ' + e.message);
    }
  });
  SpreadsheetApp.getUi().alert('Standings rebuild done:\n\n' + report.join('\n'));
}

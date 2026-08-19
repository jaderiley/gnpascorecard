/**
 * GNPA one-off DATA CLEANUP - normalises team & player names so the app's
 * dropdowns and the rebuilt standings stop showing the same team/person twice.
 *
 * WHAT IT TOUCHES (and nothing else):
 *   - Matches tab:  Home/Away Team, Home/Away Captain
 *   - Frames tab:   Home/Away Team, Home/Away Player
 *   - Players tab:  Team, Player
 *   - Historical Seed tab:  col A (Team), col B (Player)
 *   - Team Codes tab:  Team column (maps typos, then removes duplicate rows)
 * It NEVER changes scores, dates, results, or the Verified flags.
 *
 * HOW IT NORMALISES every name (generic, automatic):
 *   - collapses runs of whitespace to one space + trims
 *   - title-cases words but keeps ALL-CAPS initials (JP, SP) and lower-cases
 *     the name particles the app already lower-cases (van, de, du, le, ...)
 *   - strips accents (e-acute -> e, etc.) so accent-inconsistent spellings merge
 *   Then it applies the per-league EXPLICIT maps below for true letter-typos
 *   (Crange->Grange, Touk->Taak, V.Taak->van Taak, etc.) that no generic rule
 *   could safely guess.
 *
 * SAFETY: it's idempotent (running twice changes nothing the 2nd time) and it
 * only ever rewrites the columns listed above. Review the maps, then run
 * cleanupSuperNorth() from the Apps Script editor. It calls
 * rebuildStandingsForLeague() at the end (defined in Code.gs).
 *
 * NOT auto-fixed (need your call - flagged, left untouched):
 *   - Super North has two STRAY Super-South teams in its data:
 *     "Crucible CueTastrophe" and "OB Chalked Up". These look like matches
 *     submitted to the wrong league. Decide whether to delete those rows or
 *     re-file them, then re-run.
 */

// Particles the app's titleCase() lower-cases. Keep in sync with index.html.
var CLEAN_PARTICLES = ['van','de','vd','du','von','le','la','di','el','den','der','des'];

function stripAccents_(s) {
  // NFD-decompose, then drop combining marks U+0300..U+036F by char code.
  var d = s.normalize('NFD'), out = '';
  for (var i = 0; i < d.length; i++) {
    var c = d.charCodeAt(i);
    if (c >= 768 && c <= 879) continue; // U+0300..U+036F combining marks
    out += d.charAt(i);
  }
  return out;
}

// Canonicalise a PERSON name: whitespace, case, particles, accents.
function canonName_(name) {
  var s = String(name == null ? '' : name).replace(/\s+/g, ' ').trim();
  if (!s) return '';
  s = stripAccents_(s);
  var words = s.split(' ');
  for (var i = 0; i < words.length; i++) {
    var w = words[i];
    if (!w) continue;
    // keep short ALL-CAPS tokens as initials (JP, SP, AB)
    if (w.length <= 3 && w === w.toUpperCase() && /[A-Z]/.test(w)) continue;
    var tc = w.charAt(0).toUpperCase() + w.slice(1).toLowerCase();
    if (i > 0 && CLEAN_PARTICLES.indexOf(tc.toLowerCase()) !== -1) tc = tc.toLowerCase();
    words[i] = tc;
  }
  return words.join(' ');
}

// Apply explicit raw->canonical map (exact, case-sensitive) THEN canonName.
function fixPlayer_(name, map) {
  var raw = String(name == null ? '' : name);
  var trimmed = raw.replace(/\s+/g, ' ').trim();
  if (map && Object.prototype.hasOwnProperty.call(map, trimmed)) return map[trimmed];
  return canonName_(raw);
}

// Fix a TEAM name: trim + collapse spaces, then explicit map. No title-casing
// (teams have deliberate casing like "OB", "8's").
function fixTeam_(name, map) {
  var s = String(name == null ? '' : name).replace(/\s+/g, ' ').trim();
  if (map && Object.prototype.hasOwnProperty.call(map, s)) return map[s];
  return s;
}

// ===========================================================================
//  PER-LEAGUE MAPS
// ===========================================================================

// Super North - team typos -> canonical (canonical = the names in standings).
var NORTH_TEAM_MAP = {
  'Cue strikers':          'Legends Cue Strikers',
  'Innesdale SpinDoctors': 'Innesdale The Spin Doctors',
  'Magicians':             'Legends Magicians',
  'Rack attack':           'Legends Rack Attack'
  // NOT mapped (stray Super-South teams, review separately):
  //   'Crucible CueTastrophe', 'OB Chalked Up'
};

// Super North - player letter-typos -> canonical (case/space/accent are handled
// generically, so only genuine misspellings & abbreviations live here).
var NORTH_PLAYER_MAP = {
  'Esmari Le Crange': 'Esmari le Grange',
  'Adrian V.Taak':    'Adrian van Taak',
  'Adrian Van Touk':  'Adrian van Taak',
  'Dirk V.taak':      'Dirk van Taak',
  'Drik Van Touk':    'Dirk van Taak',
  'Bianca V.Niekerk': 'Bianca van Niekerk',
  'Jared v.d Merwe':  'Jared van der Merwe',
  'Sp Minnaar':       'SP Minnaar'
};

// ===========================================================================
//  ENGINE
// ===========================================================================

function cleanupSuperNorth() {
  var report = cleanupLeague_('Super North', NORTH_TEAM_MAP, NORTH_PLAYER_MAP);
  SpreadsheetApp.getUi().alert('Super North cleanup:\n\n' + report);
}

// Returns a text report. teamMap/playerMap may be {} for a no-op normalise.
function cleanupLeague_(leagueName, teamMap, playerMap) {
  var sheetId = lookupLeagueSheetId(leagueName);
  if (!sheetId) return 'League not configured: ' + leagueName;
  var ss = SpreadsheetApp.openById(sheetId);
  var changed = 0;

  // helper: rewrite one column (1-based) with a transform fn
  function fixCol(sheet, col, fn) {
    if (!sheet || sheet.getLastRow() < 2) return;
    var rng = sheet.getRange(2, col, sheet.getLastRow() - 1, 1);
    var vals = rng.getValues();
    var dirty = false;
    for (var i = 0; i < vals.length; i++) {
      var v = vals[i][0];
      if (v === '' || v == null) continue;
      var nv = fn(v);
      if (nv !== String(v)) { vals[i][0] = nv; dirty = true; changed++; }
    }
    if (dirty) rng.setValues(vals);
  }
  var T = function (v) { return fixTeam_(v, teamMap); };
  var P = function (v) { return fixPlayer_(v, playerMap); };

  // Matches: Home Team(5) Away Team(6) Home Captain(7) Away Captain(8)
  var m = ss.getSheetByName(MATCHES_TAB);
  fixCol(m, 5, T); fixCol(m, 6, T); fixCol(m, 7, P); fixCol(m, 8, P);

  // Frames: Home Team(5) Away Team(6) Home Player(9) Away Player(10)
  var f = ss.getSheetByName(FRAMES_TAB);
  fixCol(f, 5, T); fixCol(f, 6, T); fixCol(f, 9, P); fixCol(f, 10, P);

  // Players: Team(5) Player(7)
  var pl = ss.getSheetByName(PLAYERS_TAB);
  fixCol(pl, 5, T); fixCol(pl, 7, P);

  // Historical Seed: col A = Team (both sub-tables), col B = Player / number.
  var seed = ss.getSheetByName('Historical Seed');
  fixCol(seed, 1, T); fixCol(seed, 2, P);

  // Team Codes: map team names, then drop rows that became duplicates.
  var removed = dedupeTeamCodes_(ss, teamMap);

  // Rebuild standings from the now-clean data.
  rebuildStandingsForLeague(leagueName);

  return 'cells rewritten: ' + changed +
         '\nduplicate team-code rows removed: ' + removed +
         '\nstandings rebuilt.';
}

// Maps the Team Codes 'Team' column then removes rows whose team duplicates an
// earlier row (keeps the first / lowest-code row). Returns count removed.
function dedupeTeamCodes_(ss, teamMap) {
  var sheet = ss.getSheetByName(TEAM_CODES_TAB);
  if (!sheet || sheet.getLastRow() < 2) return 0;
  var n = sheet.getLastRow() - 1;
  var rng = sheet.getRange(2, 1, n, 2);
  var vals = rng.getValues();
  var seen = {}, keep = [], removed = 0;
  for (var i = 0; i < vals.length; i++) {
    var team = fixTeam_(vals[i][0], teamMap);
    var code = vals[i][1];
    if (!team) continue;
    var key = team.toLowerCase();
    if (seen[key]) { removed++; continue; }
    seen[key] = true;
    keep.push([team, code]);
  }
  // clear old data rows, write the kept set back
  sheet.getRange(2, 1, n, 2).clearContent();
  if (keep.length) sheet.getRange(2, 1, keep.length, 2).setValues(keep);
  return removed;
}

// ===========================================================================
//  REMOVE STRAY / MIS-FILED TEAMS (whole rows)
// ===========================================================================
// Deletes every Matches/Frames/Players/Team-Codes row belonging to a team that
// does not belong in this league (e.g. a match submitted to the wrong league),
// then rebuilds standings. Matches by exact team name (case/space-insensitive).
function deleteStrayNorth() {
  var report = deleteStrayTeams_('Super North', ['Crucible CueTastrophe', 'OB Chalked Up']);
  SpreadsheetApp.getUi().alert('Super North stray-team removal:\n\n' + report);
}

function deleteStrayTeams_(leagueName, strays) {
  var sheetId = lookupLeagueSheetId(leagueName);
  if (!sheetId) return 'League not configured: ' + leagueName;
  var ss = SpreadsheetApp.openById(sheetId);
  var bad = {};
  strays.forEach(function (t) { bad[String(t).trim().toLowerCase()] = true; });
  var isBad = function (v) { return bad[String(v == null ? '' : v).trim().toLowerCase()] === true; };

  // Drop any row where one of teamCols (1-based) holds a stray team.
  function purge(sheet, teamCols) {
    if (!sheet || sheet.getLastRow() < 2) return 0;
    var w = sheet.getLastColumn(), n = sheet.getLastRow() - 1;
    var rng = sheet.getRange(2, 1, n, w);
    var vals = rng.getValues();
    var keep = [], dropped = 0;
    for (var i = 0; i < vals.length; i++) {
      var drop = false;
      for (var k = 0; k < teamCols.length; k++) {
        if (isBad(vals[i][teamCols[k] - 1])) { drop = true; break; }
      }
      if (drop) dropped++; else keep.push(vals[i]);
    }
    if (dropped) {
      rng.clearContent();
      if (keep.length) sheet.getRange(2, 1, keep.length, w).setValues(keep);
    }
    return dropped;
  }

  var matches = purge(ss.getSheetByName(MATCHES_TAB), [5, 6]);
  var frames  = purge(ss.getSheetByName(FRAMES_TAB),  [5, 6]);
  var players = purge(ss.getSheetByName(PLAYERS_TAB), [5]);
  var codes   = purge(ss.getSheetByName(TEAM_CODES_TAB), [1]);

  rebuildStandingsForLeague(leagueName);
  return 'Matches removed: ' + matches +
         '\nFrames removed: ' + frames +
         '\nPlayers removed: ' + players +
         '\nTeam-code rows removed: ' + codes +
         '\nstandings rebuilt.';
}

// Optional: self-check the canon rules without touching any sheet.
function cleanup_selftest() {
  var t = [
    ['Tinus van eeden', 'Tinus van Eeden'],
    ['Elmore De Reuck', 'Elmore de Reuck'],
    ['Ayanda  Mathaba', 'Ayanda Mathaba'],
    ['JP du Plessis', 'JP du Plessis'],
    ['Harry Janse Van Vuuren', 'Harry Janse van Vuuren']
  ];
  var out = [];
  for (var i = 0; i < t.length; i++) {
    var got = canonName_(t[i][0]);
    out.push((got === t[i][1] ? 'OK  ' : 'FAIL ') + JSON.stringify(t[i][0]) + ' -> ' + JSON.stringify(got));
  }
  Logger.log(out.join('\n'));
  return out.join('\n');
}

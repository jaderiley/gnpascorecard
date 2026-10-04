// ============================================================
//  player.gs : per-player match history  (?action=playerMatches)
// ============================================================
//  GET ?action=playerMatches&league=<L>&team=<T>&player=<P>
//
//  Returns every VERIFIED match the player played in that league, with the
//  per-frame detail, plus season totals that reconcile with the Player
//  Standings row (live matches + Historical Seed where the league has one).
//
//  Data joins on the Submitted timestamp, exactly as rebuildPlayerStandings
//  does (formatTimestamp). Players rows give the authoritative per-match
//  won/played (walkover frames count for the present player since
//  2026-10-04); Frames rows give the frame-by-frame story. Each tab is read
//  once with getRange().getValues().
//
//  Depends on Code.gs globals: lookupLeagueSheetId, formatTimestamp,
//  jsonResponse, MATCHES_TAB / FRAMES_TAB / PLAYERS_TAB, *_HEADERS,
//  OWN_FRAMES_PCT_LEAGUE, readPlayerSeed (seed.gs).
// ============================================================

var PLAYER_CACHE_TTL_ = 1200; // 20 min, same as standings

// Same normalisation idea as roster.gs (trim, lower-case, single-space).
function normPlayerName_(s) {
  return String(s == null ? '' : s).trim().toLowerCase().replace(/\s+/g, ' ');
}

function getPlayerMatchesResponse_(params) {
  var league = String(params.league || '').trim();
  var team = String(params.team || '').trim();
  var player = String(params.player || '').trim();
  if (!league) return jsonResponse({ ok: false, message: 'league parameter required' });
  if (!team) return jsonResponse({ ok: false, message: 'team parameter required' });
  if (!player) return jsonResponse({ ok: false, message: 'player parameter required' });

  var cache = CacheService.getScriptCache();
  var gen = '0';
  try { gen = cache.get('gen:' + league) || '0'; } catch (ce) { /* best effort */ }
  var key = 'player:' + league + ':' + gen + ':' + team + ':' + normPlayerName_(player);
  // Cache keys are limited to 250 chars; hash if a name is absurdly long.
  if (key.length > 240) key = key.substring(0, 200) + ':' + key.length;
  try {
    var hit = cache.get(key);
    if (hit) return ContentService.createTextOutput(hit).setMimeType(ContentService.MimeType.JSON);
  } catch (ce2) { /* compute it */ }

  var sheetId = lookupLeagueSheetId(league);
  if (!sheetId) return jsonResponse({ ok: false, message: 'Unknown league: ' + league });

  var result = buildPlayerMatches_(SpreadsheetApp.openById(sheetId), league, team, player);
  var payload = JSON.stringify(result);
  try { cache.put(key, payload, PLAYER_CACHE_TTL_); } catch (ce3) { /* too big or cache down */ }
  return ContentService.createTextOutput(payload).setMimeType(ContentService.MimeType.JSON);
}

function playerDateString_(d) {
  if (d instanceof Date) return Utilities.formatDate(d, 'Africa/Johannesburg', 'yyyy-MM-dd');
  return String(d || '').substring(0, 10);
}

function readTabValues_(ss, tabName, nCols) {
  var sh = ss.getSheetByName(tabName);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, nCols).getValues();
}

// Pure builder (takes the spreadsheet) so it can be exercised in tests.
function buildPlayerMatches_(ss, league, team, player) {
  var wantTeam = normPlayerName_(team);
  var wantName = normPlayerName_(player);

  var matchRows = readTabValues_(ss, MATCHES_TAB, MATCHES_HEADERS.length);
  var verified = {}; // key -> match info
  matchRows.forEach(function (row) {
    if (row[12] !== true) return;
    verified[formatTimestamp(row[0])] = {
      date: playerDateString_(row[1]),
      home: String(row[4] || ''),
      away: String(row[5] || ''),
      hScore: Number(row[8]) || 0,
      aScore: Number(row[9]) || 0
    };
  });

  // Players tab: this player's per-match totals + team frames-per-position
  // (denominator for the non-Ladies standings %).
  var canonName = '', canonTeam = '';
  var mine = {};          // key -> { fw, fp, side, pos, team }
  var teamMatchFrames = {}; // key -> max framesPlayed among the team's rows
  var playerRows = readTabValues_(ss, PLAYERS_TAB, PLAYERS_HEADERS.length);
  playerRows.forEach(function (row) {
    var key = formatTimestamp(row[0]);
    if (!verified[key]) return;
    if (normPlayerName_(row[4]) !== wantTeam) return;
    var fp = Number(row[8]) || 0;
    if (fp > (teamMatchFrames[key] || 0)) teamMatchFrames[key] = fp;
    if (normPlayerName_(row[6]) !== wantName) return;
    var fw = Number(row[7]) || 0;
    if (!canonName) { canonName = String(row[6]).trim(); canonTeam = String(row[4]).trim(); }
    if (!mine[key]) mine[key] = { fw: 0, fp: 0, side: String(row[3] || ''), pos: String(row[5] || ''), team: String(row[4] || '') };
    mine[key].fw += fw;
    mine[key].fp += fp;
  });

  // Frames tab: detail for the matches this player appears in.
  var framesByKey = {};
  var frameRows = readTabValues_(ss, FRAMES_TAB, FRAMES_HEADERS.length);
  frameRows.forEach(function (row) {
    var key = formatTimestamp(row[0]);
    var m = mine[key];
    if (!m) return;
    var type = String(row[3] || '').trim().toLowerCase();
    var isHome = normPlayerName_(row[4]) === wantTeam;
    var isAway = normPlayerName_(row[5]) === wantTeam;
    var side = isHome ? 'home' : (isAway ? 'away' : '');
    if (!side) return;
    var myNames = String(isHome ? row[8] : row[9]).split('+');
    var found = false;
    for (var i = 0; i < myNames.length; i++) {
      if (normPlayerName_(myNames[i]) === wantName) { found = true; break; }
    }
    if (!found) return;
    var winner = String(row[12] || '').trim().toLowerCase();
    var f = {
      type: type,
      pos: String(isHome ? row[6] : row[7]),
      oppPos: String(isHome ? row[7] : row[6]),
      opp: String(isHome ? row[9] : row[8]).replace(/\s*\+\s*/g, ' + ').trim(),
      won: winner === side
    };
    if (type === 'walkover') f.wo = true;
    if (type === 'race') {
      f.pts = Number(isHome ? row[10] : row[11]) || 0;
      f.oppPts = Number(isHome ? row[11] : row[10]) || 0;
    }
    (framesByKey[key] = framesByKey[key] || []).push(f);
  });

  var matches = [];
  var liveW = 0, liveP = 0, wins = 0, losses = 0, draws = 0;
  Object.keys(mine).forEach(function (key) {
    var m = mine[key], info = verified[key];
    var home = (m.side.toLowerCase() === 'home');
    var teamScore = home ? info.hScore : info.aScore;
    var oppScore = home ? info.aScore : info.hScore;
    var frames = framesByKey[key] || [];
    var fwDetail = 0;
    frames.forEach(function (f) { if (f.won) fwDetail++; });
    var res = teamScore > oppScore ? 'W' : (teamScore < oppScore ? 'L' : 'D');
    if (res === 'W') wins++; else if (res === 'L') losses++; else draws++;
    liveW += m.fw; liveP += m.fp;
    matches.push({
      date: info.date,
      submitted: key,
      team: m.team,
      opp: home ? info.away : info.home,
      home: home,
      teamScore: teamScore,
      oppScore: oppScore,
      result: res,
      pos: m.pos,
      fw: m.fw,
      fp: m.fp,
      // false when the frame list does not add up to the Players row
      // (should not happen; the UI shows the Players numbers regardless).
      detailOk: frames.length === m.fp && fwDetail === m.fw,
      frames: frames
    });
  });
  matches.sort(function (a, b) {
    return a.date < b.date ? 1 : (a.date > b.date ? -1 : (a.submitted < b.submitted ? 1 : -1));
  });

  // Historical Seed (Super North / South): totals only, no per-match list.
  var seedW = 0, seedP = 0, seedFound = false, seedTeamMax = 0;
  try {
    if (typeof readPlayerSeed === 'function') {
      var seed = readPlayerSeed(ss) || {};
      Object.keys(seed).forEach(function (k) {
        var s = seed[k];
        if (normPlayerName_(s.team) !== wantTeam) return;
        if (Number(s.framesPlayed) > seedTeamMax) seedTeamMax = Number(s.framesPlayed);
        if (normPlayerName_(s.name) === wantName) {
          seedFound = true;
          seedW += Number(s.framesWon) || 0;
          seedP += Number(s.framesPlayed) || 0;
        }
      });
    }
  } catch (se) { /* seed unreadable: fall back to live-only totals */ }

  var totalW = liveW + seedW, totalP = liveP + seedP;
  // Same denominator rule as rebuildPlayerStandings.
  var teamFrames = seedTeamMax;
  Object.keys(teamMatchFrames).forEach(function (k) { teamFrames += teamMatchFrames[k]; });
  var ownFrames = (String(league).trim() === OWN_FRAMES_PCT_LEAGUE);
  var denom = ownFrames ? totalP : (teamFrames || totalP);
  var pct = denom > 0 ? (totalW / denom) * 100 : 0;
  var ownPct = totalP > 0 ? (totalW / totalP) * 100 : 0;

  return {
    ok: true,
    league: league,
    team: canonTeam || team,
    player: canonName || player,
    matches: matches,
    totals: {
      matches: matches.length,
      wins: wins, losses: losses, draws: draws,
      framesWon: totalW,
      framesPlayed: totalP,
      pct: pct,            // matches the Player Standings % column
      ownPct: ownPct,      // framesWon / framesPlayed
      liveFramesWon: liveW,
      liveFramesPlayed: liveP,
      seedFramesWon: seedW,
      seedFramesPlayed: seedP,
      seedFound: seedFound
    }
  };
}

var GNPA_VER_PLAYER = '31adfb2e';

/**
 * ONE-OFF SEED — Ladies League 2026 full roster reload (2026-07-09)
 *
 * Replaces the placeholder/partial Ladies roster (loaded 2026-07-05) with the
 * final, authoritative team list from the league manager. Players were
 * reshuffled (e.g. Adaleen Naidoo moved Funky Chicks -> Terminators) and the
 * Terminators moved from the Legends division to Musketeers, so this does a
 * full REPLACE of both the Roster and Team Codes tabs — not an append.
 *
 * Team codes are fresh random 4-digit codes (not sequential / not guessable),
 * one per team, shared with that team's captain.
 *
 * Run once from the Apps Script editor:  Run -> seedLadies2026()
 * Idempotent: re-running clears and rewrites the same two tabs.
 */
function seedLadies2026() {
  var LEAGUE = 'Ladies';
  var sheetId = lookupLeagueSheetId(LEAGUE);
  if (!sheetId) throw new Error('No sheet configured for league "' + LEAGUE + '" in Config.');
  var ss = SpreadsheetApp.openById(sheetId);

  // [team, code, [ [player, isCaptain], ... ] ]
  var TEAMS = [
    ["Crucibles Blazing 8's", '7900', [
      ['Letitia Calitz', false], ['Alex Fourie', false], ['Michelle Human', true],
      ['Mellanie Ann Uys', false], ['Suné Poolman', false], ['Cameron Marais', false]]],
    ['Musketeers Funky Chicks', '2346', [
      ['Ramona Maré', true], ['Martie Botes', false], ['Salma Ally', false],
      ['Chevonne Lyth', false], ['Doreen du Toit', false], ['Igerdine Gaabusi', false]]],
    ['Musketeers Terminators', '2665', [
      ['Wanita Maritz', true], ['Lucindy Ferreira', false], ['Sandy De Klerk', false],
      ['Adaleen Naidoo', false], ['Sinazo Nxedlana', false], ['Celeste Brown', false],
      ['Lauren Afrika', false]]],
    ['OBS Mafia', '9923', [
      ['Chantelle Penny', false], ['Waltraut Bothma', true], ['Esmari Le Grange', false],
      ['Antoinette Penny', false], ['Laurette le Roux', false], ['Carina Penny', false]]],
    ['OB Flawless Force', '5141', [
      ['Vonnett Murray', true], ['Martie Jansen', false], ['Vinette Nel', false],
      ['Sonette Martin', false], ['Nomoya Dladla', false], ['Lizel Hartzenberg', false]]],
    ['OB Chicks with Sticks', '4133', [
      ['Nicole Spengler', true], ['Leticia Grobler', false], ['Simoné Grobbelaar', false],
      ['Tanika Jonck', false], ['Suanét Esterhuizen', false], ['Chantel Rossouw', false]]],
    ['OB Cue Catz', '2669', [
      ['Clarissa Koen', false], ['Angelique Potgieter', false], ['Colleen van Rooyen', false],
      ['Simoné van Deemter', false], ['Mari Potgieter', false], ['Cait-Lynn van der Walt', true]]],
    ['Legends Bitch Squad', '6889', [
      ['Ulindie Bruwer', true], ['Roma Breetzke', false], ['Daleen von Gordon', false],
      ['Theresa Senekal', false], ['Marlise van Zyl', false], ['Ella Jacobs', false]]],
    ['Legends Cherries on Fire', '8235', [
      ['Caroline de Beer', true], ['Lianie Terblanche', false], ['Franci Henning', false],
      ['Anastasia Bezuidenhout', false], ['Delmarie Lessing-Venter', false], ['Jaqueline Jacobs', false]]],
    ['Legends Dreamers', '2542', [
      ['Moya de Bruto', false], ['Shjantell-Meri Lee', false], ['Bonita Bernice du Plessis', false],
      ['Natasha Reid', false], ['Leandre de Lange', false], ['Judy Grundling', true]]],
    ['Legends Ladypool', '8717', [
      ['Lynette van Stryp', true], ['Truda Brummelkamp', false], ['Lindy van der Merwe', false],
      ['Enid du Plessis', false], ['DelMarie Snyman', false], ['Patricia van der Walt', false]]],
    ['Legends Silly Shots', '3877', [
      ['Susan Booyens', true], ['Retha Steyn', false], ['Arlene Deysel', false],
      ['Tanja Saunders', false], ['Annelise Roos', false], ['Cheri van Staden', false]]],
    ['Legends Pooligans 2.0', '4032', [
      ['Shantel van Staden', true], ['Debby van der Vyver', false], ['Zeandra Crafford', false],
      ['Chané vd Westhuizen', false], ['Nicole Redelinghuys', false], ['Chantelle Pieterse', false]]],
    ['Legends Ubuntu', '9938', [
      ['Cindi van Staden', true], ['Kaylynne van Staden', false], ['Patrys Rossouw', false],
      ['Renette le Roux', false], ['Chené Grace', false], ['Charlotte Grace', false]]],
    ['Legends Angels', '2450', [
      ['Madelein Viljoen', true], ['Charney Bahré', false], ['Sendra Bulunga', false],
      ['Leandri Hawkins', false], ['Wanita van Heerden', false], ['Lynnique Botha', false]]],
    ['Legends Queens', '6728', [
      ['Amanda Pretorius', true], ['Shirley de Beer', false], ['Annalize Vermaak', false],
      ['Yolandie Hands', false], ['Leonie van den Berg', false], ['Marlize Lourens', false]]],
    ['Legends Obsidian Queens', '9239', [
      ['Avalon Uitzinger', false], ['Marissa Hattingh', false], ['Benette Hattingh', false],
      ['Jackie van der Westhuizen', false], ['Thato Legodi', false], ['Stephanie Pass', true]]],
    ['Legends Drinkerbells', '6218', [
      ['Christine Venter', true], ['Zelda Delport', false], ['Amina Taffa', false],
      ['Genna Abdulrahim', false], ['Myrinda Lemmer', false], ['Mika Terbalance', false]]]
  ];

  // ---- Team Codes tab: clear + rewrite ----
  var codeSheet = ss.getSheetByName(TEAM_CODES_TAB) || ss.insertSheet(TEAM_CODES_TAB);
  codeSheet.clear();
  codeSheet.getRange(1, 1, 1, TEAM_CODES_HEADERS.length).setValues([TEAM_CODES_HEADERS]).setFontWeight('bold');
  codeSheet.setFrozenRows(1);
  codeSheet.getRange(2, 2, Math.max(codeSheet.getMaxRows() - 1, 1), 1).setNumberFormat('@'); // codes as text
  var codeRows = TEAMS.map(function (t) { return [t[0], t[1]]; });
  codeSheet.getRange(2, 1, codeRows.length, 2).setValues(codeRows);

  // ---- Roster tab: clear + rewrite ----
  var rSheet = ss.getSheetByName(ROSTER_TAB) || ss.insertSheet(ROSTER_TAB);
  rSheet.clear();
  rSheet.getRange(1, 1, 1, ROSTER_HEADERS.length).setValues([ROSTER_HEADERS]).setFontWeight('bold');
  rSheet.setFrozenRows(1);
  var rosterRows = [];
  TEAMS.forEach(function (t) {
    t[2].forEach(function (p) { rosterRows.push([t[0], p[0], p[1], '']); });
  });
  rSheet.getRange(2, 1, rosterRows.length, ROSTER_HEADERS.length).setValues(rosterRows);
  rSheet.getRange(2, 3, rosterRows.length, 1).insertCheckboxes(); // Captain? column

  // ---- Verify against the live union (catches any leftover seed names) ----
  var live = getRosterForLeague(ss);
  var report = 'Ladies seed done: ' + TEAMS.length + ' teams, ' + rosterRows.length + ' players.\n' +
    'Live union now reports ' + live.teams.length + ' teams.\n' +
    'Teams: ' + live.teams.join(', ');
  Logger.log(report);
  return report;
}

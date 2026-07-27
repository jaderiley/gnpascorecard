/**
 * colorsync.gs - one-off: copy the Ladies league's tab colours onto every
 * other GNPA league sheet, matching tabs by name. Safe to re-run (skips tabs
 * already the right colour, and any tab name Ladies doesn't have).
 *
 * Ran 2026-07-27 from the Master project's Apps Script editor: 64 tabs
 * recoloured across all 8 other leagues. Ladies scheme at that time:
 *   green (#00ff00): Refresh, Manager, Matches
 *   red   (#ff0000): Team Codes, Team Standings, Player Standings, Frames, Players
 */
function syncGnpaTabColors() {
  var LEAGUES = {
    'Vets Tier 1': '17UKU0bzL4V9ajRLJ0wnxFjfTgMEDhKjM_DsxyGwSaOw',
    'Vets Tier 2': '1U-0EUJeemM10eEuOccUsUXsg5xnOL65YeuXhsBS6vGk',
    'Super North': '1GXzVZECbFBuLH_IcQvX1SHYWG7xKiyB1zb3mjXEZ4UY',
    'Super South': '1h_BE71zH0-vV8m6wYnnWnlr-vg7qwpUvSscgtDvAnj0',
    'Premier':     '1tvU6pL1Nos8F1XQFRuwxvJztZpp6TFSoxJJdUKRS6Ng',
    'Ladies':      '1ZfV6x36MEbtNPn3thnQDt_UXOSrbc_PGB7in0hk8pKM',
    'Tshwane':     '1AZ9I2ICPgCnwoPi2ehmSlvTxLqxsOGmqV5pBhUQL9Ho',
    '3-Man':       '1AI1Hpzup-xbOtZkf0yskptCQStGNbNiMQJjE0kWhiBU',
    'Juniors':     '1UWTV1JZth_RW04jeSmovHIW48uofhVecjyGzPT66y-w'
  };
  var src = SpreadsheetApp.openById(LEAGUES['Ladies']);
  var colors = {};
  src.getSheets().forEach(function (s) {
    var c = s.getTabColor();
    if (c) colors[s.getName()] = c;
  });
  Logger.log('Ladies template: ' + JSON.stringify(colors));
  var changes = 0;
  Object.keys(LEAGUES).forEach(function (league) {
    if (league === 'Ladies') return;
    var ss = SpreadsheetApp.openById(LEAGUES[league]);
    ss.getSheets().forEach(function (sheet) {
      var name = sheet.getName();
      var want = colors[name];
      if (!want) return;
      var cur = sheet.getTabColor() || '';
      if (cur.toLowerCase() === want.toLowerCase()) return;
      sheet.setTabColor(want);
      changes++;
      Logger.log(league + ': ' + name + ' -> ' + want + '  (was ' + (cur || 'none') + ')');
    });
  });
  Logger.log('done, ' + changes + ' tab(s) recoloured');
}

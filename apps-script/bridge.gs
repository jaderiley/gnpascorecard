/**
 * bridge.gs — generic, secret-gated Google Sheets read/write endpoint.
 *
 * Turns this already-deployed web app into a small "Sheets bridge" so tooling
 * (a local MCP server) can read and write any spreadsheet this account can open,
 * via one authenticated POST — no Google Cloud service account, no extra OAuth.
 *
 * SECURITY
 *   - Every call must carry the shared secret in `secret`, matched against the
 *     Script Property BRIDGE_SECRET (Project Settings -> Script Properties).
 *     No secret is ever stored in this file or the repo.
 *   - Optional Script Property BRIDGE_ALLOWLIST (comma-separated spreadsheet
 *     IDs) locks writes to specific sheets. If unset, any sheet the account can
 *     open is allowed (the secret is the control).
 *   - Wired in from doPost:  if (data.action === 'bridge') return
 *     jsonResponse(handleBridge(data));  — it runs BEFORE handleSubmission and
 *     is fully isolated from the match-submission path (which sends no `action`).
 *
 * REQUEST  { action:'bridge', secret, op, spreadsheetId, tab, ... }
 *   op 'read'   { range? }                      -> { ok, values }
 *   op 'write'  { cell='A1', values:[[...]], clearFromRow? } -> { ok, wrote:{rows,cols} }
 *   op 'append' { values:[[...]] }              -> { ok, appended, atRow }
 *   op 'clear'  { range? , fromRow? }           -> { ok }
 *   op 'tabs'   { }                             -> { ok, tabs:[names] }
 */

function handleBridge(data) {
  // ---- auth ----
  var props = PropertiesService.getScriptProperties();
  var secret = props.getProperty('BRIDGE_SECRET');
  if (!secret) return { ok: false, message: 'BRIDGE_SECRET not set on this project.' };
  if (String(data.secret || '') !== secret) return { ok: false, message: 'unauthorized' };

  var id = String(data.spreadsheetId || '').trim();
  if (!id) return { ok: false, message: 'spreadsheetId required' };

  var allow = (props.getProperty('BRIDGE_ALLOWLIST') || '').split(',')
    .map(function (s) { return s.trim(); }).filter(String);
  if (allow.length && allow.indexOf(id) === -1) {
    return { ok: false, message: 'spreadsheetId not in BRIDGE_ALLOWLIST' };
  }

  var ss;
  try { ss = SpreadsheetApp.openById(id); }
  catch (e) { return { ok: false, message: 'cannot open spreadsheet: ' + e.message }; }

  var op = String(data.op || '').toLowerCase();

  if (op === 'tabs') {
    return { ok: true, tabs: ss.getSheets().map(function (s) { return s.getName(); }) };
  }

  var tabName = String(data.tab || '').trim();
  if (!tabName) return { ok: false, message: 'tab required' };
  var sheet = ss.getSheetByName(tabName);

  if (op === 'read') {
    if (!sheet) return { ok: false, message: 'no tab named "' + tabName + '"' };
    var rng = data.range ? sheet.getRange(data.range)
                         : sheet.getDataRange();
    return { ok: true, values: rng.getValues() };
  }

  if (op === 'write') {
    if (!sheet) sheet = ss.insertSheet(tabName);
    var values = data.values;
    if (!Array.isArray(values) || !values.length || !Array.isArray(values[0])) {
      return { ok: false, message: 'values must be a non-empty 2D array' };
    }
    // Optional: wipe existing rows from `clearFromRow` down before writing.
    if (data.clearFromRow && sheet.getLastRow() >= data.clearFromRow) {
      sheet.getRange(data.clearFromRow, 1,
        sheet.getLastRow() - data.clearFromRow + 1,
        Math.max(sheet.getLastColumn(), 1)).clearContent();
    }
    var anchor = sheet.getRange(data.cell || 'A1');
    sheet.getRange(anchor.getRow(), anchor.getColumn(), values.length, values[0].length)
      .setValues(values);
    return { ok: true, wrote: { rows: values.length, cols: values[0].length } };
  }

  if (op === 'append') {
    if (!sheet) sheet = ss.insertSheet(tabName);
    var vals = data.values;
    if (!Array.isArray(vals) || !vals.length || !Array.isArray(vals[0])) {
      return { ok: false, message: 'values must be a non-empty 2D array' };
    }
    var startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, vals.length, vals[0].length).setValues(vals);
    return { ok: true, appended: vals.length, atRow: startRow };
  }

  if (op === 'clear') {
    if (!sheet) return { ok: false, message: 'no tab named "' + tabName + '"' };
    if (data.range) { sheet.getRange(data.range).clearContent(); }
    else if (data.fromRow && sheet.getLastRow() >= data.fromRow) {
      sheet.getRange(data.fromRow, 1, sheet.getLastRow() - data.fromRow + 1,
        Math.max(sheet.getLastColumn(), 1)).clearContent();
    } else { sheet.clearContents(); }
    return { ok: true };
  }

  return { ok: false, message: 'unknown op: ' + op };
}

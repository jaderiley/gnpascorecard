/**
 * version.gs - backend drift detector.
 *
 * Every .gs file in the repo ends with one line:  var GNPA_VER_<FILE> = '<hash>';
 * The hash is a short sha256 of that file's content EXCLUDING the stamp line.
 * tools/stamp_versions.py rewrites the stamps (run it BEFORE pasting into the
 * script editor). getBackendVersions() reports whatever stamps the LIVE script
 * actually contains, so a stale or missing file shows up by comparing against
 * the repo: GET <exec URL>?action=version  (route lives in Code.gs doGet).
 *
 * A file that was never pasted shows 'missing'. seed.gs is not in the repo and
 * is not tracked here.
 */

function getBackendVersions() {
  var out = {};
  // BEGIN GNPA_VER_NAMES (managed by tools/stamp_versions.py)
  var names = {
    'bridge.gs': function () { return typeof GNPA_VER_BRIDGE !== 'undefined' ? GNPA_VER_BRIDGE : 'missing'; },
    'cleanup.gs': function () { return typeof GNPA_VER_CLEANUP !== 'undefined' ? GNPA_VER_CLEANUP : 'missing'; },
    'Code.gs': function () { return typeof GNPA_VER_CODE !== 'undefined' ? GNPA_VER_CODE : 'missing'; },
    'compact.gs': function () { return typeof GNPA_VER_COMPACT !== 'undefined' ? GNPA_VER_COMPACT : 'missing'; },
    'import.gs': function () { return typeof GNPA_VER_IMPORT !== 'undefined' ? GNPA_VER_IMPORT : 'missing'; },
    'manager.gs': function () { return typeof GNPA_VER_MANAGER !== 'undefined' ? GNPA_VER_MANAGER : 'missing'; },
    'player.gs': function () { return typeof GNPA_VER_PLAYER !== 'undefined' ? GNPA_VER_PLAYER : 'missing'; },
    'refresh.gs': function () { return typeof GNPA_VER_REFRESH !== 'undefined' ? GNPA_VER_REFRESH : 'missing'; },
    'roster.gs': function () { return typeof GNPA_VER_ROSTER !== 'undefined' ? GNPA_VER_ROSTER : 'missing'; },
    'version.gs': function () { return typeof GNPA_VER_VERSION !== 'undefined' ? GNPA_VER_VERSION : 'missing'; }
  };
  // END GNPA_VER_NAMES
  Object.keys(names).forEach(function (file) {
    out[file] = names[file]();
  });
  return out;
}

var GNPA_VER_VERSION = '2434ce80';

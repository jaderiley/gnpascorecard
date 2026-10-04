#!/usr/bin/env python
"""stamp_versions.py - backend drift detector for the GNPA Apps Script copies.

Each apps-script/*.gs file ends with one line:
    var GNPA_VER_<FILE> = '<hash>';
where <hash> is the first 8 hex chars of sha256 over the file's content with
that line removed (CRLF normalised to LF, trailing whitespace trimmed).

  python tools/stamp_versions.py            stamp every .gs file (idempotent)
  python tools/stamp_versions.py --check    exit 1 if any file is unstamped/stale
  python tools/stamp_versions.py --live     compare repo stamps with the LIVE script
                                            (GET ?action=version, one cheap request)

Run the stamp step BEFORE pasting files into the Apps Script editor. Stdlib only.
"""
import hashlib
import json
import re
import sys
import urllib.request
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
GS_DIR = ROOT / "apps-script"
EXEC_URL = ("https://script.google.com/macros/s/AKfycbxOZ-8jH0bL-DjbWsLZx59sxUXR8PHl"
            "IS3mEr0Sv4YUhQWRVygt1Qt0zdlwCQqIsPqzJQ/exec")

STAMP_RE = re.compile(r"^var GNPA_VER_\w+ = '[0-9a-f]*';[ \t]*\r?\n?", re.M)
NAMES_RE = re.compile(r"(  // BEGIN GNPA_VER_NAMES[^\n]*\n)(.*?)(  // END GNPA_VER_NAMES)", re.S)


def var_name(path):
    return "GNPA_VER_" + re.sub(r"[^A-Za-z0-9]+", "_", path.stem).upper()


def gs_files():
    return sorted(GS_DIR.glob("*.gs"), key=lambda p: p.name.lower())


def digest(text):
    """Hash of the content without its stamp line."""
    body = STAMP_RE.sub("", text.replace("\r\n", "\n")).rstrip()
    return hashlib.sha256(body.encode("utf-8")).hexdigest()[:8]


def stamped(text, name, eol):
    body = STAMP_RE.sub("", text.replace("\r\n", "\n")).rstrip("\n \t")
    h = hashlib.sha256(body.encode("utf-8")).hexdigest()[:8]
    out = body + "\n\n" + f"var {name} = '{h}';" + "\n"
    return out.replace("\n", eol), h


def names_block(files):
    rows = []
    for p in files:
        n = var_name(p)
        rows.append(f"    '{p.name}': function () {{ return typeof {n} !== 'undefined' ? {n} : 'missing'; }},")
    if rows:
        rows[-1] = rows[-1].rstrip(",")
    return "\n".join(rows) + "\n"


def read(p):
    # newline='' keeps the file's own line endings
    with open(p, encoding="utf-8", newline="") as f:
        return f.read()


def write(p, s):
    with open(p, "w", encoding="utf-8", newline="") as f:
        f.write(s)


def repo_stamps():
    """{file name: stamp or None} read from the repo as it is right now."""
    out = {}
    for p in gs_files():
        m = re.search(r"^var GNPA_VER_\w+ = '([0-9a-f]*)';", read(p), re.M)
        out[p.name] = m.group(1) if m else None
    return out


def stamp_all():
    files = gs_files()
    changed = []
    # version.gs first: its names block depends on the file set, and its own hash covers it.
    ordered = sorted(files, key=lambda p: p.name != "version.gs")
    for p in ordered:
        text = read(p)
        eol = "\r\n" if text.count("\r\n") * 2 >= max(text.count("\n"), 1) else "\n"
        if p.name == "version.gs":
            nl = text.replace("\r\n", "\n")

            # rebuild the names map between the markers
            def repl(m):
                return m.group(1) + "  var names = {\n" + names_block(files) + "  };\n" + m.group(3)
            nl = NAMES_RE.sub(repl, nl)
            text = nl.replace("\n", eol)
        new, h = stamped(text, var_name(p), eol)
        if new != read(p):
            write(p, new)
            changed.append(p.name)
        print(f"{p.name:24s} {var_name(p):28s} {h}")
    print(f"\n{len(changed)} file(s) rewritten" + (": " + ", ".join(changed) if changed else " (already current)"))
    return changed


def check_repo():
    bad = []
    for p in gs_files():
        text = read(p)
        m = re.search(r"^var GNPA_VER_\w+ = '([0-9a-f]*)';", text, re.M)
        if not m or m.group(1) != digest(text):
            bad.append(p.name)
    return bad


def live_versions(timeout=15):
    """Return (versions dict, None) or (None, reason). One cheap GET."""
    req = urllib.request.Request(EXEC_URL + "?action=version", headers={"User-Agent": "gnpa-version-check"})
    try:
        with urllib.request.urlopen(req, timeout=timeout) as r:
            raw = r.read().decode("utf-8", "replace")
    except Exception as e:
        return None, f"unreachable ({type(e).__name__})"
    try:
        data = json.loads(raw)
    except ValueError:
        return None, "non-JSON reply"
    if isinstance(data.get("versions"), dict):
        return data["versions"], None
    return None, "predates"   # old script: falls through to the 'endpoint live' banner


def compare_live():
    """Return (level, message); level in 'ok' | 'warn' | 'info'."""
    unstamped = check_repo()
    if unstamped:
        return "warn", "repo stamps stale, run tools/stamp_versions.py: " + ", ".join(unstamped)
    live, why = live_versions()
    if live is None:
        if why == "predates":
            return "info", "live script predates version check (paste updated files, then redeploy)"
        return "info", f"could not check live script: {why}"
    repo = repo_stamps()
    drift = []
    for name, want in repo.items():
        got = live.get(name, "missing")
        if got != want:
            drift.append(f"{name} ({'not pasted' if got == 'missing' else 'live ' + str(got) + ' vs repo ' + str(want)})")
    if drift:
        return "warn", "live script differs from repo: " + "; ".join(drift)
    return "ok", f"live script matches repo ({len(repo)} files)"


def main():
    if "--check" in sys.argv:
        bad = check_repo()
        print("stale/unstamped: " + ", ".join(bad) if bad else "all stamps current")
        return 1 if bad else 0
    if "--live" in sys.argv:
        level, msg = compare_live()
        print(f"[{level}] {msg}")
        return 1 if level == "warn" else 0
    stamp_all()
    return 0


if __name__ == "__main__":
    sys.exit(main())

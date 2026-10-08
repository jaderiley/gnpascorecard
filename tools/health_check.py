#!/usr/bin/env python
"""health_check.py - read-only synthetic check of the live GNPA Pool App.

Wired into claude-dev/_hooks/system_check.py (daily scorecard). It NEVER writes:
only public GETs plus bridge READS (op=read) of each league's Matches tab.

Checks
  * site      gnpascoring.co.za answers 200, looks like the app, and its
              index.html equals the repo's origin/main copy (Pages deploy lag)
  * leagues   ?action=leagues answers; every Config league is in gnpa_cli's map
  * per league (all but GNPA Test League), in parallel:
      - ?action=standings and ?action=roster answer valid JSON (one retry,
        mirroring the app, because a cold standings call sometimes dies on
        Google's echo hop with a 404 "Drive page not found")
      - standings maths: played = W+L+D, points = 2W+D, framesWon <= framesPlayed,
        % = framesWon/framesPlayed, sum W = sum L, sum D even, no duplicate
        teams/players, every standings team on the roster
      - unseeded leagues: sum of played = 2 x verified matches (stale standings
        => tick the league's Refresh box)
      - stuck cards: fixtures with no verified copy older than STUCK_DAYS,
        minus a known baseline (Super South's 35 June fixtures)
  Apps Script version drift is checked separately (stamp_versions.compare_live).

Call volume per run: 1 leagues + 9 x (standings + roster + 1 bridge read) = 28.

  python tools/health_check.py          print results, exit 1 if any warn
Stdlib only. Reuses gnpa_cli.py (the gnpa skill) for the bridge + league map.
"""
import hashlib
import importlib.util
import json
import subprocess
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor
from datetime import date, timedelta
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
CLI_PATH = Path(r"C:\Users\jadea\.agents\skills\gnpa\scripts\gnpa_cli.py")
SITE = "https://gnpascoring.co.za/"
SKIP = {"GNPA Test League"}
SEEDED = {"Super North", "Super South"}   # Historical Seed tab: standings != app matches
STUCK_DAYS = 14
# Known, accepted backlogs: fixtures dated on/before `through` are counted but
# never flagged. Super South's June round was never verified (2026-10-04 audit);
# the league stopped submitting after 24/06.
BASELINE = {"Super South": {"through": "2026-06-30", "count": 35,
                            "note": "35 June fixtures never verified (2026-10-04 audit)"}}

OK, INFO, WARN = "ok", "info", "warn"


def _cli():
    spec = importlib.util.spec_from_file_location("gnpa_cli_hc", CLI_PATH)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def _get(url, timeout=90):
    req = urllib.request.Request(url, headers={"User-Agent": "gnpa-health-check"})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return r.status, r.read()


def _endpoint(cli, action, league=None, tries=2):
    """GET a public action; returns (json|None, seconds, attempts, last_error)."""
    url = cli.ENDPOINT + "?action=" + action
    if league:
        url += "&league=" + urllib.parse.quote(league)
    last, t0 = None, time.time()
    for i in range(1, tries + 1):
        try:
            _, raw = _get(url)
            data = json.loads(raw.decode("utf-8", "replace"))
            if isinstance(data, dict) and data.get("ok"):
                return data, time.time() - t0, i, None
            last = f"ok=false: {str(data.get('message') if isinstance(data, dict) else data)[:80]}"
        except urllib.error.HTTPError as e:
            last = f"HTTP {e.code}"
        except ValueError:
            last = "non-JSON reply (Google error page)"
        except Exception as e:
            last = type(e).__name__
        time.sleep(2)
    return None, time.time() - t0, tries, last


# ── site ──────────────────────────────────────────────────────────────────
def check_site():
    try:
        status, raw = _get(SITE, timeout=20)
    except Exception as e:
        return [(WARN, "GNPA site", f"{SITE} unreachable ({type(e).__name__})")]
    body = raw.decode("utf-8", "replace")
    if status != 200 or "GNPA" not in body or "macros/s/" not in body:
        return [(WARN, "GNPA site", f"{SITE} answered {status} but does not look like the app")]
    live = hashlib.sha256(body.replace("\r", "").encode()).hexdigest()
    try:
        repo = subprocess.run(["git", "-C", str(ROOT), "show", "origin/main:index.html"],
                              capture_output=True, timeout=20).stdout.decode("utf-8", "replace")
        same = hashlib.sha256(repo.replace("\r", "").encode()).hexdigest() == live
    except Exception:
        same = None
    if same is False:
        return [(WARN, "GNPA site", "live index.html differs from repo origin/main "
                 "(Pages still deploying, or local repo needs a git fetch)")]
    return [(OK, "GNPA site", f"200, {len(raw) // 1024} KB, matches repo main")]


# ── per league ─────────────────────────────────────────────────────────────
def check_standings(name, st, roster_teams):
    errs = []
    teams = st.get("teams") or []
    seen = set()
    sw = sl = sd = 0
    for t in teams:
        n = t["team"]
        w, l, d, p = t["w"], t["l"], t["d"], t["played"]
        sw, sl, sd = sw + w, sl + l, sd + d
        if n.lower() in seen:
            errs.append(f"duplicate team '{n}'")
        seen.add(n.lower())
        if p != w + l + d:
            errs.append(f"{n}: played {p} != W+L+D {w + l + d}")
        if t["points"] != 2 * w + d:
            errs.append(f"{n}: points {t['points']} != 2W+D {2 * w + d}")
        if t["framesWon"] > t["framesPlayed"]:
            errs.append(f"{n}: frames won {t['framesWon']} > played {t['framesPlayed']}")
        if t["framesPlayed"] and abs(t["pct"] - 100 * t["framesWon"] / t["framesPlayed"]) > 0.15:
            errs.append(f"{n}: % {t['pct']:.1f} != frames ratio")
        if roster_teams is not None and n.lower() not in roster_teams:
            errs.append(f"standings team '{n}' not on the roster")
    if sw != sl:
        errs.append(f"total wins {sw} != total losses {sl}")
    if sd % 2:
        errs.append(f"total draws {sd} is odd")
    pseen = set()
    for pl in st.get("players") or []:
        if not pl["team"]:          # 'LESS THAN x% ...' heading row
            continue
        k = (pl["team"].lower(), pl["player"].lower())
        if k in pseen:
            errs.append(f"duplicate player row {pl['team']} / {pl['player']}")
        pseen.add(k)
        if pl["framesWon"] > pl["framesPlayed"]:
            errs.append(f"player {pl['player']}: won {pl['framesWon']} > played {pl['framesPlayed']}")
    return errs, sum(t["played"] for t in teams)


def check_matches(cli, name, info, today):
    rows = [r for r in cli.read_tab(info["id"], "Matches")[1:] if r and str(r[0]).strip()]
    def is_v(r):
        return len(r) > 12 and r[12] is True

    def pair(r):    # order-insensitive fixture: (teamA, teamB) -> (scoreA, scoreB)
        h, a = str(r[4]).strip().lower(), str(r[5]).strip().lower()
        hs, as_ = cli.num(r[8]), cli.num(r[9])
        return ((h, a), (hs, as_)) if h <= a else ((a, h), (as_, hs))

    verified = [r for r in rows if is_v(r)]
    v_scores = defaultdict(set)    # teams -> verified score tuples
    v_dates = defaultdict(list)    # teams -> verified match dates
    for r in verified:
        teams, sc = pair(r)
        v_scores[teams].add(sc)
        v_dates[teams].append(date.fromisoformat(cli.match_date(r[1])[:10]))
    stuck, dupes = {}, 0   # (teams, score) -> earliest date
    for r in rows:
        if is_v(r):
            continue
        teams, sc = pair(r)
        d = cli.match_date(r[1])[:10]
        # A retap/re-dated copy of a fixture that IS verified (same teams + same
        # score, or same teams within 10 days) is junk, not a stuck card.
        try:
            near = any(abs((date.fromisoformat(d) - vd).days) <= 10 for vd in v_dates[teams])
        except ValueError:
            near = False
        if sc in v_scores[teams] or near:
            dupes += 1
            continue
        k = (teams, sc)            # retaps of an unverified fixture count once
        if k in stuck:
            dupes += 1
        stuck[k] = min(stuck.get(k, d), d)
    stuck = {(d,) + k[0]: 1 for k, d in stuck.items()}
    cutoff = (today - timedelta(days=STUCK_DAYS)).isoformat()
    base = BASELINE.get(name)
    old = [k for k in stuck if k[0] <= cutoff and not (base and k[0] <= base["through"])]
    base_n = sum(1 for k in stuck if base and k[0] <= base["through"])
    young = sum(1 for k in stuck if k[0] > cutoff)
    return {"rows": len(rows), "verified": len(verified), "old": sorted(old),
            "young": young, "baseline": base_n, "dupes": dupes}


def check_league(cli, name, info, today):
    out = []
    roster, rt, ra, rerr = _endpoint(cli, "roster", name)
    st, sti, sa, serr = _endpoint(cli, "standings", name)
    if roster is None:
        out.append((WARN, f"GNPA {name}", f"roster endpoint failed x{ra}: {rerr} (app shows 'Loading teams...')"))
    if st is None:
        out.append((WARN, f"GNPA {name}", f"standings endpoint failed x{sa}: {serr}"))
    if roster is None or st is None:
        return out
    notes = []
    if sa > 1 or ra > 1:
        notes.append(f"needed a retry ({'standings' if sa > 1 else 'roster'})")
    if max(sti, rt) > 45:
        notes.append(f"slow ({max(sti, rt):.0f}s)")
    roster_teams = {t.lower() for t in roster.get("teams") or []}
    errs, played_sum = check_standings(name, st, roster_teams)
    try:
        m = check_matches(cli, name, info, today)
    except BaseException as e:      # gnpa_cli raises SystemExit on bridge failure
        m = None
        notes.append(f"Matches read failed ({str(e)[:60]})")
    if m and name not in SEEDED and played_sum != 2 * m["verified"]:
        errs.append(f"standings stale: {played_sum // 2} matches counted vs {m['verified']} "
                    f"verified (tick the league's Refresh box)")
    nteams = len(st.get("teams") or [])
    summary = f"{nteams} teams in standings / {len(roster_teams)} on roster"
    if m:
        summary += f", {m['rows']} match rows ({m['verified']} verified)"
        if m["young"]:
            summary += f", {m['young']} fixture(s) awaiting verification"
        if m["dupes"]:
            summary += f", {m['dupes']} unverified retap/duplicate row(s)"
        if m["baseline"]:
            summary += f", {m['baseline']} known-baseline unverified"
            exp = BASELINE[name]["count"]
            if m["baseline"] != exp:
                summary += f" (baseline was {exp})"
    if notes:
        summary += "; " + "; ".join(notes)
    if errs:
        out.append((WARN, f"GNPA {name}", "; ".join(errs[:4]) + (f" (+{len(errs) - 4} more)" if len(errs) > 4 else "")))
    if m and m["old"]:
        ex = ", ".join(f"{d} {h} v {a}" for d, h, a in m["old"][:2])
        out.append((WARN, f"GNPA {name}", f"{len(m['old'])} fixture(s) unverified > {STUCK_DAYS} days, "
                    f"not in standings (e.g. {ex}); ask the league manager to verify"))
    out.append((OK, f"GNPA {name}", summary))
    return out


def run(today=None):
    """Return [(level, area, message)]. Never raises."""
    today = today or date.today()
    res = []
    try:
        res += check_site()
    except Exception as e:
        res.append((WARN, "GNPA site", f"check crashed: {e}"))
    try:
        cli = _cli()
    except Exception as e:
        return res + [(WARN, "GNPA leagues", f"cannot load gnpa_cli.py ({e})")]
    lg, _, la, lerr = _endpoint(cli, "leagues")
    if lg is None:
        return res + [(WARN, "GNPA leagues", f"?action=leagues failed x{la}: {lerr} (Apps Script down?)")]
    live = list(lg.get("leagues") or [])
    unknown = [n for n in live if n not in cli.LEAGUES]
    if unknown:
        res.append((WARN, "GNPA leagues", f"Config has leagues not in gnpa_cli map: {unknown}"))
    names = [n for n in live if n in cli.LEAGUES and n not in SKIP]
    with ThreadPoolExecutor(max_workers=4) as ex:
        futs = [ex.submit(check_league, cli, n, cli.LEAGUES[n], today) for n in names]
        for n, f in zip(names, futs):
            try:
                res += f.result()
            except BaseException as e:
                res.append((WARN, f"GNPA {n}", f"check crashed: {str(e)[:100]}"))
    return res


def main():
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
    t0 = time.time()
    res = run()
    for level, area, msg in res:
        print(f"{'⚠ ' if level == WARN else '  '}{area}: {msg}")
    print(f"({time.time() - t0:.0f}s)")
    return 1 if any(r[0] == WARN for r in res) else 0


if __name__ == "__main__":
    sys.exit(main())

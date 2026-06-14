# GNPA backend (Google Apps Script)

The backend is a Google Apps Script web app **bound to the Master sheet**
(`GNPA League — Master`). It is not served from this GitHub repo — these files
are kept here for version control. The live source of truth is the script
editor (Master sheet → Extensions → Apps Script).

## Files
- **Code.gs** — router, submission handler, standings rebuild. *(modified)*
- **roster.gs** — Roster tab, dual-code captain gate, code generation. *(new)*
- **seed.gs** — historical-season seed data + installers. *(unchanged — not
  duplicated here; the copy in the script editor is current.)*

## What changed
- **Dual-code submission gate** — `handleSubmission` now calls
  `validateCaptainCodes` (roster.gs) before writing. Both opposing captains
  must enter their own code; one valid submission still produces one pending
  Matches row. Manager `Verified` approval is untouched.
- **Per-league percentage** — `rebuildPlayerStandings` ranks the **Ladies**
  league on own frames played (unchanged) and **every other league** on the
  team's total frames for the season (`OWN_FRAMES_PCT_LEAGUE = 'Ladies'`).
- **Roster serving** — `handleRosterRequest` now reads the authoritative
  `Roster` tab first, falling back to seed/Players data.

## Roster tab schema (per league sheet)
`Team | Player | Captain? | Code` — appended tab, no existing schema shifts.

## Deploy
1. Master sheet → Extensions → Apps Script.
2. Paste **Code.gs** over the existing file; add a new file **roster.gs** and
   paste its contents. Leave **seed.gs** as-is. Save.
3. Deploy → Manage deployments → edit (pencil) → **New version** → Deploy.
   (The `/exec` URL stays the same, so the app needs no endpoint change.)
4. Reload the Master sheet. New menu items appear under **GNPA League**:
   - *Set up Roster tabs (all leagues)* — creates/seeds each `Roster` tab.
   - Tick **Captain?** for each team's captain(s) in every Roster tab.
   - *Generate captain codes (all leagues)* — fills blank codes for captains.
   - *Reset one captain code* — re-issues a single code if it leaks.
5. *Rebuild standings — ALL leagues* to recompute percentages with the new
   formula.

## Rollout safety
`ROSTER_GATE_STRICT = false` (roster.gs): until a league has captain codes,
submissions behave exactly as before — nothing breaks mid-season. Once a
league has at least one captain code, the gate enforces for that league. Set
to `true` to require codes everywhere.

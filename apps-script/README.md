# GNPA backend (Google Apps Script)

The backend is a Google Apps Script web app **bound to the Master sheet**
(`GNPA League — Master`). It is not served from this GitHub repo — these files
are kept here for version control. The live source of truth is the script
editor (Master sheet → Extensions → Apps Script).

## Files
- **Code.gs** — router, submission handler, standings rebuild. *(modified)*
- **roster.gs** — Team Codes tab, dual team-code gate, code generation, and
  an optional player Roster tab. *(new)*
- **seed.gs** — historical-season seed data + installers. *(unchanged — not
  duplicated here; the copy in the script editor is current.)*

## What changed
- **Dual team-code submission gate** — `handleSubmission` calls
  `validateTeamCodes` (roster.gs) before writing. Each team has ONE season-long
  code (e.g. `0001`, `0002`); both teams enter their own code to confirm a
  result. Anyone on the team can use it — no specific captain required. One
  valid submission still produces one pending Matches row. Manager `Verified`
  approval is untouched.
- **Per-league percentage** — `rebuildPlayerStandings` ranks the **Ladies**
  league on own frames played (unchanged) and **every other league** on the
  team's total frames for the season (`OWN_FRAMES_PCT_LEAGUE = 'Ladies'`).
- **Roster serving** — `handleRosterRequest` reads the optional `Roster` tab
  first, falling back to seed/Players data.

## Tab schemas (per league sheet)
- **Team Codes** — `Team | Code` (the gate). Codes are stored as text so
  leading zeros survive; assigned sequentially per league.
- **Roster** *(optional, dropdowns only)* — `Team | Player | Captain? | Code`.

Both are appended tabs — no existing schema shifts.

## Deploy
1. Master sheet → Extensions → Apps Script.
2. Paste **Code.gs** over the existing file; add a new file **roster.gs** and
   paste its contents. Leave **seed.gs** as-is. Save.
3. Deploy → Manage deployments → edit (pencil) → **New version** → Deploy.
   (The `/exec` URL stays the same, so the app needs no endpoint change.)
4. Reload the Master sheet. New menu items appear under **GNPA League**:
   - *Set up team codes (all leagues)* — creates/tops up each `Team Codes` tab
     with sequential codes (`0001`, `0002`, …). Open the tab to share codes.
   - *Reset one team code* — re-issues a single team's code if it leaks.
   - *Set up player rosters (optional)* — builds the `Roster` dropdowns.
5. *Rebuild standings — ALL leagues* to recompute percentages with the new
   formula.

## Rollout safety
`ROSTER_GATE_STRICT = false` (roster.gs): until a league has team codes,
submissions behave exactly as before — nothing breaks mid-season. Once a
league has team codes, the gate enforces for that league. Set to `true` to
require codes everywhere.

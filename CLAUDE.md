# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project overview

醫院排班 (Hospital Nurse Scheduling System) — a pure client-side static web app with no build step, no package manager, and no test suite. It parses a nurse roster Excel file, runs a simulated-annealing scheduler to auto-fill shifts under a set of hard/soft constraints, lets the user hand-edit cells with live re-validation, and exports the result back to a styled `.xlsx`.

Three files hold the entire app:
- `index.html` — DOM structure only; loads `xlsx-js-style` (SheetJS fork) and FontAwesome from CDN, plus local `style.css` and `app.js`.
- `app.js` — all logic: Excel parsing, the scheduling algorithm, the validator, table rendering/interaction, and Excel export.
- `style.css` — glassmorphic theme with light/dark mode driven by `data-theme` on `<html>` and CSS custom properties (`--bg-day`, `--color-night`, etc.).

## Running / testing

There is no build, lint, or test tooling in this repo. To work on it, just open `index.html` in a browser (or serve the directory with any static file server) and exercise the UI directly with a real roster `.xlsx` file. Per `.agents/skills/nurse_scheduling_rules/SKILL.md`, do **not** attempt automated browser testing after changes — visually/logically confirm the code is correct and report back instead.

## Architecture (`app.js`)

Everything operates on module-level global state, not a framework/store:
- `employees[]` — one entry per parsed roster row: `{ level, id, name, prefStr, preferences[], isHeadNurse, originalSchedule[], isPrefilledFlags[], rowIndex }`.
- `currentSchedule` — a 2D array `[employeeIndex][dayIndex]` mirroring `employees[].originalSchedule` but mutable; this is what gets rendered, validated, scheduled, and exported.
- `originalExcelData` — the raw SheetJS worksheet object, kept around so export can `Object.assign` new values onto existing cells and preserve everything else (styles, comments) not explicitly touched.
- `config` — user-tunable constraints (`minOff`/`maxOff`, `maxConsecutive`, per-shift daily `demand`), bound live to the settings inputs.

Pipeline: `handleFile` → `parseExcelData` (locates the date header row by scanning for ≥28 consecutive integers starting at column E, then reads employee rows below it, then bottom summary rows) → `currentSchedule` initialized as a copy of parsed data → `renderScheduleAndValidate` (calls `validateSchedule` then `renderTable` + `renderStatsAndChart`).

Editing a cell (`showCellDropdown`) writes into `currentSchedule` and, if the cell was originally prefilled, **also** back into `emp.originalSchedule` — this is the constraint the scheduler treats as immovable on the next run. Clearing a prefilled cell unlocks it (`isPrefilledFlags[d] = false`) so the algorithm can freely assign it.

`runAutoScheduling` (simulated annealing): first does a greedy pass per day to hit the exact daily headcount targets (`config.demand.day/evening/night`) using only cells that are `null` in `originalSchedule`, then repeatedly swaps two non-prefilled employees' shifts on a random day, accepting/rejecting via the Metropolis criterion (`Math.exp(-delta/temp)`), driven by `calculateScheduleCost`. Swapping (not reassigning) is what keeps the daily headcount invariant exactly satisfied throughout the search, so cost only needs to penalize per-employee constraints (min/max off days, max consecutive workdays, night-shift adjacency, shift-preference balance) — see the comment at the end of `calculateScheduleCost`.

`validateSchedule` is a second, independent implementation of the same constraints (used for live UI validation/highlighting rather than the search's cost function) — when adding or changing a scheduling rule, update the logic in **both** `calculateScheduleCost` and `validateSchedule`, plus the shift-code semantics in `isOff()` if the rule touches which codes count as a day off.

## Domain rules (see `.agents/skills/nurse_scheduling_rules/SKILL.md` for full detail)

These are load-bearing and easy to get subtly wrong — read the skill file before touching scheduling logic:

- Shift codes: `8`=white/day, `4`=evening, `12`=night, `OFF`=generated day off. `VV` counts as a day off; `SS` does **not**, and both are locked against the scheduler but remain hand-editable in the UI. Other prefilled strings (`W`, `FF`, `V`, `AV`, `VP`) are preserved but similarly locked.
- A day-1 cell containing patterns like `1大`/`3大`/`6白` is textual noise from the source sheet and must be ignored (treated as an empty, schedulable cell) — see `isIgnoredFirstDayText`.
- Hard per-employee limits: 8–9 off days/month (configurable), max 6 consecutive workdays (configurable), no night shift (`12`) immediately after `VV`/`SS`, no evening shift (`4`) immediately before `VV`/`SS`.
- Shift preference (`prefStr`, e.g. `白大`, `小`) restricts which shift codes an employee may be auto-assigned; when it lists multiple shift types the scheduler should keep counts across those types roughly balanced.
- A row is the head nurse (`isHeadNurse`) when `level === "N4"` and the preference field is blank — she is excluded from scheduling, validation, and daily headcount statistics entirely; her existing cells are left untouched.
- Global hard constraint: every day, non-head-nurse headcount must be exactly `config.demand.{day,evening,night}` (default 5/5/5) for each shift type.

## Excel export (`exportToExcel`)

Exports by mutating a clone of the original worksheet object rather than building a fresh one, so unrelated formatting/comments survive. Rules to preserve if you touch this function:
- Cell writes use `Object.assign({}, originalCell, {...})` specifically to keep the original cell's comment (`.c`) property intact.
- Comments are forced `hidden: true` and their text is manually wrapped every 15 characters, to prevent popup comment bubbles obscuring adjacent cells.
- Generated `OFF` is rewritten to `FF` on export.
- Prefilled cells get a `#FFFF00` fill; `4`/`12` shift text is red; everything is Arial 10pt, centered, with a `#D9D9D9` thin border.
- Column widths are hardcoded (not computed): A=5.44, B=9.56, C=7.11, D=6.56, E onward=4.44.
- Bottom summary rows (detected during parsing by matching cell values `"12"`/`"8"`/`"4"`/`"85"` in the first four columns) get their day counts recomputed from `currentSchedule` on export.

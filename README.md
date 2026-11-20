# AllocationLab

Internal workbench for designing, comparing and analysing **standard strategic and target asset allocations** for Financial Strategies' model profiles. It is a prototype. It covers the firm's standard profiles only and does not handle individual client portfolios.

> **Data notice.** All allocation data in this app is **synthetic** apart from one **illustrative** Balanced 60/40 allocation (Cash 2%, Fixed Interest 38%, Australasian Equities 17%, Global Equities 39%, Property 4%, Multi-Sector 0%). Nothing ships marked *Approved*. The app produces no investment recommendations or return forecasts.

## Quick start

Requires Node 18+ (developed on Node 22).

```bash
npm install
```

```bash
npm run dev
```

Open the URL Vite prints (normally http://localhost:5173).

| Command | What it does |
|---|---|
| `npm run dev` | Development server with hot reload |
| `npm test` | Run the unit tests (Vitest) |
| `npm run typecheck` | Strict TypeScript check |
| `npm run build` | Type-check and build to `dist/` |
| `npm run preview` | Serve the production build |

## Features

| Area | What you can do |
|---|---|
| **Dashboard** | See every profile's strategic allocation, provenance status, version, total, growth/defensive split and validation status, plus a stacked chart of all profiles. |
| **Allocation Designer** | Edit a scenario's weights and per-asset min/max limits. It shows live totals, the growth/defensive split, the Multi-Sector look-through setting, the difference against the strategic or target baseline, and validation messages. Edits stay drafts until you click Save. Revert and Reset-to-baseline are explicit actions. |
| **Profile Comparison** | Compare strategic, target and proposed (a saved scenario) allocations side by side. Shows pp differences for every pair, growth/defensive/unclassified rows, a stacked chart and a diverging pp-difference chart. |
| **Scenarios** | Create a scenario from a strategic baseline, a target baseline or a blank start. You can also duplicate, rename, delete and open scenarios, and compare up to 6 side by side. |
| **Assumptions** | Revise baselines, constraint sets, the Multi-Sector look-through fraction and validation tolerances. Every record carries a status, source and version. Each revision creates a new version and keeps the old one in history. A record can be marked *Approved* only through a typed confirmation, and approved records are locked. Includes an audit log and a reset to demo data. |
| **Import / Export** | Export baselines and/or scenarios as CSV, or download a template. Imports are parsed, validated and previewed before anything is applied (see below). |

### Data statuses

`approved` · `illustrative` · `synthetic` · `draft` · `missing`. Scenarios are always labelled **Hypothetical**.

## Calculation rules

All calculations are pure functions in `src/lib/` and are covered by unit tests.

- **Totals.** Weights are percentages. A total is valid when `|total − 100| ≤ tolerance`, with a default tolerance of **0.01 pp** (configurable). Floating-point noise is removed before comparison, so `0.1 + 0.2` counts as `0.3`.
- **Missing data.** A blank weight is `null`, which means *missing*. It is never treated as 0. Allocations with missing weights are marked **Incomplete**, and charts and tables label them explicitly.
- **Minimum Cash.** This is the Cash `min` in a profile's constraint set. It starts at **2%** for the illustrative Balanced 60/40 profile and is unset for the others, which produces a warning.
- **Min/max constraints.** These are configurable per asset class and per profile. Scenarios copy them when created, so later edits to a profile's limits don't change existing scenarios.
- **Growth/defensive.** Cash and Fixed Interest are defensive. Australasian Equities, Global Equities and Property are growth.
- **Multi-Sector.** Multi-Sector is split using a configurable look-through growth fraction (0–1). If no fraction is set, or the value is out of range, its weight is reported as **unclassified** and excluded from the split, with a warning. Scenarios can inherit the global fraction, override it, or force unclassified.
- **50/50 vs 60/40.** Profiles whose names state a split (Balanced 50/50 and Balanced 60/40) get a warning when the computed growth share differs from the named split by more than the nominal-split tolerance (default 0.5 pp).
- **No silent changes.** Nothing is ever rebalanced, normalised, clamped or rounded in storage. Validation only reports problems; it never fixes them.

## CSV format

```
record_type,profile,kind,name,status,source,version,cash,fixed_interest,australasian_equities,global_equities,property,multi_sector
scenario,balanced_60_40,proposed,Example proposal,hypothetical,Example,,2,38,17,39,4,0
baseline,balanced_50_50,target,,draft,Example source,,3,47,15,31,4,
```

- `record_type`: `baseline` or `scenario`.
- `profile`: one of `defensive`, `conservative`, `moderate`, `balanced_50_50`, `balanced_60_40`, `growth`, `high_growth`.
- `kind`: `strategic` or `target` for baselines; `proposed` for scenarios.
- Weights are percentages, and a trailing `%` is accepted. A blank cell means missing.

**Import safety rules** (`planImport` in `src/lib/csv.ts`):

1. A file with missing or duplicate columns is rejected outright. Invalid rows are listed with line numbers and skipped.
2. By default every valid row is imported as a **new hypothetical scenario**, so no baseline changes.
3. A baseline row can be applied as a **new baseline version** only if you select that action for the row *and* tick a confirmation box. The previous version is kept in history.
4. **Approved baselines can never be overwritten by import.** For those rows, only *import as scenario* or *skip* is offered.
5. An import cannot set a record to `approved`. Rows marked approved are recorded as `draft`.

## Persistence

State is saved to the browser's `localStorage` under the key `allocationlab:state:v1`. If the saved data is corrupt or unrecognised, the app shows the demo data instead and displays a notice; it does not delete the saved copy. **Assumptions → Reset demo data** restores the fixtures. To keep work, export it as CSV. There is no backend.

The seed data lives in `src/data/*.json`.

## Project structure

```
src/
  lib/            Pure, framework-free domain logic (+ *.test.ts)
    types.ts        Domain types
    reference.ts    Asset classes, profiles, colours
    allocation.ts   Totals, growth/defensive split, pp differences
    validation.ts   Allocation & constraint validation messages
    scenarios.ts    Immutable scenario operations (create/duplicate/save/revert/reset)
    csv.ts          CSV parse/serialise, import planning
    format.ts       Number parsing & formatting
  state/          App state: reducer, fixtures loader, persistence, selectors
  components/     Reusable UI (tables, badges, inputs, Recharts charts)
  pages/          Dashboard, Designer, Comparison, Scenarios, Assumptions, Import/Export
  data/           JSON fixtures (synthetic / illustrative / missing)
```

## Tests

`npm test` runs 46 tests covering:

- Allocation totals, tolerance and floating-point handling, and missing data
- The minimum Cash weight and other min/max constraints, including infeasible constraint sets
- The growth/defensive split and Balanced 50/50 vs 60/40
- Multi-Sector look-through: unset, set and out-of-range fractions
- Scenario isolation: deep-frozen baselines that scenario operations must not mutate; duplicates independent of their source; save/revert/reset semantics; and reducer-level checks that baselines are unchanged after scenario edits
- Baseline versioning, approval locking, and the rule that missing data cannot be approved
- CSV parsing, round-tripping, and the import safety rules

## Not in this iteration

- Monte Carlo, optimisation, return or risk forecasts (deliberately excluded)
- Authentication, user identity on approvals, multi-user storage or a backend
- Client portfolios
- CSV import and export of constraint sets and assumptions (only allocations are exchanged)
- Dark mode

## Reference spreadsheets

The repository also contains `Datasheet.xlsx` and `Recommendations Master Equity Weighting.xlsm`. These are the firm's existing working files. **AllocationLab does not read them, and none of their figures are copied into the fixtures.** Any future link to them should go through the CSV import path, so their values arrive as draft or hypothetical data and are never treated as automatically approved.

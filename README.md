# AllocationLab

Internal workbench for designing, comparing and analysing **standard strategic and target asset allocations** for Financial Strategies' model profiles, and for checking client holdings against them. It is a prototype.

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
| `npm run build:offline` | Build **one self-contained file**, `dist-offline/AllocationLab.html`, that opens straight from disk with no server or internet |
| `npm run golden` | Regenerate the recommendation-engine golden fixtures from the real workbook (needs LibreOffice) |

## Features

| Area | What you can do |
|---|---|
| **Dashboard** | See every profile's strategic allocation, provenance status, version, total, growth/defensive split and validation status, plus a stacked chart of all profiles. |
| **Allocation Designer** | Edit a scenario's weights and per-asset min/max limits. It shows live totals, the growth/defensive split, the Multi-Sector look-through setting, the difference against the strategic or target baseline, and validation messages. Edits stay drafts until you click Save. Revert and Reset-to-baseline are explicit actions. |
| **Profile Comparison** | Compare strategic, target and proposed (a saved scenario) allocations side by side. Shows pp differences for every pair, growth/defensive/unclassified rows, a stacked chart and a diverging pp-difference chart. |
| **Scenarios** | Create a scenario from a strategic baseline, a target baseline or a blank start. You can also duplicate, rename, delete and open scenarios, and compare up to 6 side by side. |
| **Assumptions** | Revise baselines, constraint sets, the Multi-Sector look-through fraction and validation tolerances. Every record carries a status, source and version. Each revision creates a new version and keeps the old one in history. A record can be marked *Approved* only through a typed confirmation, and approved records are locked. Includes an audit log and a reset to demo data. |
| **Client Holdings** | Load a client details file and a holdings file. For each client the page shows actual weights against the profile's strategic or target allocation, drift beyond tolerance, growth assets against the client's growth target, and whether cash covers monthly fees and withdrawals. Output is descriptive only. See [Client holdings](#client-holdings). |
| **Import / Export** | Export baselines and/or scenarios as CSV, or download a template. Imports are parsed, validated and previewed before anything is applied (see below). |

### Data statuses

`approved` · `illustrative` · `synthetic` · `draft` · `missing`. Scenarios are always labelled **Hypothetical**.

## Recommendations (the workbook as an app)

The **Recommendations** page runs the formulas of `Recommendations Master Equity Weighting.xlsm` for one client.

1. **Paste holdings.** Paste the Apex export. Names are matched to FS names, and term deposits become "Bank Term Deposit rate% dd/mm/yyyy".
2. **Set the client.** Choose the risk profile (Custom = any growth %), fees and withdrawals, money added or removed, and any overrides.
3. **Calculate.** You get per-holding recommendations, sector targets and the workbook's own error checks. You can then adjust overrides and export a CSV.

The **Implementation (rounds)** tab is the workbook's Implementation Advice table:
- record trades in up to six dated rounds, each with three sign-off initials;
- see the workbook's own "Remaining", plus broker, brokerage, spreads, rationale and the custodian's asset name;
- export each round's trade instructions;
- run the computer check against the Apex transaction report;
- save or open the implementation file (client data, so it is kept wherever you choose).

The **Letters & documents** tab fills the firm's Word templates the way the workbook's macros do:
- Letter of Advice, Email Advice and Follow-up letter;
- the Summary of changes, Portfolio for advice and Costs for advice tables, and the investment rationale with fact-sheet links;
- client details (name, address, client type, advisers) are typed in, because the macros read them from the Master Portfolio Risk Profiles workbook;
- choose the template .docx from the P: drive; it stays in the browser, and the filled letter is downloaded. Without a template you get a plain document with the tables;
- **Annual Review**: the Annual Review template filled with the client wording, signatories, signature table, the performance, sector-return and asset-allocation tables read from the pasted Apex or OneAnswer performance report, the commentary insert, and the advice tables. The figures read from the report are shown for checking before the document is made;
- **Statement of Advice appendices**: the product table (manager, website, description, broker, costs, fact sheet, PDS) for every held or recommended holding, in a landscape section of the Appendices template.

Firm-level inputs are editable on the **Settings** page and versioned there:
- profile percentages;
- the property model;
- fixed-interest, Australasian and global equity limits, splits and model fund lists.

How it works: the workbook's sheets are extracted into `src/engine/workbook/model.json` and evaluated by a TypeScript formula engine. The macros that change inputs are ported, and results are checked against the real workbook recalculated in LibreOffice: 110 golden scenarios across every section, with zero differences. Details, deviations and how to update the workbook are in [docs/recommendations-engine.md](docs/recommendations-engine.md).

| Command | What it does |
|---|---|
| `npm run extract:workbook -- "path/to/workbook.xlsm"` | Re-extract a newer version of the workbook |
| `npm run golden:full` | Regenerate the full-portfolio golden fixtures (needs LibreOffice; about 1 hour) |

Not yet ported:
- OneAnswer and MMC imports (and the OneAnswer transaction report);
- standing orders and regular purchases;
- client-master lookups.

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

## Client holdings

You can load three inputs (templates and synthetic samples are on the page):

1. **Investment datasheet** (optional but usually needed): `Datasheet.xlsx`, the *Data* sheet of the recommendations workbook, or a CSV export of it. It needs a header row with `Asset Class` and at least one of `Apex Name`, `FS Name` or `OA Name`. The app reads .xlsx/.xlsm files itself, with no library and without running macros.
2. **Client details:** a file with `Client number, Risk Profile, Growth Target %, Annual Fee % and/or Monthly Fees, Monthly Withdrawals`, plus optional `Cash Cover Months, Drift Tolerance pp, Notes`. You can also enter clients one at a time on the page.
3. **Holdings:** `Client number` (or `Portfolio Id`), `Asset Name`, `Asset Value` (or `Asset value NZD`). `Asset Class` and `Asset Sub-Class` are optional. Upload a CSV or tab-separated file, or paste straight from Excel.

**How each holding gets its asset class.** The first rule that applies wins:

1. A manual assignment you made for that investment name.
2. The Asset Class column in the holdings file.
3. The datasheet: an exact name match on Apex, FS or OA name, ignoring case and spacing. If that fails, it tries again after removing the trailing detail suffix that platforms add to term deposits and bonds, such as `(188,000.00; 3.53%; 18 Months; 24/12/2026)`.

Holdings that still have no class, or match datasheet rows with conflicting classes, are listed for you to assign. Until then they are excluded from the class weights, with a warning. The app never fuzzy-matches names.

Calculation rules (`src/lib/holdings.ts` and `src/lib/datasheet.ts`, with tests):

- **Weights.** Weight = class value ÷ the client's total value. Unclassified or unmapped value stays in the total and is reported separately.
- **Drift.** Drift = actual weight − the profile's strategic or target weight, flagged beyond ±tolerance. The default tolerance is 2 pp, matching the existing workbook; a per-client value overrides it.
- **Growth target.** The actual growth share is compared with the client's growth target. If the client's target differs from the profile model's growth share, that is noted. The app does not build a client-specific allocation.
- **Fees.** Monthly fees = (annual fee % × total portfolio value ÷ 12) + any fixed monthly fee.
- **Cash coverage.** Required cash = (monthly fees + monthly withdrawals) × cover months. The default is 12 months; a per-client value overrides it.
- **Bad data.** Rows with missing or invalid values, unknown profiles and duplicate clients are rejected with line numbers. Negative values are kept but warned about.

**Data handling.** Client files, manually entered clients and the datasheet are held only in the open tab's memory. They are cleared on reload or by **Clear client data**, and client numbers are never put in the URL. The only things remembered in browser storage are the asset-class label mapping, manual class assignments keyed by investment name (no client numbers or values), and the default settings. The prototype has no authentication, so do not load real client data until the firm's privacy and compliance process has approved it.

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

`npm test` runs 356 tests covering:

- Allocation totals, tolerance and floating-point handling, and missing data
- The minimum Cash weight and other min/max constraints, including infeasible constraint sets
- The growth/defensive split and Balanced 50/50 vs 60/40
- Multi-Sector look-through: unset, set and out-of-range fractions
- Scenario isolation: deep-frozen baselines that scenario operations must not mutate; duplicates independent of their source; save/revert/reset semantics; and reducer-level checks that baselines are unchanged after scenario edits
- Baseline versioning, approval locking, and the rule that missing data cannot be approved
- CSV parsing, round-tripping, and the import safety rules
- Client holdings: money parsing, CSV/TSV validation, class mapping, weights, drift and tolerance overrides, growth target gaps, annual-% fees and cash coverage
- Implementation rounds (12 golden scenarios: Remaining, broker/spread/rationale, exposure, checks) and the computer check
- Letter tables (Excel number formats, sort order, totals that add up, buy/sell ordering) and the Word filler on a synthetic template: text split across runs, whole-word and case matching, trust wording, tables, bullets, hyperlinks, signature bookmarks and removed cost sections
- Annual review and SoA: Apex and OneAnswer report parsing (synthetic reports), the annual review template tables, wording and highlights, and the SoA landscape section
- The recommendations workbook: 110 golden scenarios against the real workbook, Excel-semantics unit tests, and a check that the evaluator reproduces Excel's saved results
- The .xlsx reader and datasheet matching, using a synthetic workbook fixture: Apex/FS/OA names, term-deposit suffix stripping, ambiguous and not-found names, and classification precedence

## Not in this iteration

- Monte Carlo, optimisation, return or risk forecasts (deliberately excluded)
- Authentication, user identity on approvals, multi-user storage or a backend
- Client-specific target allocations, trade or rebalancing suggestions, and exporting client results
- CSV import and export of constraint sets and assumptions (only allocations are exchanged)
- Dark mode

## Reference spreadsheets

The repository also contains `Datasheet.xlsx` and `Recommendations Master Equity Weighting.xlsm`. These are the firm's existing working files. **AllocationLab does not read them, and none of their figures are copied into the fixtures.** Any future link to them should go through the CSV import path, so their values arrive as draft or hypothetical data and are never treated as automatically approved.

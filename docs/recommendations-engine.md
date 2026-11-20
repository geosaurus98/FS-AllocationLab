# Recommendations engine — porting the workbook

Goal: an app version of `Recommendations Master Equity Weighting.xlsm` that produces **the same numbers as the workbook**, delivered as a single offline HTML file (`npm run build:offline`).

Decisions so far:
- The core engine comes first.
- It must match the workbook exactly, verified by automated golden tests.
- Advisers run it as an offline single file.
- It uses the workbook's own risk profiles.

## How the workbook is put together

| Part | Size | Role |
|---|---|---|
| `Portfolio` sheet | 3,152 formulas | Holdings by section (rows 3–205); current %, target %, change, recommended value, overrides; check tables (BC–BN) |
| `Asset Allocation` sheet | 706 formulas | Tactical profile table (`Targets`), fixed-interest ladder, Australasian/Global/Property fund distributions, limits |
| `Data` sheet | 492 investments | Investment datasheet: names (OA/Apex/FS), asset class, type, brokerage, rationale |
| `Balanced` sheet | 657 formulas | Industry comparison and the model Balanced portfolio (reference) |
| VBA | ~20,000 lines, 100 modules | Custodian imports, model-asset insertion, implementation rounds, trade stats, Word documents, Outlook emails |

### Portfolio sheet layout

| Section | Holding rows | Total row | Target cell |
|---|---|---|---|
| Cash | 4–10 | 11 | `G11` |
| Fixed Interest | 13–71 | 72 | `G72` |
| Australasian Equities | 74–119 | 120 | `G120` |
| Global Equities | 122–167 | 168 | `G168` |
| Property | 170–191 | 192 | `G192` |
| Multi Sector | 194–200 | 201 | `G201` |

The columns are:
- **A** — name (FS name)
- **B** — current value
- **C** — current %
- **D** — % target (total rows only)
- **E** — change
- **F** — recommended value
- **G** — override (holding rows) or sector target (total rows)
- **H** — recommended %
- **J** — type/style from the datasheet

### Inputs (per client)

- **Holdings** in A/B, filled by the custodian import macros.
- **`Portfolio_Total` (B204)**, set by import to the holdings total.
- **`Risk_Profile` (BE1).** Selecting a profile runs `ComboBox1_Change`, which writes the profile's percentages × 100 into D11, D72, D120, D168, D192 and D201.
- **Money settings:**
  - `Fees_and_Withdrawals` (BH2)
  - `Add_remove_to_Portfolio` (BD11)
  - `Cash_Total_Override` (BD12)
  - Sector total overrides BD72, BD120, BD168 and BD192
- **Per-holding overrides** in column G.
- **Custom profile settings:** Asset Allocation H3:H6.
- **Firm settings on the Asset Allocation sheet:**
  - Distributions and model fund lists
  - Minimum trade limits (fund $250, shares $2,500–5,000, bonds $5,000, TDs $1,000/$10,000)
  - Bank exposure cap (10%)
  - Managed-fund share of fixed interest (30%)

## Profiles (tactical table)

| | Conservative | Moderate | Balanced | Growth | Aggressive | Custom |
|---|---|---|---|---|---|---|
| Lower risk | 75% | 60% | 40% | 20% | 10% | adviser |
| Cash | 7.5 | 5 | 2 | 2 | 2 | adviser |
| Fixed income | 67.5 | 55 | 38 | 18 | 8 | adviser |
| Growth | 25 | 40 | 60 | 80 | 90 | adviser |
| Aus. equities | derived | derived | 16 | derived | derived | derived |
| Global equities | derived | derived | 40 | derived | derived | derived |
| Property | derived | 4 | 4 | derived | derived | derived |

How the derived cells and the extra profiles work:
- **Property (derived columns).** `Balanced property ÷ Balanced growth × growth`.
- **Australasian and global equities (derived columns).** `Balanced AE (or GE) ÷ (Balanced growth − Balanced property) × (growth − property)`.
- **Custom property.** The Custom property cell refers to itself and is solved by Excel's iterative calculation. Its fixed point equals the rule above, so the engine uses the closed form.
- **Current.** Uses the client's current sector percentages.
  - Quirk: when the client has no cash, the current cash cell is blank text. That makes the Current profile's lower risk and cash 0 (`IFERROR` of `"" + n`).
- **Master Risk Alloc.** Uses values the `FetchClientInfo` macro fetches from the client master file.

## Sector targets (implemented: `src/engine/sectorTargets.ts`)

All six are ported formula for formula, keeping the workbook's mix of `=""` and `=0` blank tests (`src/engine/excel.ts`).

- **Cash (G11).** If a cash total override above 0 is set, use it. Otherwise:
  - Base = profile cash ÷ lower risk × (FI% + cash%) × portfolio total.
  - Result = the larger of that base and fees & withdrawals.
- **Fixed interest (G72).** Rules in priority order:
  1. If a sector override is set: current + override.
  2. If a cash total override is set: total − recommended (AE + GE + property) − cash override + add/remove. This is the balancing item, so it depends on per-holding results.
  3. If add/remove is set: the same rule as 4, scaled on (total + add/remove).
  4. Otherwise: FI's share of defensive, minus any cash lifted above normal to cover fees.
- **Australasian (G120), global (G168) and property (G192).**
  - If a sector override is set: current + override.
  - Otherwise: profile sector ÷ growth × (100 − FI% − cash%) × total, scaled on (total + add/remove) when add/remove is set.
- **Multi sector (G201).** MS target% × final portfolio total.

## Golden tests

```
scripts/golden/gen_sector_targets.py   # seeded synthetic scenarios
scripts/golden/run_workbook.py         # fills a copy of the workbook, recalculates in LibreOffice (macros never run)
scripts/golden/trim_fixture.py         # → src/engine/__golden__/sector-targets.json
src/engine/sectorTargets.golden.test.ts
```

To regenerate the fixtures:

```bash
npm run golden
```

This needs LibreOffice with Python UNO.

A recalculation check of the unmodified workbook gave **0 differences** between LibreOffice and Excel's saved results across 3,091 formula cells. Every holding cell in that template was empty, though; the golden scenarios add holdings.

## Phases

1. **Profiles and sector targets** — done, golden-tested.
2. **Property, multi-sector and cash rows** — done, golden-tested (40 scenarios covering all five property bands, purchases, sales, under-limit holds and overrides). The "Update Sheet" model-asset insertion is ported for property.
3. **Remaining per-holding sections.**
   - Port the `Recalculate` macro: it adds model funds with a distribution to each section with value 0.
   - Then cash, multi-sector and property row formulas, followed by global equities (active/passive, gold, hedging), then Australasian equities (NZ/AUS split, passive/active, direct shares).
3. **Fixed interest.** Managed-fund share, the bond and term-deposit ladder by maturity year, perpetual and convertible bonds, maturity versus hold date, rounding to $100, minimum bond/TD limits and bank exposure.
4. **Checks and outputs.** Error checks (BD2), exposure table, top 10, fixed-interest breakdown, and the recommended-portfolio and changes view.
5. **Later (not yet scheduled).** Custodian imports (Apex/OA/MMC), implementation rounds and trade stats, documents (Letter of Advice, SoA, annual review) and emails.

## Known quirks to preserve (and review with the firm)

- **Blank cash target.** `Cash_Target` blank versus 0 changes results: some formulas test `=""` and others `=0`.
- **Current profile with no cash.** As described above, a client with no cash gets 0 for lower risk and cash.
- **Dates.** Bond and TD ladder years use `TODAY()`, so results depend on the date the workbook is opened. The engine will take an explicit "as at" date.
- **Multi-sector rows (likely a workbook bug).** Each row's formula is shifted by one row. The first multi-sector holding gets the target divided by `COUNTIF(B193:B200,"<>0")`, which also counts the empty rows, so normally the target ÷ 8. Every other multi-sector holding gets $0. All shipped profiles have 0% multi-sector, so this only matters when an adviser sets a multi-sector target.
- **Property: flat AUS/INT rows.** Extra AUS and international rows (beyond the first) ignore the bands and only apply when the pot is at least $30,000, a value hard-coded in the formula.
- **Property: kept holdings.** A holding with override 0 whose name is not anywhere in Asset Allocation column B is deducted from the property pot.
- **Hard-coded fund name.** The cash-sufficiency check refers to "Nikko AM NZ Cash Fund" by name; the fund has since been renamed Amova.

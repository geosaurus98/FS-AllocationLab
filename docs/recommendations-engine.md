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

## How the app runs the workbook

The app evaluates **the workbook's own formulas**. It doesn't use hand-written approximations of them.

1. **Extraction.** `scripts/engine/extract_workbook.py` reads four sheets from the `.xlsm` into `src/engine/workbook/model.json`: Portfolio, Asset Allocation, Data and ValidationLists. It keeps every value, every formula (with its array-formula flag), the named ranges and the iterative-calculation setting. Macros are not read.
2. **Evaluation.** `src/engine/workbook/` is a TypeScript formula evaluator:
   - **Parser:** `parser.ts`.
   - **Excel value and comparison rules:** `values.ts`.
   - **Evaluator:** `evaluator.ts` implements the ~45 functions the workbook uses. It follows legacy Excel's implicit intersection, with element-by-element behaviour in array formulas and SUMPRODUCT. Circular references are solved by iteration.
3. **Macro behaviour that changes inputs** is ported in `scenario.ts`, `importHoldings.ts` and `settingsCells.ts`:
   - **Apex import:** names → FS names, term deposits as "Bank Term Deposit rate% dd/mm/yyyy", holdings merged and placed by asset class.
   - **AdjustPortfolio:** Data rows flagged "Add" are inserted.
   - **Gold_distribution.**
   - **ResetInvestmentsMaturing:** 90-day hold window.
   - **Profile selection (ComboBox1_Change).**
   - **Asset Allocation `Worksheet_Calculate`:** inserts bond/TD ladder rows, property, global and Australasian model rows. It is repeated until nothing more is added.
   - **"Update Sheet" button:** optional.
4. **Settings.** Firm settings (Settings page) are written into the Asset Allocation input cells before each run. Cells that hold formulas are never overwritten.

**Updating to a newer workbook:**

```bash
python3 scripts/engine/extract_workbook.py "path/to/Recommendations Master Equity Weighting.xlsm"
```

Then re-run the golden scenarios (below) and the unit tests.

The hand-written modules from the first phases (`sectorTargets.ts`, `property.ts`, `cashMs.ts`, `profiles.ts`) are kept as readable reference implementations. They are tested against the same golden data.

## Verification

- **Excel's saved results.** On the template, the evaluator reproduces every value Excel last saved on the Portfolio and Asset Allocation sheets (`evaluator.test.ts`).
- **LibreOffice golden scenarios.** `scripts/golden/run_workbook.py` fills a copy of the real workbook with a synthetic scenario and recalculates it in LibreOffice, headless and with macros never run. It applies the same macro emulation and the en-NZ locale, then records every holding row. The evaluator must match all of them.
  - **Sector targets:** 40 scenarios.
  - **Property, multi-sector and cash:** 40 scenarios.
  - **Full portfolios:** 30 scenarios, all sections, with term deposits around the 90-day maturity window, dated bonds, NZ/AUS shares, equivalents, gold and overrides.
- **Excel semantics.** Unit tests cover the behaviours the workbook depends on (`functions.test.ts`).

**LibreOffice difference.** LibreOffice 24.2 doesn't implement Excel 365's `FILTER`, which the global/Australasian "equivalents" text uses. It returns `#NAME?` there. Golden comparisons therefore run the evaluator with `libreOfficeCompat = true`. The app always uses Excel behaviour, so `FILTER` works as in Excel 365.

**Macro emulation, not tested against Excel.** The macro steps are ported from the VBA source and checked by review. In Excel, the insertion event fires on every recalculation, in an order that depends on how cells were edited. The app runs it after import and profile selection until nothing changes. In rare cases Excel may therefore leave extra $0 placeholder rows that the app does not.

## Phases

1. **Profiles and sector targets** — done.
2. **Property, multi-sector and cash** — done.
3. **Workbook evaluator** — done. It covers global and Australasian equities, fixed interest and checks.
4. **Recommendations screen** — done:
   - paste the Apex holdings;
   - set profile, Custom growth %, fees and withdrawals, add/remove and cash override;
   - get per-holding recommendations, sector summary, workbook checks, overrides and a CSV export.
5. **Implementation (rounds)** — done. This is the workbook's Implementation Advice table:
   - **Rounds 1–6:** the amounts go into the Portfolio round columns (AC, AG, AK, AO, AS, AW), each with a date and three sign-off initials.
   - **Remaining:** the workbook's own BA formula. Round amounts on actual bonds count against the matching "YYYY Bonds (BBB+…)" placeholder rows.
   - **Holding details:** broker and brokerage by custodian, spreads, rationale and the custodian's asset name (SafeNameChange).
   - **Bank exposure and checks** are recalculated with the rounds.
   - **Trade instructions per round:** sells, then buys, as a CSV.
   - **Computer check** (Transactions_Match) against the Apex transaction report.
   - **Implementation file:** saved where the adviser chooses, never in the browser.
   - **Verification:** 12 golden round scenarios in LibreOffice match exactly, including Remaining, broker/spread/rationale columns, exposure and checks.
   - **Speed:** round edits recalculate only Portfolio AC:BN, which are the only cells that depend on them.
6. **Letters & documents** — done (Letter of Advice, Email Advice, Follow-up letter):
   - **Tables** (`src/engine/workbook/documentTables.ts`): ports of Sort_Portfolio, Portfolio_For_Advice, Costs_For_Advice, InvestmentRationale and Simple_AR_table, with the Excel number formats of the cells that are pasted.
   - **Word filler** (`src/lib/docx.ts`, `src/lib/letters.ts`): ReplaceTextInWord (whole words, case-insensitive with Word's case matching, across runs), the tables at their placeholders with the widths and font sizes the macros set, rationale headings and bullets with "Link to the fact sheet here.", the AdvisorTable signature table built from the adviser bookmarks, which are then deleted, and Cost1–Cost3 removed when there are no costs.
   - **Placeholders:** Letter of Advice and Follow-up use `Investmentrationale`, `Portfolioforadvice`, `InsertTable`, `ChangeTable`; Email Advice uses `Portfolio_for_advice` and `Costs_For_Advice` instead, plus `holder0`–`holder4` wording by client type.
   - **Inputs the app cannot reach:** the templates (P: drive) are chosen by the adviser; client details from the Master Portfolio Risk Profiles workbook are typed in. Nothing is uploaded.
   - **Not ported:** the "hold the portfolio at the current date" prompt and the save-to-P: step (the browser downloads the file instead).
7. **Later (not yet scheduled):**
   - OneAnswer and MMC imports;
   - the OneAnswer transaction report for the computer check;
   - standing orders and regular purchases;
   - the implementation dashboard (bond summary tables);
   - more than six rounds (the workbook inserts extra columns);
   - the Statement of Advice and annual review documents (they need performance, tax and holder details);
   - the external "Bonds Allocation in Years" file used by the Bond_Allocation macro;
   - client master lookups (FetchClientInfo).

## Known quirks to preserve (and review with the firm)

- **Blank cash target.** `Cash_Target` blank versus 0 changes results: some formulas test `=""` and others `=0`.
- **Current profile with no cash.** As described above, a client with no cash gets 0 for lower risk and cash.
- **Dates.** Bond and TD ladder years use `TODAY()`, so results depend on the date the workbook is opened. The engine will take an explicit "as at" date.
- **Multi-sector rows (likely a workbook bug).** Each row's formula is shifted by one row. The first multi-sector holding gets the target divided by `COUNTIF(B193:B200,"<>0")`, which also counts the empty rows, so normally the target ÷ 8. Every other multi-sector holding gets $0. All shipped profiles have 0% multi-sector, so this only matters when an adviser sets a multi-sector target.
- **Property: flat AUS/INT rows.** Extra AUS and international rows (beyond the first) ignore the bands and only apply when the pot is at least $30,000, a value hard-coded in the formula.
- **Property: kept holdings.** A holding with override 0 whose name is not anywhere in Asset Allocation column B is deducted from the property pot.
- **Datasheet entries to check.** "ANZ Bank New Zealand Limited (ANB190) 4.63% 20/02/2030" is typed as a Term Deposit. Apex code WNZ-0726 maps to an FS name with a different coupon (6.355% versus 3.455%).
- **Computer check (deliberate differences).** The macro looks for "Bond" in the Data sheet's *Style* column, which never contains it, so bond names were never trimmed and bonds never matched; the app uses the *Type* column. The macro also kept only the last transaction per asset; the app adds them together.
- **Letter tables (deliberate differences).** In the Summary table the macro never resets its running section total after Cash, so the first "All Other … (No Changes)" figure would subtract cash funds; the app resets it. It treats any 3-letter name as a cash currency in every section; the app does so only in Cash. A bond or term deposit whose name does not end in a date stops the macro with an error; the app shows "Sell". The stray column-T formula on the "Multi Sector Total" row would add a rationale bullet for the total; the app ignores it. The macro means to remove the Cost1–Cost3 sections when there are no costs, but its test can never be true; the app removes them. Column U (Fact Sheet) calls `GetHyperlinkAddress`, which is not defined in the workbook (the module defines `GetlinkAddress`), so Excel shows it blank; the letters use the Data sheet's FactSheet value, as the macro does.
- **Hard-coded fund name.** The cash-sufficiency check refers to "Nikko AM NZ Cash Fund" by name; the fund has since been renamed Amova.

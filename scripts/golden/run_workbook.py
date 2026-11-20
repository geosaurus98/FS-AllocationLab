"""
Golden-result generator for the recommendation engine.

Fills a COPY of "Recommendations Master Equity Weighting.xlsm" with a scenario, recalculates it in
LibreOffice (headless, macros never executed), and records the workbook's results. The TypeScript
engine is then tested against these results (src/engine/*.golden.test.ts).

Usage:
    python3 scripts/golden/run_workbook.py scenarios.json out.json

Requires LibreOffice with Python UNO (`soffice`, `import uno`). Scenarios must contain synthetic data only.

Scenario shape (all money in NZD, percentages as in the workbook):
{
  "id": "s1",
  "profile": "Balanced",                 # Targets header: Conservative..Custom, Current, Master Risk Alloc
  "custom": {"lowerRisk":0.3,"cash":0.02,"fi":0.28,"growth":0.7},   # optional, Asset Allocation H3:H6
  "targetsPct": {"cash":2,"fi":38,"ae":16,"ge":40,"prop":4,"ms":0},  # optional; default = profile x100 (ComboBox behaviour)
  "holdings": {"cash":[["NZD Call Account",1000]], "fi":[...], "ae":[...], "ge":[...], "prop":[...], "ms":[...]},
  "portfolioTotal": null,                # default: sum of holdings (as the import macros set it)
  "feesAndWithdrawals": 0,               # Portfolio!BH2 (annual $)
  "addRemove": null,                     # Portfolio!BD11
  "cashTotalOverride": null,             # Portfolio!BD12
  "sectorOverrides": {"fi":null,"ae":null,"ge":null,"prop":null},   # BD72, BD120, BD168, BD192
  "portfolioDate": "2026-11-20"
}
"""
import json, os, shutil, sys, tempfile
sys.path.insert(0, os.path.dirname(__file__))
from lo import LO

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
WORKBOOK = os.path.join(ROOT, 'Recommendations Master Equity Weighting.xlsm')

# Portfolio sheet sections: (key, first holding row, last holding row, total row)
SECTIONS = [('cash', 4, 10, 11), ('fi', 13, 71, 72), ('ae', 74, 119, 120), ('ge', 122, 167, 168), ('prop', 170, 191, 192), ('ms', 194, 200, 201)]
TARGET_ROWS = {'cash': 3, 'fi': 4, 'ae': 6, 'ge': 7, 'prop': 8, 'ms': 9}  # rows within Targets (1-based, header = 1)
TARGET_CELLS = {'cash': 'D11', 'fi': 'D72', 'ae': 'D120', 'ge': 'D168', 'prop': 'D192', 'ms': 'D201'}
SECTOR_OVERRIDE_CELLS = {'fi': 'BD72', 'ae': 'BD120', 'ge': 'BD168', 'prop': 'BD192'}


def addr(sheet, a):
    return sheet.getCellRangeByName(a)


def put(sheet, a, v):
    c = addr(sheet, a)
    if v is None:
        c.setFormula('')
    elif isinstance(v, str):
        c.setString(v)
    else:
        c.setValue(float(v))


def val(sheet, a):
    c = addr(sheet, a)
    t = c.getType().value  # EMPTY, VALUE, TEXT, FORMULA
    if t == 'EMPTY':
        return None
    if t == 'FORMULA':
        err = c.getError()
        if err:
            return {'error': c.getString() or f'Err:{err}'}
        # com.sun.star.sheet.FormulaResult: VALUE=1, STRING=2, ERROR=4
        if c.FormulaResultType2 == 2:
            return c.getString()
        return c.getValue()
    if t == 'TEXT':
        return c.getString()
    return c.getValue()


def settle(lo, aa, rounds=60):
    """Recalculate until the Custom profile's self-referencing property cell (H9) stops changing.

    The workbook relies on Excel's iterative calculation for H9. Each Excel recalculation continues
    from the previous value, so a workbook in use sits at the converged fixed point; one LibreOffice
    pass may stop short because of the workbook's 0.001 max-change setting.
    """
    prev = None
    for _ in range(rounds):
        lo.doc.calculateAll()
        v = aa.getCellRangeByName('H9').getValue()
        if prev is not None and abs(v - prev) < 1e-14:
            return
        prev = v


# Named ranges scanned by Trigger_Recalculate.Recalculate and the Portfolio section each feeds.
MODEL_RANGES = [('ae', ['B35:B70']), ('ge', ['B79:B114', 'J73:J77']), ('prop', ['B119:B149'])]


def recalculate_macro(p, aa):
    """Port of the Recalculate macro: every model asset whose amount (2 columns right) is non-zero,
    whose name has more than one word and does not contain "Direct", is added to its section with
    value 0 if it is not already there."""
    for key, ranges in MODEL_RANGES:
        first, last = next((f, l) for k, f, l, _ in SECTIONS if k == key)
        for rng in ranges:
            cells = aa.getCellRangeByName(rng)
            for i in range(cells.Rows.Count):
                c = cells.getCellByPosition(0, i)
                name = c.getString()
                if not name or 'Direct' in name:
                    continue
                if len(name.strip().split(' ')) <= 1:
                    continue
                amt = aa.getCellByPosition(c.CellAddress.Column + 2, c.CellAddress.Row)
                # Empty <> 0 is False in VBA; an error value would stop the macro (type mismatch) —
                # treated here as "not added".
                if amt.getError() != 0 or amt.getValue() == 0:
                    continue
                name = name.strip()
                empty_row = None
                exists = False
                for r in range(first, last + 1):
                    a = p.getCellRangeByName(f'A{r}').getString()
                    if a == name:
                        exists = True
                        break
                    if a == '':
                        empty_row = r
                        break
                if not exists and empty_row is not None:
                    put(p, f'A{empty_row}', name)
                    put(p, f'B{empty_row}', 0)


def _insert(lo, p, key, name):
    """UpdateDescriptions / UpdateGlobal / UpdateProperty / UpdateAus: add name to the section if absent."""
    first, last = next((f, l) for k, f, l, _ in SECTIONS if k == key)
    for r in range(first, last + 1):
        if p.getCellRangeByName(f'A{r}').getString() == name:
            return False
    for r in range(first, last + 1):
        c = p.getCellRangeByName(f'A{r}')
        if c.getType().value == 'EMPTY':
            put(p, f'A{r}', name)
            put(p, f'B{r}', 0)
            lo.doc.calculate()
            return True
    return False


class _Abort(Exception):
    pass


def _num(cell):
    """VBA Value2 of a cell for numeric comparison; errors/text raise (type mismatch aborts the event)."""
    if cell.getError() != 0:
        raise _Abort()
    t = cell.getType().value
    if t == 'EMPTY':
        return 0.0
    if t == 'TEXT' or (t == 'FORMULA' and cell.FormulaResultType2 == 2):
        s = cell.getString()
        if s == '':
            return ''
        raise _Abort()
    return cell.getValue()


def worksheet_calculate_insert(lo, p, aa):
    """Port of Asset Allocation Worksheet_Calculate (sheet2.cls): one pass, live recalculation.
    Returns True if anything was inserted."""
    inserted = False
    try:
        # ProcessBondsAndTDs: BondsTDs = H20:I27
        for r in range(20, 28):
            for col, label in (('H', 'bonds'), ('I', 'tds')):
                c = aa.getCellRangeByName(f'{col}{r}')
                if c.getError() != 0:
                    continue
                v = _num(c)
                if v == '' or v == 0:
                    continue
                year = aa.getCellRangeByName(f'B{r}').getString()
                desc = f'{year} Bonds (BBB+ or higher, subject to liquidity)' if label == 'bonds' else f'Term Deposits {year}'
                inserted |= _insert(lo, p, 'fi', desc)
        # ProcessProperties: Properties = B119:D149
        for r in range(119, 150):
            amt = aa.getCellRangeByName(f'D{r}')
            if amt.getError() != 0:
                continue
            name = aa.getCellRangeByName(f'B{r}').getString()
            if name and _num(amt) != '' and _num(amt) > 0:
                inserted |= _insert(lo, p, 'prop', name)
        # ProcessGlobalEquities: Global_Equities = B79:G114
        limit = aa.getCellRangeByName('Q75').getValue()
        for r in range(79, 115):
            d = aa.getCellRangeByName(f'D{r}')
            if d.getError() != 0:
                continue
            g = aa.getCellRangeByName(f'G{r}')
            dv = _num(d)
            if dv == '' or not dv > 0:
                continue
            gv = _num(g)
            if gv == '':
                continue
            mv = _num(aa.getCellRangeByName(f'M{r}'))
            if (0 if mv == '' else mv) < 2 and gv > limit:
                name = aa.getCellRangeByName(f'B{r}').getString()
                if name not in ('Active', 'Direct', 'Passive'):
                    inserted |= _insert(lo, p, 'ge', name)
        # ProcessAustralasianEquities: AA_Australasian_Equities = B35:B70
        for r in range(35, 71):
            txt = aa.getCellRangeByName(f'B{r}').getString()
            if txt and 'direct' not in txt.lower() and len(txt.strip().split(' ')) > 1:
                dv = _num(aa.getCellRangeByName(f'D{r}'))
                if dv != '' and dv != 0:
                    inserted |= _insert(lo, p, 'ae', txt)
    except _Abort:
        pass  # a run-time error ends the event handler (its error message box is not reproduced)
    return inserted


def import_macros(lo, p, aa):
    """Steps of PortfolioSetup that change inputs: AdjustPortfolio, Gold_distribution, ResetInvestmentsMaturing."""
    data = lo.sheet('Data')
    # AdjustPortfolio: Data rows flagged "Add" (column W) are added to their section with value 0.
    last_data = max(r for r in range(1, 1200) if data.getCellByPosition(22, r - 1).getString() != '' or r == 1)
    def last_portfolio_row():
        return max(r for r in range(1, 260) if p.getCellRangeByName(f'A{r}').getString() != '')
    for i in range(2, last_data + 1):
        if data.getCellByPosition(22, i - 1).getString() != 'Add':
            continue
        name = data.getCellByPosition(3, i - 1).getString()
        section = data.getCellByPosition(4, i - 1).getString()
        lastp = last_portfolio_row()
        header = next((r for r in range(1, lastp + 1) if p.getCellRangeByName(f'A{r}').getString().lower() == section.lower()), None)
        if header is None:
            continue
        if any(p.getCellRangeByName(f'A{r}').getString().lower() == name.lower() for r in range(header, lastp + 1)):
            continue
        r = header
        while p.getCellRangeByName(f'A{r}').getString() != '':
            r += 1
        put(p, f'A{r}', name)
        put(p, f'B{r}', 0)
    # Gold_distribution: K74:K77 from the gold assets held.
    gold = [aa.getCellRangeByName(f'J{r}').getString() for r in range(74, 78)]
    held = {}
    total = 0.0
    for r in range(1, last_portfolio_row() + 1):
        a = p.getCellRangeByName(f'A{r}').getString()
        if a and a in gold and a not in held:
            held[a] = p.getCellRangeByName(f'B{r}').getValue()
            total += held[a]
    for i, g in enumerate(gold):
        if g in held:
            v = held[g] / total if total != 0 else (1 if g == 'Sprott Physical Gold Trust ETV' else 0)
        else:
            v = 0
        put(aa, f'K{74 + i}', v)
    # ResetInvestmentsMaturing
    put(p, 'BD14', 90)


def run(lo, sc):
    p = lo.sheet('Portfolio')
    aa = lo.sheet('Asset Allocation')
    profile = sc['profile']
    if sc.get('custom'):
        cu = sc['custom']
        for cell, k in (('H3', 'lowerRisk'), ('H4', 'cash'), ('H5', 'fi'), ('H6', 'growth')):
            put(aa, cell, cu[k])
    # 1. Import: holdings, portfolio total and money settings (what the custodian import macros do).
    total = 0.0
    for key, first, last, _ in SECTIONS:
        rows = sc.get('holdings', {}).get(key, [])
        if len(rows) > last - first + 1:
            raise ValueError(f'too many {key} holdings')
        for i in range(last - first + 1):
            r = first + i
            if i < len(rows):
                name, value = rows[i][0], rows[i][1]
                put(p, f'A{r}', name)
                put(p, f'B{r}', value)
                put(p, f'G{r}', rows[i][2] if len(rows[i]) > 2 else None)
                total += float(value)
            else:
                put(p, f'A{r}', None); put(p, f'B{r}', None); put(p, f'G{r}', None)
    put(p, 'B204', sc['portfolioTotal'] if sc.get('portfolioTotal') is not None else total)
    put(p, 'BH2', sc.get('feesAndWithdrawals', 0))
    put(p, 'BD11', sc.get('addRemove'))
    put(p, 'BD12', sc.get('cashTotalOverride'))
    for k, cell in SECTOR_OVERRIDE_CELLS.items():
        put(p, cell, (sc.get('sectorOverrides') or {}).get(k))
    put(p, 'B2', sc.get('portfolioDate', '2026-11-20'))
    if sc.get('importMacros', True):
        import_macros(lo, p, aa)
    settle(lo, aa)
    # 2. Select the risk profile: ComboBox1_Change copies the profile column x100 into column D.
    headers = [aa.getCellByPosition(c, 1).getString() for c in range(2, 10)]
    col = 2 + headers.index(profile)
    targets = {k: aa.getCellByPosition(col, r).getValue() for k, r in TARGET_ROWS.items()}
    put(p, 'BE1', profile)
    tp = sc.get('targetsPct') or {}
    for k, cell in TARGET_CELLS.items():
        put(p, cell, tp[k] if k in tp else targets[k] * 100)
    settle(lo, aa)
    # 3. Asset Allocation Worksheet_Calculate event: add model rows, repeated until nothing changes.
    if sc.get('insertModelAssets', True):
        for _ in range(10):
            if not worksheet_calculate_insert(lo, p, aa):
                break
            settle(lo, aa)
    # 4. Optional "Update Sheet" button (Trigger_Recalculate.Recalculate).
    if sc.get('updateSheetButton'):
        recalculate_macro(p, aa)
        settle(lo, aa)
    out = {'id': sc['id'], 'targetsUsed': targets, 'sectors': {}, 'rows': {}}
    for key, first, last, tot in SECTIONS:
        out['sectors'][key] = {c: val(p, f'{c}{tot}') for c in 'BCDEFGHI'}
        out['rows'][key] = [{c: val(p, f'{c}{r}') for c in 'ABCDEFGHJK'} for r in range(first, last + 1) if val(p, f'A{r}') is not None]
    if sc.get('debugCells'):
        out['debug'] = {f'{sh}!{a}': (val(lo.sheet(sh), a), lo.sheet(sh).getCellRangeByName(a).getFormula()[:200]) for sh, a in (x.split('!') for x in sc['debugCells'])}
    out['property'] = {'C117': val(aa, 'C117'), 'amounts': {aa.getCellRangeByName(f'B{r}').getString(): val(aa, f'D{r}') for r in range(119, 150) if aa.getCellRangeByName(f'B{r}').getString()}}
    out['portfolio'] = {k: val(p, a) for k, a in (('total', 'B202'), ('recommendedTotal', 'F202'), ('finalPortfolioTotal', 'F204'), ('errors', 'BD2'), ('status', 'BC2'))}
    out['assetAllocationTargets'] = {k: val(aa, a) for k, a in (('cash', 'M3'), ('fi', 'M4'), ('ae', 'M5'), ('ge', 'M6'), ('prop', 'M7'), ('ms', 'M8'))}
    out['profileTable'] = {h: {k: aa.getCellByPosition(2 + i, r).getValue() for k, r in {**TARGET_ROWS, 'lowerRisk': 2, 'growth': 5}.items()} for i, h in enumerate(headers) if h}
    return out


def main():
    scenarios = json.load(open(sys.argv[1]))
    tmp = tempfile.mkdtemp(prefix='golden-')
    results = []
    lo = LO(port=int(os.environ.get('LO_PORT', '2002')))
    try:
        for sc in scenarios:
            path = os.path.join(tmp, 'wb.xlsm')
            shutil.copyfile(WORKBOOK, path)  # fresh copy per scenario
            lo.open(path)
            try:
                results.append({'scenario': sc, 'workbook': run(lo, sc)})
                print(f"  {sc['id']} done ({len(results)}/{len(scenarios)})", flush=True)
            finally:
                lo.doc.close(True)
    finally:
        lo.close()
        shutil.rmtree(tmp, ignore_errors=True)
    json.dump(results, open(sys.argv[2], 'w'), indent=1, default=str)
    print(f'{len(results)} scenarios written to {sys.argv[2]}')


if __name__ == '__main__':
    main()

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
    out = {'id': sc['id'], 'targetsUsed': targets, 'sectors': {}, 'rows': {}}
    for key, first, last, tot in SECTIONS:
        out['sectors'][key] = {c: val(p, f'{c}{tot}') for c in 'BCDEFGHI'}
        out['rows'][key] = [{c: val(p, f'{c}{r}') for c in 'ABCDEFGH'} for r in range(first, last + 1) if val(p, f'A{r}') is not None]
    out['portfolio'] = {k: val(p, a) for k, a in (('total', 'B202'), ('recommendedTotal', 'F202'), ('finalPortfolioTotal', 'F204'), ('errors', 'BD2'), ('status', 'BC2'))}
    out['assetAllocationTargets'] = {k: val(aa, a) for k, a in (('cash', 'M3'), ('fi', 'M4'), ('ae', 'M5'), ('ge', 'M6'), ('prop', 'M7'), ('ms', 'M8'))}
    out['profileTable'] = {h: {k: aa.getCellByPosition(2 + i, r).getValue() for k, r in {**TARGET_ROWS, 'lowerRisk': 2, 'growth': 5}.items()} for i, h in enumerate(headers) if h}
    return out


def main():
    scenarios = json.load(open(sys.argv[1]))
    tmp = tempfile.mkdtemp(prefix='golden-')
    results = []
    lo = LO()
    try:
        for sc in scenarios:
            path = os.path.join(tmp, 'wb.xlsm')
            shutil.copyfile(WORKBOOK, path)  # fresh copy per scenario
            lo.open(path)
            try:
                results.append({'scenario': sc, 'workbook': run(lo, sc)})
            finally:
                lo.doc.close(True)
    finally:
        lo.close()
        shutil.rmtree(tmp, ignore_errors=True)
    json.dump(results, open(sys.argv[2], 'w'), indent=1, default=str)
    print(f'{len(results)} scenarios written to {sys.argv[2]}')


if __name__ == '__main__':
    main()

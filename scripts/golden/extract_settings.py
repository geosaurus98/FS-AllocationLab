"""Extract the firm-level inputs of the recommendations workbook into src/engine/defaults/*.json.

Only literal input cells are read (never formulas), so the app's defaults equal the workbook's.
Usage: python3 scripts/golden/extract_settings.py
"""
import json, os, re
import openpyxl

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
WB = os.path.join(ROOT, 'Recommendations Master Equity Weighting.xlsm')
wb = openpyxl.load_workbook(WB)  # formulas, not values
aa = wb['Asset Allocation']


def lit(cell):
    v = aa[cell].value
    if isinstance(v, str) and v.startswith('='):
        return None
    return v


def num(cell):
    v = lit(cell)
    return float(v) if isinstance(v, (int, float)) else None


def text(cell):
    v = aa[cell].value
    if v is None or (isinstance(v, str) and v.startswith('=')):
        return ''
    return str(v).strip()


def model_list(rows, name_col='B', dist_col='C'):
    out = []
    for r in rows:
        n = text(f'{name_col}{r}')
        out.append({'name': n, 'distribution': num(f'{dist_col}{r}')})
    return out


tactical = {}
for col, prof in zip('CDEFG', ['Conservative', 'Moderate', 'Balanced', 'Growth', 'Aggressive']):
    tactical[prof] = {'lowerRisk': num(f'{col}3'), 'cash': num(f'{col}4'), 'fi': num(f'{col}5'), 'growth': num(f'{col}6'),
                      'ae': num(f'{col}7'), 'ge': num(f'{col}8'), 'prop': num(f'{col}9'), 'ms': num(f'{col}10') or 0}
tactical['Custom'] = {'lowerRisk': num('H3'), 'cash': num('H4'), 'fi': num('H5'), 'growth': num('H6'), 'ms': num('H10') or 0}
tactical['Master Risk Alloc'] = {'cash': num('J4'), 'fi': num('J5'), 'ae': num('J7'), 'ge': num('J8'), 'prop': num('J9'), 'ms': num('J10')}


def band(from_cell, to_cell, jrow):
    return {'from': num(from_cell) if from_cell else 0, 'to': num(to_cell) if to_cell else None,
            'nzFunds': num(f'J{jrow}'), 'nzShares': num(f'J{jrow + 1}'), 'aus': num(f'J{jrow + 2}'), 'int': num(f'J{jrow + 3}')}


property_settings = {
    'bands': [band(None, 'I119', 119), band('H124', 'I124', 124), band('H129', 'I129', 129), band('H134', 'I134', 134), band('H139', None, 139)],
    'nzShares': model_list(range(119, 129)),
    'nzFunds': model_list(range(130, 136)),
    'ausBanded': model_list([137])[0],
    'ausFlat': model_list(range(138, 143)),
    'intBanded': model_list([144])[0],
    'intFlat': model_list(range(145, 150)),
    'flatMinimumPot': 30000,  # hard-coded in D138:D142 / D145:D149
    'fundLimit': num('K146'),
    'shareLimit': num('K147'),
}

general = {
    'portfolioTolerance': num('Q10'),
    'portfolioRounding': num('S10'),
    'portfolioMinRemainder': num('U10'),
    'assetClassTolerance': num('W10'),
}

# Every text value in Asset Allocation column B: the property pot excludes "kept" holdings whose
# name is NOT in this column (COUNTIF($B:$B, name)=0).
column_b = sorted({str(c.value).strip() for c in aa['B'] if isinstance(c.value, str) and not c.value.startswith('=') and c.value.strip()})

out = {
    'source': 'Recommendations Master Equity Weighting.xlsm (Asset Allocation sheet)',
    'status': 'approved',
    'tactical': tactical,
    'property': property_settings,
    'general': general,
    'assetAllocationColumnB': column_b,
}
path = os.path.join(ROOT, 'src', 'engine', 'defaults', 'workbook-settings.json')
json.dump(out, open(path, 'w'), indent=1)
print('wrote', path)

"""Extract the recommendations workbook into a JSON model the TypeScript evaluator runs.

Reads Portfolio, Asset Allocation, Data and ValidationLists: every literal value and formula (with
its array-formula flag), the defined names and the iterative-calculation settings. Macros are not
read. Re-run whenever the workbook changes:

    python3 scripts/engine/extract_workbook.py ["path/to/workbook.xlsm"]

Writes src/engine/workbook/model.json and src/engine/workbook/cached-values.json (Excel's last
saved results, used to verify the evaluator).
"""
import json, os, sys
import openpyxl
from openpyxl.worksheet.formula import ArrayFormula

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
SRC = sys.argv[1] if len(sys.argv) > 1 else os.path.join(ROOT, 'Recommendations Master Equity Weighting.xlsm')
SHEETS = ['Portfolio', 'Asset Allocation', 'Data', 'ValidationLists']
OUT = os.path.join(ROOT, 'src', 'engine', 'workbook')

wb = openpyxl.load_workbook(SRC)
wbv = openpyxl.load_workbook(SRC, data_only=True)


def enc(v):
    if isinstance(v, bool):
        return ['b', v]
    if isinstance(v, (int, float)):
        return ['n', float(v)]
    import datetime
    if isinstance(v, datetime.time):
        return ['n', (v.hour * 3600 + v.minute * 60 + v.second) / 86400]
    if isinstance(v, datetime.timedelta):
        return ['n', v.total_seconds() / 86400]
    if hasattr(v, 'isoformat'):  # datetime/date -> Excel serial
        base = datetime.datetime(1899, 12, 30)
        d = v if isinstance(v, datetime.datetime) else datetime.datetime(v.year, v.month, v.day)
        return ['n', (d - base).total_seconds() / 86400]
    s = str(v)
    if s.startswith('#') and s.rstrip('!?') in ('#N/A', '#VALUE', '#REF', '#DIV/0', '#NUM', '#NAME', '#NULL'):
        return ['e', s]
    return ['s', s]


model = {'source': os.path.basename(SRC), 'sheets': {}, 'names': {}, 'iterate': bool(wb.calculation.iterate),
         'iterateCount': int(wb.calculation.iterateCount or 100), 'iterateDelta': float(wb.calculation.iterateDelta or 0.001)}
cached = {}
for name in SHEETS:
    ws = wb[name]
    cells = []
    cv = {}
    max_r = max_c = 0
    for row in ws.iter_rows():
        for c in row:
            v = c.value
            if v is None:
                continue
            max_r = max(max_r, c.row); max_c = max(max_c, c.column)
            if isinstance(v, ArrayFormula):
                cells.append([c.row, c.column, 'a', v.text[1:] if v.text.startswith('=') else v.text])
            elif isinstance(v, str) and v.startswith('='):
                cells.append([c.row, c.column, 'f', v[1:]])
            else:
                k, p = enc(v)
                cells.append([c.row, c.column, k, p])
                continue
            r = wbv[name].cell(c.row, c.column).value
            if r is not None:
                cv[f'{c.row},{c.column}'] = enc(r)
    model['sheets'][name] = {'maxRow': max_r, 'maxCol': max_c, 'cells': cells}
    cached[name] = cv
for n, d in wb.defined_names.items():
    if n.startswith('_xlpm') or d.attr_text.startswith('#'):
        continue
    model['names'][n] = d.attr_text
os.makedirs(OUT, exist_ok=True)
json.dump(model, open(os.path.join(OUT, 'model.json'), 'w'), separators=(',', ':'))
json.dump(cached, open(os.path.join(OUT, 'cached-values.json'), 'w'), separators=(',', ':'))
print({k: len(v['cells']) for k, v in model['sheets'].items()}, 'names', len(model['names']))

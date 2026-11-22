"""Trim run_workbook.py output to what src/engine/workbook/golden.test.ts compares."""
import json, sys
data = json.load(open(sys.argv[1]))
mode = sys.argv[3] if len(sys.argv) > 3 else 'full'
out = {'mode': mode, 'scenarios': []}
for d in data:
    w = d['workbook']
    out['scenarios'].append({
        'scenario': {k: v for k, v in d['scenario'].items() if k != 'debugCells'},
        'rows': {k: [[r['A'], r['F']] for r in rows] for k, rows in w['rows'].items()},
        'sectors': {k: [v['F'], v['G']] for k, v in w['sectors'].items()},
        'checks': w['portfolio']['errors'],
    })
json.dump(out, open(sys.argv[2], 'w'), separators=(',', ':'))
print(len(out['scenarios']), 'scenarios ->', sys.argv[2])

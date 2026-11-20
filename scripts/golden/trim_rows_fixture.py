"""Keep what the per-holding golden tests need from run_workbook.py output."""
import json, sys
data = json.load(open(sys.argv[1]))
sections = sys.argv[3].split(',') if len(sys.argv) > 3 else ['cash', 'prop', 'ms']
out = []
for d in data:
    w = d['workbook']
    out.append({'scenario': d['scenario'], 'workbook': {
        'sectors': {k: {c: v.get(c) for c in 'BEFG'} for k, v in w['sectors'].items()},
        'rows': {k: [{c: r.get(c) for c in 'ABFGK'} for r in w['rows'][k]] for k in sections},
        'property': w.get('property'),
        'portfolio': {'finalPortfolioTotal': w['portfolio']['finalPortfolioTotal']},
    }})
json.dump(out, open(sys.argv[2], 'w'), indent=1)
print(f'{len(out)} fixtures -> {sys.argv[2]}')

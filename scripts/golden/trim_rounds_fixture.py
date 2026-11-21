"""Trim run_workbook.py output for round scenarios (src/engine/workbook/rounds.golden.test.ts)."""
import json, sys
data = json.load(open(sys.argv[1]))
out = []
for d in data:
    w = d['workbook']
    out.append({
        'scenario': d['scenario'],
        'rows': {k: [{c: r.get(c) for c in ['A', 'E', 'F', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'T', 'BA']} for r in rows] for k, rows in w['rows'].items()},
        'implementation': w.get('implementation'),
        'checks': w['portfolio']['errors'],
    })
json.dump(out, open(sys.argv[2], 'w'), separators=(',', ':'))
print(len(out), 'round scenarios ->', sys.argv[2])

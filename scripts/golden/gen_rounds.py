"""Round scenarios built on top of full-portfolio results (needs a run_workbook.py output file)."""
import json, os, random, sys

rnd = random.Random(31)
full = json.load(open(sys.argv[1]))
POOLS = json.load(open(os.path.join(os.path.dirname(__file__), 'name-pools.json')))
BONDS = [x['name'] for x in POOLS['Fixed Interest'] if x['type'] == 'Bond' and x['name'][-4:].isdigit()]
out = []
for d in rnd.sample(full, int(sys.argv[3]) if len(sys.argv) > 3 else 12):
    sc = dict(d['scenario'])
    sc['id'] = sc['id'].replace('full', 'rounds')
    sc['implementationOutputs'] = True
    sc['custodian'] = rnd.choice(['Apex', 'OA'])
    rounds = {}
    for key, rows in d['workbook']['rows'].items():
        for r in rows:
            e = r.get('E')
            if not isinstance(e, (int, float)) or e == 0 or rnd.random() < 0.3:
                continue
            n = rnd.choice([1, 1, 2, 3])
            parts = sorted(rnd.random() for _ in range(n - 1))
            fracs = [b - a for a, b in zip([0] + parts, parts + [rnd.choice([1, 1, 0.8])])]
            rounds[r['A']] = [round(e * f, 2) for f in fracs]
    # actual bonds bought against the "YYYY Bonds" placeholders
    adds = []
    for r in d['workbook']['rows'].get('fi', []):
        if 'Bonds (BBB+' in r['A'] and isinstance(r.get('F'), (int, float)) and r['F'] > 0:
            year = r['A'][:4]
            cands = [b for b in BONDS if b.endswith(year)]
            if cands:
                b = rnd.choice(cands)
                adds.append(['fi', b])
                rounds[b] = [round(r['F'] * rnd.choice([0.5, 1.0]), 2)]
    sc['addHoldings'] = adds
    sc['rounds'] = rounds
    out.append(sc)
json.dump(out, open(sys.argv[2], 'w'), indent=1)
print(len(out), 'round scenarios;', sum(len(s['addHoldings']) for s in out), 'added bonds')

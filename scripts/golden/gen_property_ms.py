"""Synthetic scenarios exercising property, multi-sector and cash rows (deterministic seed)."""
import json, random, sys

rnd = random.Random(424242)
PROFILES = ['Conservative', 'Moderate', 'Balanced', 'Growth', 'Aggressive', 'Custom']
MODEL_PROP = ['Goodman Property Trust', 'Precinct Properties New Zealand Limited', 'Vital Healthcare Property Trust Units',
              'Property For Industry Limited', 'Kernel Wealth NZ Commercial Property Fund',
              'Vanguard Australian Property Securities Index ETF', 'Kernel Wealth Global Property (NZD Hedged) Fund (PIE)']
OTHER_PROP = ['Argosy Property Trust', 'Kiwi Property Group Ltd', 'Mint Australasian Property Fund', 'Smart NZ Property ETF',
              'Fisher Funds Property and Infrastructure Fund', 'Goodman Group', 'Dexus AREIT Fund', 'Not In Datasheet Property Co']
MS = ['Milford Balanced Fund', 'OneAnswer Balanced Fund (MAC)', 'Simplicity Growth Investment Fund']
CASH = ['NZD', 'Amova NZ Cash Fund', 'AUD', 'USD']


def m(lo, hi):
    return round(rnd.uniform(lo, hi), 2)


def scenario(i):
    profile = PROFILES[i % len(PROFILES)]
    # portfolio sizes chosen so the property pot crosses every band (0-10k, 10-20k, 20-25k, 25-30k, >30k)
    size = rnd.choice([60000, 150000, 300000, 450000, 600000, 700000, 900000, 1500000, 3000000])
    sc = {'id': f'pm-{i:02d}', 'profile': profile, 'holdings': {}}
    if profile == 'Custom':
        g = round(rnd.uniform(0.3, 0.9), 3)
        c = round(rnd.uniform(0.01, 0.05), 3)
        sc['custom'] = {'lowerRisk': round(1 - g, 3), 'cash': c, 'fi': round(1 - g - c, 3), 'growth': g}
    cash = rnd.sample(CASH, rnd.choice([1, 1, 2]))
    sc['holdings']['cash'] = [[n, m(5000, 60000)] for n in cash]
    sc['holdings']['fi'] = [['Harbour NZ Corporate Bond Fund', round(size * 0.35, 2)]]
    props = []
    for name in rnd.sample(MODEL_PROP, rnd.randint(0, 4)) + rnd.sample(OTHER_PROP, rnd.randint(0, 3)):
        v = rnd.choice([m(500, 8000), m(5000, 40000), 0.0]) if rnd.random() < 0.9 else m(100, 400)
        row = [name, v]
        r = rnd.random()
        if r < 0.12:
            row.append(0)                       # keep as is
        elif r < 0.2:
            row.append(rnd.choice([m(300, 6000), -round(v * 0.5, 2)]))
        props.append(row)
    sc['holdings']['prop'] = props
    if rnd.random() < 0.4:
        sc['holdings']['ms'] = [[n, m(1000, 50000)] + ([0] if rnd.random() < 0.2 else []) for n in rnd.sample(MS, rnd.randint(1, 3))]
        sc['targetsPct'] = {'ms': rnd.choice([0, 2, 5])}
    sc['feesAndWithdrawals'] = rnd.choice([0, m(5000, 40000)])
    if rnd.random() < 0.2:
        sc['addRemove'] = rnd.choice([m(5000, 80000), -m(5000, 30000)])
    if rnd.random() < 0.15:
        sc['sectorOverrides'] = {'prop': rnd.choice([m(2000, 15000), -m(1000, 5000)])}
    return sc


json.dump([scenario(i) for i in range(int(sys.argv[2]) if len(sys.argv) > 2 else 40)], open(sys.argv[1], 'w'), indent=1)

"""Generate synthetic scenarios for sector-target golden tests (deterministic seed)."""
import json, random, sys

rnd = random.Random(20261120)
PROFILES = ['Conservative', 'Moderate', 'Balanced', 'Growth', 'Aggressive', 'Custom', 'Current', 'Master Risk Alloc']
KEYS = ['cash', 'fi', 'ae', 'ge', 'prop', 'ms']


def money(lo, hi):
    return round(rnd.uniform(lo, hi), 2)


def scenario(i):
    profile = PROFILES[i % len(PROFILES)]
    sc = {'id': f'st-{i:02d}', 'profile': profile, 'holdings': {}}
    if profile == 'Custom':
        g = round(rnd.uniform(0.1, 0.95), 3)
        cash = round(rnd.uniform(0.01, min(0.1, 1 - g)), 3)
        sc['custom'] = {'lowerRisk': round(1 - g, 3), 'cash': cash, 'fi': round(1 - g - cash, 3), 'growth': g}
    for k in KEYS:
        count = rnd.choice([0, 1, 1, 2, 3]) if k != 'ms' else rnd.choice([0, 0, 1])
        if k == 'cash' and i % 7 == 3:
            count = 0  # exercise the blank-cash quirk
        sc['holdings'][k] = [[f'Synthetic {k.upper()} holding {j + 1}', money(0, 400000)] for j in range(count)]
    if not any(sc['holdings'].values()):
        sc['holdings']['fi'] = [['Synthetic FI holding 1', 100000.0]]
    sc['feesAndWithdrawals'] = rnd.choice([0, None, money(1000, 30000), money(30000, 90000)])
    sc['addRemove'] = rnd.choice([None, None, None, 0, money(1000, 150000), -money(1000, 60000)])
    if rnd.random() < 0.3:
        sc['targetsPct'] = {'cash': round(rnd.uniform(0, 10), 2), 'fi': round(rnd.uniform(0, 70), 2)}
        if rnd.random() < 0.3:
            sc['targetsPct']['cash'] = None  # blank cell
    if rnd.random() < 0.2:
        sc.setdefault('targetsPct', {})['ms'] = round(rnd.uniform(0, 8), 2)
    ov = {}
    for k in ['fi', 'ae', 'ge', 'prop']:
        if rnd.random() < 0.15:
            ov[k] = rnd.choice([0, money(1000, 50000), -money(1000, 50000)])
    if ov:
        sc['sectorOverrides'] = ov
    return sc


json.dump([scenario(i) for i in range(int(sys.argv[2]) if len(sys.argv) > 2 else 40)], open(sys.argv[1], 'w'), indent=1)

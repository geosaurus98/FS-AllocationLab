"""Synthetic full-portfolio scenarios (all sections) for golden tests. Deterministic seed.

Names come from the workbook's Data sheet (scripts/golden/name-pools.json); values are random.
"""
import datetime, json, os, random, sys

rnd = random.Random(int(sys.argv[3]) if len(sys.argv) > 3 else 77)
POOLS = json.load(open(os.path.join(os.path.dirname(__file__), 'name-pools.json')))
AE, GE, FI = POOLS['Australasian Equities'], POOLS['Global Equities'], POOLS['Fixed Interest']
PROFILES = ['Conservative', 'Moderate', 'Balanced', 'Growth', 'Aggressive', 'Custom']
AE_MODEL = ['Simplicity NZ Share Fund', 'Harbour NZ Index Shares Fund (PIE)', 'Vanguard Australian Shares Index ETF',
            'VanEck Vectors Australian Equal Weight ETF', 'VanEck S&P/ASX Midcap ETF', 'Hyperion Australian Growth Companies Fund (PIE)',
            'Harbour Australasian Equity Fund', 'Milford Trans-Tasman Equity Fund']
GE_MODEL = ['Berkshire Hathaway Inc B Shares', 'Milford Active Growth Fund', 'Amova Global Equity Multi-Manager (Unhedged) Fund',
            'Te Ahumairangi Global Equity Fund', 'JPMorgan Global Growth & Income PLC', 'Vanguard Total World Stock Index ETF',
            'iShares MSCI World ETF (URTH)', 'Vanguard S&P 500 ETF', 'Simplicity Hedged Global Share Fund',
            'Sprott Physical Gold and Silver Trust - USD (CEF)', 'Sprott Physical Gold Trust ETV']
FI_MODEL = ['Harbour NZ Corporate Bond Fund', 'Mercer Macquarie NZ Fixed Interest Fund', 'Mercer Macquarie NZ Short Duration Fund',
            'Amova NZ Corporate Bond Fund', 'PM Capital Enhanced Yield', 'Russell Investments Global Fixed Interest Fund', 'Simplicity NZ Bond Fund']
TD_BANKS = ['ANZ Term Deposit', 'BNZ Term Deposit', 'Westpac Term Deposit', 'Heartland Bank Limited Term Deposit', 'Rabobank Term Deposit']
BASE = datetime.date(2026, 11, 20)


def m(lo, hi):
    return round(rnd.uniform(lo, hi), 2)


def pick(pool, n, **flt):
    c = [x['name'] for x in pool if all(x.get(k) == v for k, v in flt.items())]
    return rnd.sample(c, min(n, len(c)))


def holding(name, lo, hi):
    row = [name, m(lo, hi)]
    r = rnd.random()
    if r < 0.06:
        row.append(0)
    elif r < 0.1:
        row.append(rnd.choice([m(500, 8000), -round(row[1] * 0.4, 2)]))
    return row


def td_name(bank):
    d = BASE + datetime.timedelta(days=rnd.choice([-30, 20, 60, 85, 95, 150, 300, 400, 700, 1000, 1400]))
    rate = round(rnd.uniform(3.2, 5.8), 2)
    return f'{bank} {rate:.2f}% {d.strftime("%d/%m/%Y")}'


def scenario(i):
    profile = PROFILES[i % len(PROFILES)]
    size = rnd.choice([120000, 300000, 600000, 1000000, 2000000, 3500000])
    sc = {'id': f'full-{i:02d}', 'profile': profile, 'portfolioDate': '2026-11-20', 'holdings': {}}
    if profile == 'Custom':
        g = round(rnd.uniform(0.3, 0.9), 3)
        c = round(rnd.uniform(0.01, 0.05), 3)
        sc['custom'] = {'lowerRisk': round(1 - g, 3), 'cash': c, 'fi': round(1 - g - c, 3), 'growth': g}
    s = size / 1000000
    sc['holdings']['cash'] = [['NZD', m(5000, 80000) * s], ['Amova NZ Cash Fund', m(1000, 30000) * s]][: rnd.choice([1, 2])]
    fi = [holding(n, 5000 * s, 60000 * s) for n in rnd.sample(FI_MODEL, rnd.randint(0, 3))]
    fi += [[td_name(b), round(rnd.choice([25000, 50000, 100000, 188000]) * s, 2)] for b in rnd.sample(TD_BANKS, rnd.randint(0, 3))]
    fi += [holding(n, 5000 * s, 80000 * s) for n in pick(FI, rnd.randint(0, 4), type='Bond') if '(BBB+' not in n and 'Appropriate' not in n]
    sc['holdings']['fi'] = fi
    ae = [holding(n, 2000 * s, 60000 * s) for n in rnd.sample(AE_MODEL, rnd.randint(0, 3))]
    ae += [holding(n, 1000 * s, 40000 * s) for n in pick(AE, rnd.randint(0, 6), type='Share', sub='New Zealand')]
    ae += [holding(n, 1000 * s, 40000 * s) for n in pick(AE, rnd.randint(0, 4), type='Share', sub='Australia')]
    ae += [holding(n, 2000 * s, 50000 * s) for n in pick(AE, rnd.randint(0, 2), type='ETF')]
    sc['holdings']['ae'] = ae[:30]
    ge = [holding(n, 2000 * s, 120000 * s) for n in rnd.sample(GE_MODEL, rnd.randint(0, 4))]
    ge += [holding(n, 1000 * s, 40000 * s) for n in pick(GE, rnd.randint(0, 5), type='Share')]
    ge += [holding(n, 5000 * s, 90000 * s) for n in pick(GE, rnd.randint(0, 3), type='ETF', style='Passive')]
    ge += [holding(n, 5000 * s, 60000 * s) for n in pick(GE, rnd.randint(0, 2), type='Managed Fund', style='Active')]
    ge += [holding(n, 2000 * s, 30000 * s) for n in pick(GE, rnd.randint(0, 1), type='Gold')]
    sc['holdings']['ge'] = ge[:30]
    sc['holdings']['prop'] = [holding(n, 2000 * s, 30000 * s) for n in rnd.sample(['Goodman Property Trust', 'Kiwi Property Group Ltd', 'Vanguard Australian Property Securities Index ETF', 'Argosy Property Trust'], rnd.randint(0, 2))]
    sc['feesAndWithdrawals'] = rnd.choice([0, m(5000, 60000)])
    if rnd.random() < 0.2:
        sc['addRemove'] = rnd.choice([m(10000, 200000), -m(5000, 50000)])
    if rnd.random() < 0.1:
        sc['sectorOverrides'] = {rnd.choice(['fi', 'ae', 'ge']): rnd.choice([m(5000, 30000), -m(5000, 20000)])}
    for k, v in sc['holdings'].items():
        for row in v:
            row[1] = round(row[1], 2)
    return sc


n = int(sys.argv[2]) if len(sys.argv) > 2 else 30
json.dump([scenario(i) for i in range(n)], open(sys.argv[1], 'w'), indent=1)

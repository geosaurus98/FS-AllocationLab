import { describe, expect, it } from 'vitest';
import { computerCheck, custodianName, parseApexTransactions } from './transactions';

const header = ['', '', 'Portfolio Name', 'Portfolio Id', 'Event Date', '', '', 'Asset', 'Asset Type', 'Transaction Type', '', 'NZD Amount', '', '', '', '', '', '', '', '', 'Report Date'];
const row = (id: string, asset: string, type: string, amount: number, assetType = 'Managed Funds') => ['', '', 'Demo Client', id, '20/11/2026', '', '', asset, assetType, type, '', amount, '', '', '', '', '', '', '', '', '21/11/2026'];

describe('parseApexTransactions', () => {
  it('keeps this portfolio, drops amounts under $250, cleans names', () => {
    const r = parseApexTransactions(
      [header, row('DEMO-1', 'Harbour NZ Corporate Bond Fund (PIE) - XR', 'Sell', 5000), row('DEMO-1', 'Small Fund', 'Buy', -100), row('DEMO-2', 'Other', 'Buy', -9000), row('DEMO-1', 'Kernel Fund (Cash Dist)', 'Buy', -1200), row('DEMO-1', 'Demo Bank @ 5% Mat 2029 - NZD (DB100)', 'Buy', -20000, 'Corporate Bonds & Notes')],
      'DEMO-1',
    );
    expect(r.clientName).toBe('Demo Client');
    expect(r.reportDate?.toISOString().slice(0, 10)).toBe('2026-11-21');
    expect(r.transactions.map((t) => [t.assetName, t.amount])).toEqual([
      ['Demo Bank @ 5% Mat 2029', 20000],
      ['Harbour NZ Corporate Bond Fund (PIE) - XR', -5000],
      ['Kernel Fund', 1200],
    ]);
  });
  it('rejects files that are not Apex reports, and unknown portfolios', () => {
    expect(() => parseApexTransactions([header, ['x']], 'DEMO-1')).toThrow(/Apex/);
    expect(() => parseApexTransactions([header, row('DEMO-2', 'A', 'Buy', -500)], 'DEMO-1')).toThrow(/No transactions/);
  });
});

describe('computerCheck', () => {
  const holdings = [
    { name: 'Harbour NZ Corporate Bond Fund', apexName: 'Harbour NZ Corporate Bond Fund (PIE) - XR', oaName: '', type: 'Managed Fund' },
    { name: 'Demo Bond 2029', apexName: 'Demo Bank @ 5% Mat 2029 - NZD (DB100) (0.00; 5%; 5 Years; 01/01/2029)', oaName: '', type: 'Bond' },
    { name: 'Kernel Fund', apexName: 'Kernel Fund', oaName: '', type: 'Managed Fund' },
    { name: 'Not Traded', apexName: 'Not Traded', oaName: '', type: 'ETF' },
  ];
  const report = {
    clientName: 'Demo',
    reportDate: new Date(Date.UTC(2026, 10, 21)),
    transactions: [
      { assetName: 'Harbour NZ Corporate Bond Fund (PIE) - XR', eventDate: '', type: 'Sell', amount: -4950, assetType: '' },
      { assetName: 'Demo Bank @ 5% Mat 2029', eventDate: '', type: 'Buy', amount: 20000, assetType: '' },
      { assetName: 'Kernel Fund', eventDate: '', type: 'Buy', amount: 600, assetType: '' },
      { assetName: 'Kernel Fund', eventDate: '', type: 'Buy', amount: 600, assetType: '' },
      { assetName: 'Surprise ETF', eventDate: '', type: 'Buy', amount: 3000, assetType: '' },
    ],
  };
  it('matches within $100, adds earlier rounds within 5 days, sums repeated transactions', () => {
    const r = computerCheck({
      holdings,
      rounds: { 'Harbour NZ Corporate Bond Fund': [-3000, -2000], 'Demo Bond 2029': [null, 20000], 'Kernel Fund': [null, 1200], 'Not Traded': [null, 700] },
      roundDates: ['2026-11-18', '2026-11-21'],
      currentRound: 1,
      report,
      custodian: 'Apex',
    });
    expect(r.roundsIncluded).toEqual([0, 1]);
    expect(Object.fromEntries(r.lines.map((l) => [l.holding, l.status]))).toEqual({
      'Harbour NZ Corporate Bond Fund': 'match',
      'Demo Bond 2029': 'match',
      'Kernel Fund': 'match',
      'Not Traded': 'not-in-transactions',
    });
    expect(r.unmatchedTransactions.map((t) => t.assetName)).toEqual(['Surprise ETF']);
  });
  it('does not include rounds older than 5 days, and flags size differences', () => {
    const r = computerCheck({ holdings, rounds: { 'Harbour NZ Corporate Bond Fund': [-3000, -2000] }, roundDates: ['2026-11-01', '2026-11-21'], currentRound: 1, report, custodian: 'Apex' });
    expect(r.roundsIncluded).toEqual([1]);
    expect(r.lines[0]!.status).toBe('mismatch');
  });
  it('uses the OA name after the colon', () => {
    expect(custodianName({ name: 'X', apexName: '', oaName: 'NIKKONZC: Amova NZ Cash Fund', type: 'Cash' }, 'OA')).toBe('Amova NZ Cash Fund');
  });
});

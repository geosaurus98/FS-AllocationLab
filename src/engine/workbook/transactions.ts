/**
 * "Computer check" (port of Transactions_Match): compare the amounts entered in the implementation
 * rounds with the custodian's transaction report.
 *
 * Apex report layout (first sheet, header in row 1): C client name, D portfolio/client number,
 * E event date, H asset name, I asset type, J transaction type, L NZD amount, U report date.
 *
 * Deliberate differences from the macro (both make the check less noisy, neither changes advice):
 *  - bonds are recognised by the Data sheet's Type column; the macro tests the Style column, which
 *    never says "Bond", so its bond names were never trimmed and bonds never matched;
 *  - several transactions in the same asset are added together; the macro kept only the last one.
 */

export interface Transaction {
  assetName: string;
  eventDate: string;
  type: string;
  /** Sign as shown by the macro: −(report amount). Only the size is compared. */
  amount: number;
  assetType: string;
}

export interface TransactionReport {
  clientName: string;
  reportDate: Date | null;
  transactions: Transaction[];
}

const cellText = (v: unknown) => (v === null || v === undefined ? '' : String(v).trim());

/** Excel serial or d/m/y text → Date. */
export function toDate(v: unknown): Date | null {
  if (typeof v === 'number' && v > 20000 && v < 80000) return new Date(Math.round((v - 25569) * 86400000));
  const s = cellText(v);
  if (!s) return null;
  if (/^\d+(\.\d+)?$/.test(s)) return toDate(Number(s));
  let m = /^(\d{1,2})[/-](\d{1,2})[/-](\d{4})/.exec(s);
  if (m) return new Date(Date.UTC(+m[3]!, +m[2]! - 1, +m[1]!));
  m = /^(\d{4})-(\d{2})-(\d{2})/.exec(s);
  if (m) return new Date(Date.UTC(+m[1]!, +m[2]! - 1, +m[3]!));
  return null;
}

export function parseApexTransactions(grid: unknown[][], portfolioId: string): TransactionReport {
  const reportDate = toDate(grid[1]?.[20]); // U2
  if (!reportDate) throw new Error('This does not look like an Apex transaction report (no date in column U, row 2).');
  const id = portfolioId.trim();
  let clientName = '';
  let found = false;
  const transactions: Transaction[] = [];
  for (let i = 1; i < grid.length; i++) {
    const r = grid[i] ?? [];
    if (cellText(r[3]) !== id) continue;
    found = true;
    clientName = cellText(r[2]);
    const amount = Number(cellText(r[11]).replace(/[$,]/g, ''));
    if (!Number.isFinite(amount) || Math.abs(amount) < 250) continue;
    let name = cellText(r[7]);
    const assetType = cellText(r[8]);
    if (name.includes('(Cash Dist)')) name = name.replace('(Cash Dist)', '').trim();
    if (assetType === 'Corporate Bonds & Notes') name = name.split('-')[0]!.trim();
    const d = toDate(r[4]);
    transactions.push({ assetName: name, eventDate: d ? d.toISOString().slice(0, 10) : cellText(r[4]), type: cellText(r[9]), amount: -amount, assetType });
  }
  if (!found) throw new Error(`No transactions for portfolio ${id} in this report.`);
  transactions.sort((a, b) => a.assetName.localeCompare(b.assetName));
  return { clientName, reportDate, transactions };
}

export interface CheckHolding {
  name: string;
  apexName: string;
  oaName: string;
  type: string;
}

export interface CheckLine {
  holding: string;
  custodianName: string;
  roundsTotal: number;
  transactionAmount: number | null;
  status: 'match' | 'mismatch' | 'not-in-transactions';
}

export interface CheckResult {
  currentRound: number;
  roundsIncluded: number[];
  lines: CheckLine[];
  /** Transactions with no round entry. */
  unmatchedTransactions: Transaction[];
}

/** Name the custodian uses for a holding (SafeNameChange / Perform_Check). */
export function custodianName(h: CheckHolding, custodian: 'Apex' | 'OA'): string {
  if (custodian === 'Apex') {
    const n = h.apexName || h.name;
    return h.type === 'Bond' ? n.split('-')[0]!.trim() : n;
  }
  let n = h.oaName || h.name;
  if (n.includes(':')) n = n.replace(/ /g, ' ').split(':')[1]!.trim();
  return n;
}

export function computerCheck(args: {
  holdings: CheckHolding[];
  rounds: Record<string, (number | null)[]>;
  roundDates: (string | null)[];
  currentRound: number;
  report: TransactionReport;
  custodian: 'Apex' | 'OA';
}): CheckResult {
  const { holdings, rounds, roundDates, currentRound, report, custodian } = args;
  // earlier rounds within 5 days of the report date are counted together with the current round
  const included = [currentRound];
  const ref = report.reportDate ?? new Date();
  for (let i = currentRound - 1; i >= 0; i--) {
    const d = toDate(roundDates[i]);
    if (!d || Math.abs((ref.getTime() - d.getTime()) / 86400000) > 5) break;
    included.push(i);
  }
  const byName = new Map<string, number>();
  for (const t of report.transactions) {
    const k = t.assetName.toLowerCase();
    byName.set(k, (byName.get(k) ?? 0) + t.amount);
  }
  const used = new Set<string>();
  const lines: CheckLine[] = [];
  for (const h of holdings) {
    const cur = rounds[h.name]?.[currentRound];
    if (!cur) continue;
    const total = included.reduce((a, i) => a + (rounds[h.name]?.[i] ?? 0), 0);
    const cname = custodianName(h, custodian);
    const k = cname.toLowerCase();
    if (byName.has(k)) {
      const amt = byName.get(k)!;
      used.add(k);
      lines.push({ holding: h.name, custodianName: cname, roundsTotal: total, transactionAmount: amt, status: Math.abs(Math.abs(total) - Math.abs(amt)) <= 100 ? 'match' : 'mismatch' });
    } else {
      lines.push({ holding: h.name, custodianName: cname, roundsTotal: total, transactionAmount: null, status: 'not-in-transactions' });
    }
  }
  const unmatchedTransactions = report.transactions.filter((t) => !used.has(t.assetName.toLowerCase()));
  return { currentRound, roundsIncluded: included.sort((a, b) => a - b), lines, unmatchedTransactions };
}

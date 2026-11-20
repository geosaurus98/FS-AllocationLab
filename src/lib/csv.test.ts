import { describe, expect, it } from 'vitest';
import { exportAllocationsCsv, parseAllocationCsv, parseCsv, planImport } from './csv';
import { weightsFrom } from './reference';
import type { BaselineAllocation } from './types';

const header = 'record_type,profile,kind,name,status,source,version,cash,fixed_interest,australasian_equities,global_equities,property,multi_sector';

const mk = (status: BaselineAllocation['status']): BaselineAllocation => ({
  id: 'balanced6040:strategic', profileId: 'balanced6040', kind: 'strategic',
  weights: weightsFrom({ cash: 2, fixedInterest: 38, ausEquities: 17, globalEquities: 39, property: 4, multiSector: 0 }),
  status, source: 's', version: 1, effectiveDate: null, updatedAt: 'x', notes: '',
});

describe('parseCsv', () => {
  it('handles quotes, embedded commas and CRLF', () => {
    expect(parseCsv('a,"b, c","d ""q"""\r\n1,2,3\r\n')).toEqual([['a', 'b, c', 'd "q"'], ['1', '2', '3']]);
  });
});

describe('parseAllocationCsv', () => {
  it('parses valid rows and keeps blanks as missing', () => {
    const r = parseAllocationCsv(`${header}\nbaseline,balanced_60_40,strategic,,draft,test,,2,38,17,39,,0\n`);
    expect(r.fatal).toEqual([]);
    expect(r.errors).toEqual([]);
    expect(r.rows[0]?.weights.property).toBeNull();
    expect(r.rows[0]?.profileId).toBe('balanced6040');
  });

  it('rejects files with missing columns', () => {
    const r = parseAllocationCsv('record_type,profile,kind,cash\nbaseline,growth,strategic,2');
    expect(r.fatal[0]).toContain('Missing required column');
  });

  it('reports row-level errors with line numbers', () => {
    const r = parseAllocationCsv(`${header}\nbaseline,unknown,strategic,,,,,2,38,17,39,4,0\nscenario,growth,proposed,X,,,,abc,38,17,39,4,0\n`);
    expect(r.rows).toHaveLength(0);
    expect(r.errors.map((e) => e.line)).toEqual([2, 3]);
    expect(r.errors[1]?.message).toContain('Cash');
  });

  it('round-trips exported data', () => {
    const csv = exportAllocationsCsv([mk('illustrative')], []);
    const r = parseAllocationCsv(csv);
    expect(r.rows[0]?.weights).toEqual(mk('illustrative').weights);
  });
});

describe('planImport', () => {
  const row = parseAllocationCsv(`${header}\nbaseline,balanced_60_40,strategic,,approved,x,,5,35,17,39,4,0\n`).rows;

  it('defaults to creating a scenario — never overwrites baselines by default', () => {
    const [item] = planImport(row, [mk('synthetic')]);
    expect(item?.defaultAction).toBe('create-scenario');
    expect(item?.options).toContain('propose-baseline-revision');
    expect(item?.differsFromExisting).toBe(true);
  });

  it('never offers to overwrite an approved baseline', () => {
    const [item] = planImport(row, [mk('approved')]);
    expect(item?.options).not.toContain('propose-baseline-revision');
    expect(item?.notes.join(' ')).toMatch(/approved and locked/i);
  });

  it('explains that import cannot set approved status', () => {
    const [item] = planImport(row, [mk('synthetic')]);
    expect(item?.notes.join(' ')).toMatch(/cannot mark data as approved/);
  });
});

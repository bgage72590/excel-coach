import { describe, expect, it } from 'vitest';
import { parseCell } from '../src/engine/address';
import { GL_ACCOUNTS, VENDORS, WAREHOUSES, fromSerial } from '../src/engine/data';
import { Rng } from '../src/engine/rng';
import type { Exercise } from '../src/engine/types';
import {
  POWER_QUERY,
  cleanLine,
  distinctRows,
  groupTotals,
  pqAppend,
  pqCleanExport,
  pqGroup,
  pqMerge,
  pqProper,
  pqTrim,
  pqUnpivot,
  pqUpper,
} from '../src/exercises/powerquery';
import { SEEDS, exerciseSuite } from './helpers/suite';

exerciseSuite(POWER_QUERY);

describe('power query module', () => {
  it('exports five query exercises in the powerquery module', () => {
    expect(POWER_QUERY).toHaveLength(5);
    expect(new Set(POWER_QUERY.map((e) => e.id)).size).toBe(5);
    for (const ex of POWER_QUERY) {
      expect(ex.module).toBe('powerquery');
      expect(ex.variants).toEqual([]);
      const d = ex.make(new Rng(1));
      expect(ex.layout(d).answer.kind).toBe('query');
      expect(ex.layout(d).blocks.every((b) => b.kind !== 'data' || b.asTable)).toBe(true);
      expect(ex.hints.some((h) => /Close & Load/.test(h))).toBe(true);
      expect(ex.concept.tip).toMatch(/new sheet/);
    }
  });

  it('never sends a Mac learner to Load To or Only Create Connection without a Mac path', () => {
    // Excel for Mac has no Load To dialog: Close & Load always loads to a new sheet.
    for (const ex of POWER_QUERY) {
      for (const text of [...ex.hints, ex.concept.syntax, ex.concept.tip ?? '', ex.solution(ex.make(new Rng(1)))]) {
        if (/Only Create Connection|Close & Load To/.test(text)) expect(text, text).toMatch(/Windows/);
      }
    }
  });

  it('a helper query loaded to a sheet on a Mac can never be mistaken for the answer', () => {
    // On a Mac, the Items and monthly AP queries land on sheets as copies of their source Tables.
    for (const ex of [pqMerge, pqAppend] as Exercise<any>[]) {
      for (const seed of SEEDS) {
        const layout = ex.layout(ex.make(new Rng(seed)));
        if (layout.answer.kind !== 'query') throw new Error('query answer expected');
        const want = layout.answer.columns.map((c) => c.toLowerCase()).sort().join('|');
        for (const b of layout.blocks) if (b.kind === 'data') expect(b.columns.map((c) => c.header.toLowerCase()).sort().join('|')).not.toBe(want);
      }
    }
  });

  it('keeps every block within columns A–N and rows 1–130', () => {
    for (const ex of POWER_QUERY) {
      for (const seed of SEEDS) {
        const d = ex.make(new Rng(seed));
        for (const b of ex.layout(d).blocks) {
          const start = parseCell(b.at);
          const rows = b.kind === 'data' ? b.rows.length + 1 : b.values.length;
          const cols = b.kind === 'data' ? b.columns.length : b.values[0].length;
          expect(start.col + cols - 1).toBeLessThanOrEqual(14);
          expect(start.row + rows - 1).toBeLessThanOrEqual(130);
          if (b.kind === 'data') expect(b.rows.length).toBeLessThanOrEqual(120);
        }
      }
    }
  });

  it('asks for an output whose headers differ from every source Table, except the clean copy', () => {
    // The checker skips source Tables by name, so an output may share their headers; it must not share
    // them with a Table the learner might load by accident (the monthly AP queries, Items, Orders).
    for (const ex of [pqUnpivot, pqMerge, pqAppend, pqGroup] as Exercise<any>[]) {
      const d = ex.make(new Rng(7));
      const layout = ex.layout(d);
      const want = layout.answer.kind === 'query' ? layout.answer.columns.map((c) => c.toLowerCase()).sort().join('|') : '';
      for (const b of layout.blocks) {
        if (b.kind !== 'data') continue;
        expect(b.columns.map((c) => c.header.toLowerCase()).sort().join('|')).not.toBe(want);
      }
    }
  });
});

describe('Power Query text semantics', () => {
  it('Trim removes only leading and trailing whitespace', () => {
    expect(pqTrim('  Apex Supply   ')).toBe('Apex Supply');
    expect(pqTrim('a  b')).toBe('a  b');
  });

  it('Proper matches the documented Text.Proper example', () => {
    expect(pqProper('the QUICK BrOwn fOx')).toBe('The Quick Brown Fox');
  });

  it('every vendor and warehouse survives lower, upper and proper casing unchanged', () => {
    for (const name of [...VENDORS, ...WAREHOUSES]) {
      expect(name).toMatch(/^[A-Za-z]+( [A-Za-z]+)*$/);
      for (const v of [name, name.toLowerCase(), name.toUpperCase()]) expect(pqProper(v)).toBe(name);
    }
    expect(pqUpper('inv-20481')).toBe('INV-20481');
  });

  it('distinctRows keeps the first copy of identical rows only', () => {
    expect(distinctRows([['a', 1], ['b', 2], ['a', 1], ['a', 2]])).toEqual([['a', 1], ['b', 2], ['a', 2]]);
  });
});

describe('pq-clean-export data', () => {
  it.each(SEEDS)('seed %i: messy in every way the task promises, and cleans to real names', (seed) => {
    const d = pqCleanExport.make(new Rng(seed));
    const lines = d.rows.filter((r) => r !== null);
    const blanks = d.rows.filter((r) => r === null).length;
    expect(blanks).toBeGreaterThanOrEqual(3);
    expect(d.rows[0]).not.toBeNull();
    expect(d.rows[d.rows.length - 1]).not.toBeNull();

    const raw = lines.map((r) => JSON.stringify(r));
    const duplicates = raw.length - new Set(raw).size;
    expect(duplicates).toBeGreaterThanOrEqual(3);

    const texts = lines.flatMap((r) => [r.invoice, r.vendor, r.warehouse]);
    expect(texts.some((t) => t.startsWith(' '))).toBe(true);
    expect(texts.some((t) => t.endsWith(' '))).toBe(true);
    // Power Query's Trim leaves inner runs alone, so the data must never double a space inside the text.
    for (const t of texts) expect(pqTrim(t)).not.toMatch(/ {2}/);
    expect(lines.some((r) => /inv/.test(r.invoice))).toBe(true);
    expect(lines.some((r) => pqTrim(r.vendor) !== pqProper(pqTrim(r.vendor)))).toBe(true);
    expect(lines.some((r) => pqTrim(r.warehouse) !== pqProper(pqTrim(r.warehouse)))).toBe(true);

    for (const r of lines) {
      const [invoice, vendor, warehouse, amount] = cleanLine(r);
      expect(invoice).toMatch(/^INV-\d{5}$/);
      expect(VENDORS).toContain(vendor);
      expect(WAREHOUSES).toContain(warehouse);
      expect(typeof amount).toBe('number');
    }
  });

  it.each(SEEDS)('seed %i: the answer is the same whichever order the steps run in', (seed) => {
    const d = pqCleanExport.make(new Rng(seed));
    const lines = d.rows.filter((r) => r !== null);
    const expected = pqCleanExport.expected(d);
    // Remove duplicates first, clean second: same rows, because duplicates are exact copies and invoices are unique.
    const dedupeFirst = distinctRows(lines.map((r) => [r.invoice, r.vendor, r.warehouse, r.amount])).map((r) =>
      cleanLine({ invoice: r[0] as string, vendor: r[1] as string, warehouse: r[2] as string, amount: r[3] as number }),
    );
    expect(dedupeFirst).toEqual(expected);
    expect(new Set(expected.map((r) => r[0])).size).toBe(expected.length);
  });
});

describe('pq-unpivot data', () => {
  it.each(SEEDS)('seed %i: one output row per account per month, no blanks for Unpivot to drop', (seed) => {
    const d = pqUnpivot.make(new Rng(seed));
    expect(d.months).toHaveLength(6);
    for (const m of d.months) expect(m).toMatch(/^[A-Z][a-z]{2}$/);
    expect(new Set(d.rows.map((r) => r.account)).size).toBe(d.rows.length);
    for (const r of d.rows) {
      expect(GL_ACCOUNTS.map((a) => a.name)).toContain(r.account);
      for (const a of r.amounts) {
        expect(Number.isInteger(a)).toBe(true);
        expect(a).toBeGreaterThan(0);
      }
    }
    const expected = pqUnpivot.expected(d);
    expect(expected).toHaveLength(d.rows.length * d.months.length);
    expect(expected[1]).toEqual([d.rows[0].account, d.months[1], d.rows[0].amounts[1]]);
  });
});

describe('pq-merge data', () => {
  it.each(SEEDS)('seed %i: every order SKU exists once in Items', (seed) => {
    const d = pqMerge.make(new Rng(seed));
    const skus = d.items.map((i) => i.sku);
    expect(new Set(skus).size).toBe(skus.length);
    for (const o of d.orders) expect(skus.filter((s) => s === o.sku)).toHaveLength(1);
    expect(d.orders.length).toBeGreaterThanOrEqual(36);
    const expected = pqMerge.expected(d);
    expect(expected).toHaveLength(d.orders.length);
    for (const row of expected) expect(['Packaging', 'Hardware', 'Electrical', 'Safety', 'Janitorial']).toContain(row[3]);
  });
});

describe('pq-append data', () => {
  it.each(SEEDS)('seed %i: all rows from the three months, dates inside their month, the count stated in the task', (seed) => {
    const d = pqAppend.make(new Rng(seed));
    const total = d.months.reduce((n, m) => n + m.rows.length, 0);
    expect(pqAppend.expected(d)).toHaveLength(total);
    expect(pqAppend.task(d)).toContain(`all ${total} rows`);
    d.months.forEach((m, i) => {
      for (const r of m.rows) expect(fromSerial(r.date)).toMatchObject({ year: 2026, month: 7 + i });
    });
    const invoices = d.months.flatMap((m) => m.rows.map((r) => r.invoice));
    expect(new Set(invoices).size).toBe(invoices.length);
    expect(pqAppend.layout(d).blocks.filter((b) => b.kind === 'data').map((b) => (b.kind === 'data' ? b.table : ''))).toEqual(['AP_Jul', 'AP_Aug', 'AP_Sep']);
  });
});

describe('pq-group data', () => {
  it.each(SEEDS)('seed %i: totals add up to all units, one row per value', (seed) => {
    const d = pqGroup.make(new Rng(seed));
    const totals = groupTotals(d);
    const field = d.by === 'Warehouse' ? 'warehouse' : 'category';
    expect(totals).toHaveLength(new Set(d.rows.map((r) => r[field])).size);
    expect(totals.reduce((n, r) => n + (r[1] as number), 0)).toBe(d.rows.reduce((n, r) => n + r.units, 0));
    const layout = pqGroup.layout(d);
    expect(layout.answer).toEqual({ kind: 'query', columns: [d.by, 'Total units'], order: 'any' });
    expect(pqGroup.task(d)).toContain(`\`${d.by}\``);
  });

  it('uses both grouping columns across seeds', () => {
    const fields = new Set(Array.from({ length: 20 }, (_, s) => pqGroup.make(new Rng(s + 1)).by));
    expect(fields).toEqual(new Set(['Warehouse', 'Category']));
  });
});

import { describe, expect, it } from 'vitest';
import { parseCell, rangeSize, type RangeRef } from '../src/engine/address';
import { CARRIERS, VENDORS, eomonth } from '../src/engine/data';
import { gradeRules } from '../src/engine/grade';
import { Rng } from '../src/engine/rng';
import type { Block, Cell, DataBlock, Grid } from '../src/engine/types';
import { EXERCISES } from '../src/exercises';
import { isWeekend } from '../src/exercises/dates';
import { FINANCE_MISSIONS } from '../src/missions/finance';
import { OPS_MISSIONS } from '../src/missions/ops';
import {
  ALL_LANES,
  AS_OF_DATES,
  BOUNDARY_DAYS,
  BUCKETS,
  EXTRA_MISSIONS,
  LIMIT_MARGIN,
  NEW_CARRIERS,
  OVER_60_LIMITS,
  REFRESH_COLS,
  REVIEW_MONTHS,
  SUMMARY_COLUMNS,
  TOP_LANE_LEAD,
  apAging,
  bucketFor,
  callList,
  carrierCostReview,
  daysPastDue,
  laneRanking,
  limitWorks,
  over60,
  rowsAtStep,
  weeklyRefresh,
  type AgingData,
  type ApInvoice,
  type CarrierData,
  type RefreshData,
} from '../src/missions/extra';
import type { Mission } from '../src/missions/types';
import { SEEDS, blockRange, missionSuite, overlaps } from './helpers/suite';

missionSuite(EXTRA_MISSIONS);

/** Session seeds are random 32-bit numbers, so invariants are checked over many of them. */
const MANY_SEEDS = [...SEEDS, ...Array.from({ length: 300 }, (_, k) => ((k + 1) * 2654435761) >>> 0)];

/**
 * The base data for a seed plus every variant of the given steps (all by default), seeded exactly
 * as ExcelHost.check seeds them (seed + 7919 × (variant index + 1)): the states a learner meets.
 */
function allStates<D>(m: Mission<D>, seed: number, steps = m.steps.map((_, i) => i)): D[] {
  const base = m.make(new Rng(seed));
  return [base, ...steps.flatMap((s) => m.steps[s].variants.map((v, i) => v.apply(base, new Rng(seed + 7919 * (i + 1)))))];
}

/** Reads a Table's rows back out of the sheet blocks, keyed by header: what Excel will hold. */
function tableRows(blocks: Block[], table: string): Record<string, Cell>[] {
  const b = blocks.find((x): x is DataBlock => x.kind === 'data' && x.table === table)!;
  return b.rows.map((r) => Object.fromEntries(b.columns.map((c, i) => [c.header, r[i]])));
}

function cellsAt(blocks: Block[], at: string): Cell[][] {
  const b = blocks.find((x) => x.kind === 'cells' && x.at === at)!;
  return b.kind === 'cells' ? b.values : [];
}

/** Excel's text comparison in SUMIFS, COUNTIFS and = tests on plain ASCII text: case-insensitive. */
const eq = (a: Cell, b: Cell) => String(a).toLowerCase() === String(b).toLowerCase();
const total = (xs: number[]) => xs.reduce((a, b) => a + b, 0);
const near = (a: number, b: number) => Math.abs(a - b) <= 1e-9 * Math.max(1, Math.abs(b));

/** Words the copy guidelines rule out (the shared suite checks titles, tasks and hints; this adds briefs and notes). */
const BANNED = /\b(please|simply|just|easy|easily|successfully|leverage|seamless)\b|!/i;

describe('extra missions: metadata', () => {
  it('only list skills that exist', () => {
    const ids = new Set(EXERCISES.map((e) => e.id));
    for (const m of EXTRA_MISSIONS) for (const s of m.skills) expect(ids.has(s), `${m.id} skill ${s}`).toBe(true);
  });

  it('have ids no other mission uses, and the roles they were planned for', () => {
    const ids = [...FINANCE_MISSIONS, ...OPS_MISSIONS, ...EXTRA_MISSIONS].map((m) => m.id);
    expect(new Set(ids).size).toBe(ids.length);
    expect(EXTRA_MISSIONS.map((m) => [m.id, m.role])).toEqual([
      ['mission-ap-aging', 'finance'],
      ['mission-carrier-cost-review', 'ops'],
      ['mission-weekly-refresh', 'ops'],
    ]);
  });

  it('brief from a named manager, and briefs and step notes follow the copy rules', () => {
    for (const m of EXTRA_MISSIONS) {
      const d = m.make(new Rng(1));
      const brief = m.brief(d);
      expect(brief.from).toMatch(/^[A-Z][a-z]+ [A-Z][a-z]+, /);
      const notes = m.steps.flatMap((s) => (s.startNote ? [s.startNote(d)] : []));
      for (const text of [brief.subject, brief.body, ...notes]) expect(text, `copy: ${text}`).not.toMatch(BANNED);
    }
  });

  it('use typographic apostrophes in everything the learner reads', () => {
    for (const m of EXTRA_MISSIONS) {
      const d = m.make(new Rng(7));
      const brief = m.brief(d);
      const copy = [m.title, m.summary, brief.subject, brief.body, ...m.steps.flatMap((s) => [s.title, s.task(d), ...s.hints, ...(s.startNote ? [s.startNote(d)] : [])])];
      for (const text of copy) expect(text, text).not.toMatch(/'/);
    }
  });
});

describe('extra missions: sheet layout beyond the shared checks', () => {
  it.each(SEEDS)('seed %i: nothing sits in the 80 rows the host formats below a spill anchor', (seed) => {
    for (const m of EXTRA_MISSIONS) {
      const d = m.make(new Rng(seed));
      for (const step of m.steps) {
        const a = step.answer(d);
        if (a.kind !== 'spill') continue;
        const start = parseCell(a.anchor);
        const zone: RangeRef = { start, end: { row: start.row + 79, col: start.col + step.expected(d)[0].length } };
        for (const b of m.blocks(d)) expect(overlaps(zone, blockRange(b)), `${m.id} block at ${b.at}`).toBe(false);
      }
    }
  });

  it.each(SEEDS)('seed %i: the Invoices Table has room for its two new columns plus a gap', (seed) => {
    const d = apAging.make(new Rng(seed));
    const blocks = apAging.blocks(d);
    const invoices = blocks.find((b) => b.kind === 'data' && b.table === 'Invoices')!;
    const r = blockRange(invoices);
    expect(r.end.col).toBe(6); // A:F; the learner adds G and H
    const room: RangeRef = { start: { row: 1, col: 7 }, end: { row: r.end.row + 1, col: 9 } };
    for (const b of blocks) if (b !== invoices) expect(overlaps(room, blockRange(b)), `block at ${b.at}`).toBe(false);
    const added = apAging.steps.slice(0, 2).map((s) => s.answer(d));
    expect(added.map((a) => (a.kind === 'tableColumn' ? `${a.table}[${a.column}]` : a.kind))).toEqual(['Invoices[Days past due]', 'Invoices[Bucket]']);
  });

  it.each(SEEDS)('seed %i: the weekly Shipments Table is alone on the sheet, so appended weeks never hit anything', (seed) => {
    const blocks = weeklyRefresh.blocks(weeklyRefresh.make(new Rng(seed)));
    expect(blocks).toHaveLength(1);
    expect(blocks[0]).toMatchObject({ kind: 'data', table: 'Shipments', at: 'A1', asTable: true });
  });
});

// ---------- AP aging ----------

describe('AP aging: bucket boundaries', () => {
  it('buckets each boundary day the way XLOOKUP with match_mode -1 does', () => {
    expect(BOUNDARY_DAYS.map(bucketFor)).toEqual(['Current', '1–30 days', '1–30 days', '31–60 days', '31–60 days', '61–90 days', '61–90 days', 'Over 90 days']);
    expect(bucketFor(15)).toBe('1–30 days');
    expect(bucketFor(400)).toBe('Over 90 days');
    expect(() => bucketFor(-1)).toThrow();
  });

  it('puts the cutoffs XLOOKUP reads on the sheet, and labels the grid with the same buckets in order', () => {
    const blocks = apAging.blocks(apAging.make(new Rng(1)));
    expect(tableRows(blocks, 'Buckets')).toEqual(BUCKETS.map((b) => ({ 'Min days': b.min, Bucket: b.label })));
    expect(cellsAt(blocks, 'J12')[0]).toEqual(['Vendor', ...BUCKETS.map((b) => b.label)]);
    // None of the labels can be read by Excel as a number or a date.
    for (const b of BUCKETS) expect(b.label).toMatch(/[A-Za-z]/);
  });

  it.each(MANY_SEEDS)('seed %i: an open invoice sits on every boundary day, and steps 1 and 2 expect the right value there', (seed) => {
    const d = apAging.make(new Rng(seed));
    const days = apAging.steps[0].expected(d).map((r) => r[0]);
    const buckets = apAging.steps[1].expected(d).map((r) => r[0]);
    for (const b of BOUNDARY_DAYS) {
      const k = d.invoices.findIndex((i) => !i.paid && d.asOf - i.due === b);
      expect(k, `open invoice ${b} days past due`).toBeGreaterThanOrEqual(0);
      expect(days[k]).toBe(b);
      expect(buckets[k]).toBe(bucketFor(b));
    }
    // Not yet due shows 0 and Current; a paid invoice past its due date shows blanks.
    const early = d.invoices.findIndex((i) => !i.paid && i.due > d.asOf);
    expect(early).toBeGreaterThanOrEqual(0);
    expect([days[early], buckets[early]]).toEqual([0, 'Current']);
    const paid = d.invoices.findIndex((i) => i.paid && i.due < d.asOf);
    expect(paid).toBeGreaterThanOrEqual(0);
    expect([days[paid], buckets[paid]]).toEqual(['', '']);
  });

  it('computes every step on a hand-built ledger, edges included', () => {
    const asOf = AS_OF_DATES[1];
    const inv = (vendor: string, late: number, amount: number, paid = false): ApInvoice => ({ vendor, invoice: `X-${late}`, date: asOf - late - 30, due: asOf - late, amount, paid });
    const [a, b, c, e, f, g] = VENDORS;
    const d: AgingData = {
      asOf,
      limit: 5000,
      vendors: [a, b, c, e, f, g],
      invoices: [
        inv(a, -5, 100), // not due yet
        inv(a, 0, 200), // due today: Current
        inv(b, 1, 300),
        inv(b, 30, 400),
        inv(c, 31, 500),
        inv(c, 60, 600),
        inv(e, 61, 3000),
        inv(e, 90, 2000), // 61 + 90 = exactly 5,000: not more than the limit
        inv(f, 91, 5000.01),
        inv(g, 400, 9999, true), // paid: no days, no bucket, no balance
      ],
    };
    const [s1, s2, s3, s4, s5] = apAging.steps.map((s) => s.expected(d));
    expect(s1.map((r) => r[0])).toEqual([0, 0, 1, 30, 31, 60, 61, 90, 91, '']);
    expect(s2.map((r) => r[0])).toEqual(['Current', 'Current', '1–30 days', '1–30 days', '31–60 days', '31–60 days', '61–90 days', '61–90 days', 'Over 90 days', '']);
    expect(s3).toEqual([
      [300, 0, 0, 0, 0],
      [0, 700, 0, 0, 0],
      [0, 0, 1100, 0, 0],
      [0, 0, 0, 5000, 0],
      [0, 0, 0, 0, 5000.01],
      [0, 0, 0, 0, 0],
    ]);
    const open = 100 + 200 + 300 + 400 + 500 + 600 + 3000 + 2000 + 5000.01;
    s4[0].forEach((v, i) => expect(v).toBeCloseTo([300, 700, 1100, 5000, 5000.01][i] / open, 12));
    expect(over60(d, e)).toBe(5000);
    expect(s5).toEqual([[f]]);
    expect(limitWorks(d)).toBe(false); // e sits exactly on the limit: make() never builds this
    expect(daysPastDue(d.invoices[9], asOf)).toBe('');
  });
});

describe('AP aging data', () => {
  it('every state is a plausible export with sane days and buckets', () => {
    for (const seed of MANY_SEEDS) {
      for (const d of allStates<AgingData>(apAging, seed)) {
        expect(d.invoices.length).toBeGreaterThanOrEqual(30);
        expect(d.invoices.length).toBeLessThanOrEqual(120);
        expect(new Set(d.invoices.map((i) => i.invoice)).size, 'invoice numbers are unique').toBe(d.invoices.length);
        expect([...d.vendors].sort()).toEqual([...VENDORS].sort());
        for (const i of d.invoices) {
          expect(i.date).toBeLessThan(i.due);
          expect(i.date).toBeLessThanOrEqual(d.asOf);
          expect(d.vendors).toContain(i.vendor);
          expect(Math.round(i.amount * 100) / 100).toBe(i.amount);
          const days = daysPastDue(i, d.asOf);
          if (i.paid) expect(days).toBe('');
          else expect(Number.isInteger(days) && (days as number) >= 0).toBe(true);
        }
        expect(d.invoices.some((i) => !i.paid), 'there is open AP to share out').toBe(true);
      }
    }
  });

  it('the call list is never empty, never everyone, and no balance sits on the limit', () => {
    for (const seed of MANY_SEEDS) {
      for (const d of allStates<AgingData>(apAging, seed, [4])) {
        const list = callList(d);
        expect(list.length).toBeGreaterThanOrEqual(1);
        expect(list.length).toBeLessThan(d.vendors.length);
        for (const v of d.vendors) expect(Math.abs(over60(d, v) - d.limit), `${v} vs ${d.limit}`).toBeGreaterThanOrEqual(LIMIT_MARGIN);
        expect(OVER_60_LIMITS).toContain(d.limit);
      }
    }
  });

  it.each(MANY_SEEDS)('seed %i: the limit variant always changes the call list', (seed) => {
    const base = apAging.make(new Rng(seed));
    expect(AS_OF_DATES).toContain(base.asOf);
    const step = apAging.steps[4];
    const i = step.variants.findIndex((v) => /limit/.test(v.label));
    const moved = step.variants[i].apply(base, new Rng(seed + 7919 * (i + 1)));
    expect(moved.limit).not.toBe(base.limit);
    expect(callList(moved)).not.toEqual(callList(base));
  });

  it.each(SEEDS)('seed %i: grid, shares and the call list reconcile', (seed) => {
    for (const d of allStates<AgingData>(apAging, seed)) {
      const [, , grid, shares] = apAging.steps.map((s) => s.expected(d));
      const open = total(d.invoices.filter((i) => !i.paid).map((i) => i.amount));
      expect(total(grid.flat() as number[])).toBeCloseTo(open, 6);
      expect(total(shares[0] as number[])).toBeCloseTo(1, 12);
      d.vendors.forEach((v, r) => expect((grid[r][3] as number) + (grid[r][4] as number)).toBeCloseTo(over60(d, v), 6));
    }
  });
});

// ---------- carrier cost review ----------

describe('carrier cost review data', () => {
  it('every state has every carrier, real lanes and one clear priciest lane', () => {
    for (const seed of MANY_SEEDS) {
      for (const d of allStates<CarrierData>(carrierCostReview, seed)) {
        expect(d.rows.length).toBeLessThanOrEqual(120);
        expect([...d.carriers].sort()).toEqual([...CARRIERS].sort());
        for (const c of CARRIERS) expect(d.rows.filter((r) => r.carrier === c).length, c).toBeGreaterThanOrEqual(5);
        for (const r of d.rows) {
          expect(ALL_LANES).toContain(r.lane);
          expect(r.weight).toBeGreaterThan(0);
          expect(r.cost).toBeGreaterThan(0);
          expect(r.promised).toBeGreaterThan(r.ship);
        }
        const [top, next] = laneRanking(d.rows);
        expect(top.perLb).toBeGreaterThanOrEqual(next.perLb * TOP_LANE_LEAD);
      }
    }
  });

  it.each(MANY_SEEDS)('seed %i: month edges and on-the-day deliveries make the strict comparisons wrong', (seed) => {
    const d = carrierCostReview.make(new Rng(seed));
    for (const m of REVIEW_MONTHS) {
      expect(d.rows.some((r) => r.ship === m), 'a shipment on the first day').toBe(true);
      expect(d.rows.some((r) => r.ship === eomonth(m)), 'a shipment on the last day').toBe(true);
    }
    for (const c of CARRIERS) {
      const mine = d.rows.filter((r) => r.carrier === c);
      if (mine.some((r) => r.delivered <= r.promised)) expect(mine.some((r) => r.delivered === r.promised), c).toBe(true);
    }
  });

  it.each(SEEDS)('seed %i: dividing the totals differs from averaging each shipment’s rate', (seed) => {
    const d = carrierCostReview.make(new Rng(seed));
    const averaged = d.carriers.map((c) => {
      const mine = d.rows.filter((r) => r.carrier === c);
      return total(mine.map((r) => r.cost / r.weight)) / mine.length;
    });
    const want = carrierCostReview.steps[0].expected(d).map((r) => r[0] as number);
    expect(averaged.some((v, i) => Math.abs(v - want[i]) > 1e-6)).toBe(true);
  });

  it.each(MANY_SEEDS)('seed %i: the lane variant crowns another lane and the months variant moves the window', (seed) => {
    const base = carrierCostReview.make(new Rng(seed));
    const [lane] = carrierCostReview.steps[2].variants;
    expect(laneRanking(lane.apply(base, new Rng(seed + 7919)).rows)[0].lane).not.toBe(laneRanking(base.rows)[0].lane);
    const months = carrierCostReview.steps[3].variants[0];
    expect(base.months).toEqual(REVIEW_MONTHS.slice(1));
    expect(months.apply(base, new Rng(seed)).months).toEqual(REVIEW_MONTHS.slice(0, 3));
  });
});

// ---------- weekly refresh ----------

describe('weekly refresh', () => {
  it('has no inputs and no variants, and only steps 2 and 3 write to the sheet', () => {
    const d = weeklyRefresh.make(new Rng(1));
    expect(weeklyRefresh.inputs(d)).toEqual([]);
    for (const s of weeklyRefresh.steps) expect(s.variants).toEqual([]);
    expect(weeklyRefresh.steps.map((s) => !!s.onStart)).toEqual([false, true, true]);
    expect(weeklyRefresh.steps.map((s) => !!s.startNote)).toEqual([false, true, true]);
  });

  it('asks every step for the same query output, with headers unlike the source Table', () => {
    const d = weeklyRefresh.make(new Rng(1));
    for (const s of weeklyRefresh.steps) expect(s.answer(d)).toEqual({ kind: 'query', columns: SUMMARY_COLUMNS, order: 'any' });
    const source = REFRESH_COLS.map((c) => c.header.toLowerCase()).sort().join('|');
    expect(SUMMARY_COLUMNS.map((c) => c.toLowerCase()).sort().join('|')).not.toBe(source);
    expect(weeklyRefresh.steps[0].task(d)).toMatch(/\{macPqNote\}/);
    expect(weeklyRefresh.steps[0].hints[0]).toMatch(/\{fromTable:Shipments\}/);
    for (const s of weeklyRefresh.steps.slice(1)) expect(s.task(d)).toMatch(/Data › Refresh All/);
  });

  it.each(MANY_SEEDS)('seed %i: each onStart rewrites the Table with every earlier row first, in the same column order', (seed) => {
    const d = weeklyRefresh.make(new Rng(seed));
    const block = weeklyRefresh.blocks(d)[0] as DataBlock;
    let before: Grid = block.rows;
    for (const s of weeklyRefresh.steps.slice(1)) {
      const writes = s.onStart!(d);
      expect(writes).toHaveLength(1);
      const w = writes[0];
      if (w.kind !== 'table') throw new Error('expected a Table write');
      expect(w.table).toBe('Shipments');
      expect(w.columns).toEqual(block.columns.map((c) => c.header));
      expect(w.rows.length).toBeGreaterThan(before.length);
      expect(w.rows.slice(0, before.length)).toEqual(before);
      for (const row of w.rows) expect(row).toHaveLength(w.columns.length);
      before = w.rows;
    }
  });

  it.each(MANY_SEEDS)('seed %i: the expected summary changes after each onStart and matches a regroup of the rows written', (seed) => {
    const d = weeklyRefresh.make(new Rng(seed));
    const block = weeklyRefresh.blocks(d)[0] as DataBlock;
    const grids = [block.rows, ...weeklyRefresh.steps.slice(1).map((s) => (s.onStart!(d)[0] as { rows: Grid }).rows)];
    const col = (h: string) => REFRESH_COLS.findIndex((c) => c.header === h);
    const regroup = (rows: Grid) => {
      const out = new Map<string, [number, number]>();
      for (const r of rows) {
        const [cost, n] = out.get(r[col('Carrier')] as string) ?? [0, 0];
        out.set(r[col('Carrier')] as string, [cost + (r[col('Cost')] as number), n + 1]);
      }
      return out;
    };
    const keyOf = (grid: Grid) => JSON.stringify([...grid].sort((a, b) => String(a[0]).localeCompare(String(b[0]))));
    const seen = new Set<string>();
    weeklyRefresh.steps.forEach((s, i) => {
      const expected = s.expected(d) as Grid;
      const want = regroup(grids[i]);
      expect(expected).toHaveLength(want.size);
      for (const [carrier, cost, n] of expected) {
        const [c, k] = want.get(carrier as string)!;
        expect(near(cost as number, c)).toBe(true);
        expect(n).toBe(k);
        expect(Number.isInteger(n)).toBe(true);
      }
      seen.add(keyOf(expected));
    });
    expect(seen.size, 'each step expects a different summary').toBe(3);
  });

  it.each(MANY_SEEDS)('seed %i: the new carrier appears only in step 3, and step 1 already shows all four regulars', (seed) => {
    const d = weeklyRefresh.make(new Rng(seed));
    expect(NEW_CARRIERS).toContain(d.newCarrier);
    expect(CARRIERS as readonly string[]).not.toContain(d.newCarrier);
    const carriersIn = (i: number) => weeklyRefresh.steps[i].expected(d).map((r) => r[0] as string);
    expect(carriersIn(0).sort()).toEqual([...CARRIERS].sort());
    expect(carriersIn(1)).not.toContain(d.newCarrier);
    expect(carriersIn(2)).toContain(d.newCarrier);
    expect([...d.base, ...d.week2].some((r) => r.carrier === d.newCarrier)).toBe(false);
    expect(d.week3.some((r) => r.carrier === d.newCarrier)).toBe(true);
    expect(weeklyRefresh.steps[2].task(d)).toContain(d.newCarrier);
  });

  it.each(MANY_SEEDS)('seed %i: weeks arrive in order, on weekdays, with rising Ship IDs', (seed) => {
    const d: RefreshData = weeklyRefresh.make(new Rng(seed));
    const all = rowsAtStep(d, 2);
    expect(all.length).toBeLessThanOrEqual(120);
    const ids = all.map((r) => Number(r.id.slice(3)));
    ids.slice(1).forEach((id, i) => expect(id).toBeGreaterThan(ids[i]));
    for (const r of all) {
      expect(r.id).toMatch(/^SH-\d{5}$/);
      expect(isWeekend(r.date), `${r.id} ships on a weekday`).toBe(false);
    }
    expect(Math.max(...d.base.map((r) => r.date))).toBeLessThan(Math.min(...d.week2.map((r) => r.date)));
    expect(Math.max(...d.week2.map((r) => r.date))).toBeLessThan(Math.min(...d.week3.map((r) => r.date)));
  });

  it('step notes say how many rows arrived, and the second names the new carrier', () => {
    for (const seed of SEEDS) {
      const d = weeklyRefresh.make(new Rng(seed));
      const [, s2, s3] = weeklyRefresh.steps;
      expect(s2.startNote!(d)).toContain(`${d.week2.length} rows`);
      expect(s3.startNote!(d)).toContain(`${d.week3.length} rows`);
      expect(s3.startNote!(d)).toContain(d.newCarrier);
    }
  });
});

// ---------- answer keys, recomputed independently ----------

describe('answer keys re-derived from the sheet, the way the intended formulas compute them', () => {
  it('AP aging: days, approximate-match buckets, the SUMIFS grid, shares and the FILTER', () => {
    for (const seed of MANY_SEEDS.slice(0, 60)) {
      for (const d of allStates<AgingData>(apAging, seed)) {
        const blocks = apAging.blocks(d);
        const inv = tableRows(blocks, 'Invoices');
        const cutoffs = tableRows(blocks, 'Buckets');
        const asOf = cellsAt(blocks, 'K1')[0][0] as number;
        const limit = cellsAt(blocks, 'K2')[0][0] as number;
        const vendors = cellsAt(blocks, 'J13').map((r) => r[0] as string);
        const labels = (cellsAt(blocks, 'J12')[0] as string[]).slice(1);

        const days = inv.map((r) => (r.Paid === 'Y' ? '' : Math.max(0, asOf - (r['Due date'] as number))));
        const lookup = (n: number) => {
          let best: Record<string, Cell> | undefined;
          for (const c of cutoffs) if ((c['Min days'] as number) <= n && (!best || (c['Min days'] as number) > (best['Min days'] as number))) best = c;
          return best!.Bucket as string;
        };
        const buckets = days.map((n) => (n === '' ? '' : lookup(n)));
        const grid = vendors.map((v) => labels.map((b) => total(inv.filter((r, i) => eq(r.Vendor, v) && buckets[i] !== '' && eq(buckets[i], b)).map((r) => r.Amount as number))));
        const all = total(grid.flat());
        const shares = labels.map((_, c) => total(grid.map((row) => row[c])) / all);
        const list = vendors.filter((_, r) => grid[r][3] + grid[r][4] > limit);

        const [e1, e2, e3, e4, e5] = apAging.steps.map((s) => s.expected(d));
        expect(e1).toEqual(days.map((v) => [v]));
        expect(e2).toEqual(buckets.map((v) => [v]));
        e3.forEach((row, r) => row.forEach((v, c) => expect(near(v as number, grid[r][c])).toBe(true)));
        e4[0].forEach((v, c) => expect(near(v as number, shares[c])).toBe(true));
        expect(e5).toEqual(list.map((v) => [v]));
      }
    }
  });

  it('carrier review: cost per lb, on-time share, the priciest lane and the monthly grid', () => {
    for (const seed of MANY_SEEDS.slice(0, 60)) {
      for (const d of allStates<CarrierData>(carrierCostReview, seed)) {
        const blocks = carrierCostReview.blocks(d);
        const ship = tableRows(blocks, 'Shipments');
        const carriers = cellsAt(blocks, 'I2').map((r) => r[0] as string);
        const carriers2 = cellsAt(blocks, 'I11').map((r) => r[0] as string);
        const months = cellsAt(blocks, 'J10')[0] as number[];
        const of = (c: string) => ship.filter((r) => eq(r.Carrier, c));
        const sumOf = (rows: Record<string, Cell>[], h: string) => total(rows.map((r) => r[h] as number));

        const perLb = carriers.map((c) => sumOf(of(c), 'Cost') / sumOf(of(c), 'Weight lb'));
        const onTime = carriers.map((c) => of(c).filter((r) => (r['Delivered date'] as number) <= (r['Promised date'] as number)).length / of(c).length);
        // UNIQUE, two SUMIFS per lane, then the largest.
        const lanes = [...new Set(ship.map((r) => r.Lane as string))];
        const laneCost = lanes.map((l) => sumOf(ship.filter((r) => eq(r.Lane, l)), 'Cost') / sumOf(ship.filter((r) => eq(r.Lane, l)), 'Weight lb'));
        const top = laneCost.indexOf(Math.max(...laneCost));
        const monthly = carriers2.map((c) => months.map((m) => sumOf(of(c).filter((r) => (r['Ship date'] as number) >= m && (r['Ship date'] as number) <= eomonth(m)), 'Cost')));

        const [e1, e2, e3, e4] = carrierCostReview.steps.map((s) => s.expected(d));
        e1.forEach((row, i) => expect(near(row[0] as number, perLb[i])).toBe(true));
        e2.forEach((row, i) => expect(near(row[0] as number, onTime[i])).toBe(true));
        expect(e3[0][0]).toBe(lanes[top]);
        expect(near(e3[0][1] as number, laneCost[top])).toBe(true);
        e4.forEach((row, r) => row.forEach((v, c) => expect(near(v as number, monthly[r][c])).toBe(true)));
      }
    }
  });

  it('carrier review: a window built with > or < gives a different grid', () => {
    for (const seed of SEEDS) {
      const d = carrierCostReview.make(new Rng(seed));
      for (const months of [REVIEW_MONTHS.slice(0, 3), REVIEW_MONTHS.slice(1)]) {
        const state = { ...d, months };
        const want = carrierCostReview.steps[3].expected(state) as number[][];
        const strict = d.carriers.map((c) => months.map((m) => total(d.rows.filter((r) => r.carrier === c && r.ship > m && r.ship < eomonth(m)).map((r) => r.cost))));
        expect(strict.flat().some((v, i) => !near(v, want.flat()[i]))).toBe(true);
      }
    }
  });
});

// ---------- rules ----------

describe('rules accept realistic correct formulas and catch typed answers', () => {
  const good: Record<string, string[]> = {
    'mission-ap-aging#1': [
      '=IF([@Paid]="N",MAX(0,$K$1-[@[Due date]]),"")',
      '=IF(Invoices[@Paid]="Y","",MAX($K$1-Invoices[@[Due date]],0))',
      '=IF([@Paid]="Y","",IF($K$1>[@[Due date]],$K$1-[@[Due date]],0))',
    ],
    'mission-ap-aging#2': [
      '=IF([@Paid]="Y","",_xlfn.XLOOKUP([@[Days past due]],Buckets[Min days],Buckets[Bucket],,-1))',
      '=IF([@[Days past due]]="","",VLOOKUP([@[Days past due]],Buckets,2,TRUE))',
      '=IF([@Paid]="Y","",INDEX(Buckets[Bucket],MATCH([@[Days past due]],Buckets[Min days],1)))',
    ],
    'mission-ap-aging#3': [
      '=SUMIFS(Invoices[[Amount]:[Amount]],Invoices[[Bucket]:[Bucket]],K$12,Invoices[[Vendor]:[Vendor]],$J13)',
      // Filled by copy and paste, which leaves Table column names alone.
      '=SUMIFS(Invoices[Amount],Invoices[Bucket],K$12,Invoices[Vendor],$J13)',
    ],
    'mission-ap-aging#4': [
      '=SUMIFS(Invoices[[Amount]:[Amount]],Invoices[[Bucket]:[Bucket]],K12)/SUMIFS(Invoices[[Amount]:[Amount]],Invoices[[Paid]:[Paid]],"N")',
      '=SUM(K$13:K$18)/SUM($K$13:$O$18)',
    ],
    'mission-ap-aging#5': [
      '=_xlfn._xlws.FILTER(J13:J18,(N13:N18+O13:O18)>K2)',
      '=FILTER(J13:J18,SUMIFS(Invoices[Amount],Invoices[Vendor],J13:J18,Invoices[Days past due],">60")>$K$2)',
    ],
    'mission-carrier-cost-review#1': ['=SUMIF(Shipments[Carrier],I2,Shipments[Cost])/SUMIF(Shipments[Carrier],I2,Shipments[Weight lb])'],
    'mission-carrier-cost-review#2': [
      '=SUMPRODUCT(--(Shipments[Carrier]=I2),--(Shipments[Delivered date]<=Shipments[Promised date]))/COUNTIF(Shipments[Carrier],I2)',
      '=AVERAGE(IF(Shipments[Carrier]=$I2,--(Shipments[Delivered date]<=Shipments[Promised date])))',
    ],
    'mission-carrier-cost-review#3': [
      '=_xlfn.LET(_xlpm.l,_xlfn.UNIQUE(Shipments[Lane]),_xlpm.c,SUMIFS(Shipments[Cost],Shipments[Lane],_xlpm.l)/SUMIFS(Shipments[Weight lb],Shipments[Lane],_xlpm.l),_xlfn.XLOOKUP(MAX(_xlpm.c),_xlpm.c,_xlpm.l))',
      '=LET(l,UNIQUE(Shipments[Lane]),c,SUMIFS(Shipments[Cost],Shipments[Lane],l)/SUMIFS(Shipments[Weight lb],Shipments[Lane],l),MAX(c))',
      '=LET(l,UNIQUE(Shipments[Lane]),c,SUMIFS(Shipments[Cost],Shipments[Lane],l)/SUMIFS(Shipments[Weight lb],Shipments[Lane],l),INDEX(SORTBY(HSTACK(l,c),c,-1),1,{1,2}))',
    ],
    'mission-carrier-cost-review#4': [
      '=SUMIFS(Shipments[[Cost]:[Cost]],Shipments[[Carrier]:[Carrier]],$I11,Shipments[[Ship date]:[Ship date]],">="&J$10,Shipments[[Ship date]:[Ship date]],"<"&EDATE(J$10,1))',
      '=SUMIFS(Shipments[Cost],Shipments[Carrier],$I11,Shipments[Ship date],">="&J$10,Shipments[Ship date],"<"&EDATE(J$10,1))',
    ],
  };

  it('no rule fails on the solution or common equivalents', () => {
    const failures: string[] = [];
    for (const m of [apAging, carrierCostReview] as Mission<any>[]) {
      const d = m.make(new Rng(1));
      m.steps.forEach((s, i) => {
        const id = `${m.id}#${i + 1}`;
        for (const f of [s.solution(d), ...(good[id] ?? [])]) {
          for (const item of gradeRules(s.rules, [f])) if (item.status === 'fail') failures.push(`${id} ${f}: ${item.label}`);
        }
      });
    }
    expect(failures).toEqual([]);
  });

  /** The status of one rule item: 'pass', 'fail', or undefined when the step has no such rule. */
  const rule = (m: Mission<any>, step: number, f: string, id: string) => gradeRules(m.steps[step].rules, [f]).find((i) => i.id === id)?.status;

  it('flags typed labels, thresholds, names and lanes', () => {
    const fails = (m: Mission<any>, step: number, f: string) => gradeRules(m.steps[step].rules, [f]).some((i) => i.status === 'fail');
    expect(fails(apAging, 1, '=IF([@Paid]="Y","",IF([@[Days past due]]>90,"Over 90 days",IF([@[Days past due]]>60,"61–90 days","Current")))')).toBe(true);
    // One typed label and no disallowed numbers: only the typed-text rule can catch it.
    const typedCurrent = '=IF([@Paid]="Y","",IF([@[Days past due]]=0,"Current",XLOOKUP([@[Days past due]],Buckets[Min days],Buckets[Bucket],,-1)))';
    expect(rule(apAging, 1, typedCurrent, 'no-typed-numbers')).toBe('pass');
    expect(rule(apAging, 1, typedCurrent, 'require-0')).toBe('pass');
    expect(rule(apAging, 1, typedCurrent, 'no-typed-text')).toBe('fail');
    expect(fails(apAging, 0, '=IF([@Paid]="Y","",MAX(0,46295-[@[Due date]]))')).toBe(true);
    expect(fails(apAging, 4, '=FILTER(J13:J18,N13:N18+O13:O18>5000)')).toBe(true);
    expect(fails(apAging, 2, '=SUMIFS(Invoices[Amount],Invoices[Vendor],"Apex Supply",Invoices[Bucket],K$12)')).toBe(true);
    expect(fails(carrierCostReview, 2, `="${ALL_LANES[0]}"`)).toBe(true);
    expect(fails(carrierCostReview, 0, '=SUMIFS(Shipments[Cost],Shipments[Carrier],"Northline")/SUMIFS(Shipments[Weight lb],Shipments[Carrier],"Northline")')).toBe(true);
  });

  it('require Table columns wherever an added-rows variant runs, since a fixed range covering the Table grows with it', () => {
    const fixed: [Mission<any>, number, string][] = [
      [apAging, 2, '=SUMIFS($E$2:$E$58,$A$2:$A$58,$J13,$H$2:$H$58,K$12)'],
      [carrierCostReview, 0, '=SUMIFS($G$2:$G$70,$A$2:$A$70,I2)/SUMIFS($F$2:$F$70,$A$2:$A$70,I2)'],
      [carrierCostReview, 1, '=SUMPRODUCT(($A$2:$A$70=I2)*($E$2:$E$70<=$D$2:$D$70))/COUNTIFS($A$2:$A$70,I2)'],
      [carrierCostReview, 2, '=LET(l,UNIQUE($B$2:$B$70),c,SUMIFS($G$2:$G$70,$B$2:$B$70,l)/SUMIFS($F$2:$F$70,$B$2:$B$70,l),TAKE(SORTBY(HSTACK(l,c),c,-1),1))'],
      [carrierCostReview, 3, '=SUMIFS($G$2:$G$70,$A$2:$A$70,$I11,$C$2:$C$70,">="&J$10,$C$2:$C$70,"<="&EOMONTH(J$10,0))'],
    ];
    for (const [m, step, f] of fixed) {
      expect(m.steps[step].variants.some((v) => /added/.test(v.label)), `${m.id}#${step + 1} has an added-rows variant`).toBe(true);
      expect(rule(m, step, f, 'require-0'), `${m.id}#${step + 1} ${f}`).toBe('fail');
      expect(rule(m, step, m.steps[step].solution(m.make(new Rng(1))), 'require-0')).toBe('pass');
    }
  });

  it('lock Table columns in grids filled right, because the fill handle slides Table[Col] one column per cell', () => {
    // A plain column reference like Invoices[Amount]; the locked form Invoices[[Amount]:[Amount]] doesn't match.
    const sliding = /[A-Za-z_]\w*\[(?![[@#])[^\]]+\]/;
    expect(sliding.test('=SUMIFS(Invoices[Amount],Invoices[Vendor],$J13)')).toBe(true);
    expect(sliding.test('=SUMIFS(Invoices[[Amount]:[Amount]],Invoices[[Vendor]:[Vendor]],$J13)')).toBe(false);
    let grids = 0;
    for (const m of EXTRA_MISSIONS) {
      const d = m.make(new Rng(1));
      m.steps.forEach((s, i) => {
        const a = s.answer(d);
        if (a.kind !== 'cells' || a.consistency === 'none' || rangeSize(a.range).cols < 2) return;
        grids++;
        const filled = [s.solution(d), s.hints[s.hints.length - 1]];
        for (const f of filled) expect(f, `${m.id}#${i + 1}`).not.toMatch(sliding);
      });
    }
    expect(grids).toBe(3); // AP balances and shares, carrier monthly spend
  });
});

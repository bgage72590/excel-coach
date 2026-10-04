import { describe, expect, it } from 'vitest';
import type { RangeRef } from '../src/engine/address';
import { eomonth } from '../src/engine/data';
import { gradeRules, slidingRangeHint } from '../src/engine/grade';
import { Rng } from '../src/engine/rng';
import type { Block, Cell, DataBlock, ExpectedGrid } from '../src/engine/types';
import { EXERCISES } from '../src/exercises';
import {
  AS_OF_DATES,
  AT_RISK,
  CLOSED_STAGES,
  COMMISSION_MONTHS,
  FLAG_LEVELS,
  LEVELS_PCT,
  OPEN_STAGES,
  PLAN_ACCELERATOR,
  PLAN_FLAG_BELOW,
  SALES_MISSIONS,
  TIER_MINS_PCT,
  clearOfLevels,
  commissionStatement,
  isOpenStage,
  pipelineByStage,
  slippedCount,
  statement,
  tierRateFor,
  weightedAmounts,
  type CommissionData,
  type PipelineData,
  type PlanTier,
} from '../src/missions/sales';
import { SEEDS, blockRange, missionSuite, overlaps } from './helpers/suite';

missionSuite(SALES_MISSIONS);

/**
 * The base data for a seed plus every variant of every step, seeded exactly as ExcelHost.check
 * seeds them (seed + 7919 × (variant index + 1)), so these tests see the states a learner meets.
 */
function allStates<D>(m: { make(rng: Rng): D; steps: { variants: { apply(d: D, rng: Rng): D }[] }[] }, seed: number): D[] {
  const base = m.make(new Rng(seed));
  return [base, ...m.steps.flatMap((s) => s.variants.map((v, i) => v.apply(base, new Rng(seed + 7919 * (i + 1)))))];
}

/** Session seeds are random 32-bit numbers, so invariants are checked over many of them. */
const MANY_SEEDS = [...SEEDS, ...Array.from({ length: 200 }, (_, k) => ((k + 1) * 2654435761) >>> 0)];

/** Reads a Table's rows back out of the sheet blocks, keyed by header: what Excel will hold. */
function tableRows(blocks: Block[], table: string): Record<string, Cell>[] {
  const b = blocks.find((x): x is DataBlock => x.kind === 'data' && x.table === table)!;
  return b.rows.map((r) => Object.fromEntries(b.columns.map((c, i) => [c.header, r[i]])));
}

function cellsAt(blocks: Block[], at: string): Cell[][] {
  const b = blocks.find((x) => x.kind === 'cells' && x.at === at)!;
  return b.kind === 'cells' ? b.values : [];
}

/** Excel's EOMONTH(serial, 0), worked out with Date rather than the engine helper. */
const EPOCH = Date.UTC(1899, 11, 30);
const DAY = 86_400_000;
function excelEomonth(s: number): number {
  const d = new Date(EPOCH + s * DAY);
  return Math.round((Date.UTC(d.getUTCFullYear(), d.getUTCMonth() + 1, 0) - EPOCH) / DAY);
}

/** SUMIFS and COUNTIFS text criteria ignore case. */
const same = (a: Cell, b: Cell) => String(a).toLowerCase() === String(b).toLowerCase();
const total = (rows: Record<string, Cell>[], col: string) => rows.reduce((a, r) => a + (r[col] as number), 0);

function expectClose(got: ExpectedGrid, want: Cell[][], label: string) {
  expect(got.length, `${label} rows`).toBe(want.length);
  got.forEach((row, r) =>
    row.forEach((v, c) => {
      const w = want[r][c];
      if (typeof w === 'number') expect(Math.abs((v as number) - w), `${label} [${r}][${c}]`).toBeLessThanOrEqual(1e-9 * Math.max(1, Math.abs(w)));
      else expect(v, `${label} [${r}][${c}]`).toBe(w);
    }),
  );
}

describe('sales missions: metadata and copy', () => {
  it('are sales missions with five steps, a named manager and skills that exist', () => {
    const ids = new Set(EXERCISES.map((e) => e.id));
    for (const m of SALES_MISSIONS) {
      expect(m.role).toBe('sales');
      expect(m.steps.length).toBe(5);
      expect(m.brief(m.make(new Rng(1))).from).toMatch(/^[A-Z][a-z]+ [A-Z][a-z]+, /);
      for (const s of m.skills) expect(ids.has(s), `${m.id} skill ${s}`).toBe(true);
    }
  });

  it('uses typographic apostrophes and never TODAY() in anything the learner reads', () => {
    for (const m of SALES_MISSIONS) {
      const d = m.make(new Rng(7));
      const brief = m.brief(d);
      const texts = [m.title, m.summary, brief.subject, brief.body, ...m.steps.flatMap((s) => [s.title, s.task(d), ...s.hints])];
      for (const text of texts) expect(text, text).not.toMatch(/'/);
      for (const s of m.steps) expect(s.solution(d)).not.toMatch(/TODAY\(/i);
    }
  });
});

describe('sales missions: sheet layout beyond the shared checks', () => {
  it.each(SEEDS)('seed %i: the Reps Table has room for the five statement columns plus a gap', (seed) => {
    for (const d of allStates<CommissionData>(commissionStatement, seed)) {
      const blocks = commissionStatement.blocks(d);
      const reps = blocks.find((b) => b.kind === 'data' && b.table === 'Reps')!;
      const r = blockRange(reps);
      expect(r.start).toEqual({ row: 1, col: 10 }); // J1; the learner adds L:P
      const room: RangeRef = { start: { row: 1, col: 12 }, end: { row: r.end.row + 2, col: 17 } };
      for (const b of blocks) if (b !== reps) expect(overlaps(room, blockRange(b)), `block at ${b.at}`).toBe(false);
    }
    const cols = commissionStatement.steps.map((s) => s.answer(commissionStatement.make(new Rng(seed))));
    expect(cols.map((a) => (a.kind === 'tableColumn' ? `${a.table}[${a.column}]` : a.kind))).toEqual([
      'Reps[Bookings]',
      'Reps[Attainment]',
      'Reps[Rate]',
      'Reps[Commission]',
      'Reps[Flag]',
    ]);
  });

  it.each(SEEDS)('seed %i: nothing sits below the data Tables or in the column the learner adds to Opps', (seed) => {
    // Deals spans A:E with F as its gap. Opps spans A:E, the learner adds F, and G is the gap.
    for (const [m, table, lastCol] of [
      [commissionStatement, 'Deals', 6],
      [pipelineByStage, 'Opps', 7],
    ] as const) {
      const states = allStates<any>(m, seed);
      const tallest = Math.max(...states.map((d) => blockRange(m.blocks(d).find((b) => b.kind === 'data' && b.table === table)!).end.row));
      for (const d of states) {
        // Those columns stay empty below and beside the Table at its tallest, plus the gap row.
        const zone: RangeRef = { start: { row: 1, col: 1 }, end: { row: tallest + 1, col: lastCol } };
        for (const b of m.blocks(d)) if (!(b.kind === 'data' && b.table === table)) expect(overlaps(zone, blockRange(b)), `${m.id} block at ${b.at}`).toBe(false);
      }
    }
  });
});

// ---------- commission statement ----------

const PLAN: PlanTier[] = [
  { min: 0, rate: 0.04 },
  { min: 0.5, rate: 0.06 },
  { min: 0.7, rate: 0.08 },
  { min: 1, rate: 0.1 },
  { min: 1.2, rate: 0.13 },
];

describe('commission statement: tier edges', () => {
  it('pays the tier a rep lands on exactly, and the tier below one dollar short of it', () => {
    expect(tierRateFor(PLAN, 0, 60000)).toBe(0.04);
    expect(tierRateFor(PLAN, 29950, 60000)).toBe(0.04);
    expect(tierRateFor(PLAN, 30000, 60000)).toBe(0.06); // exactly 50%
    expect(tierRateFor(PLAN, 41999, 60000)).toBe(0.06); // a dollar short of 70%
    expect(tierRateFor(PLAN, 42000, 60000)).toBe(0.08); // exactly 70%
    expect(tierRateFor(PLAN, 59999, 60000)).toBe(0.08);
    expect(tierRateFor(PLAN, 60000, 60000)).toBe(0.1); // exactly quota
    expect(tierRateFor(PLAN, 72000, 60000)).toBe(0.13); // exactly 120%
    expect(tierRateFor(PLAN, 250000, 60000)).toBe(0.13);
  });

  it('a threshold-exact quotient is the same double as the typed threshold, so Excel’s XLOOKUP sees an exact match', () => {
    const typed = [0, 0.5, 0.7, 1, 1.2];
    for (let quota = 60000; quota <= 120000; quota += 5000) {
      TIER_MINS_PCT.forEach((pct, i) => {
        const bookings = (pct * quota) / 100;
        expect(Number.isInteger(bookings)).toBe(true);
        expect(bookings / quota).toBe(typed[i]);
        expect(bookings / quota).toBe(pct / 100);
      });
    }
  });

  it('reads attainment levels in whole numbers', () => {
    expect(clearOfLevels(42000, 60000)).toBe(false); // exactly 70%
    expect(clearOfLevels(42000, 60000, true)).toBe(true);
    expect(clearOfLevels(41700, 60000)).toBe(false); // 69.5%: too close to read
    expect(clearOfLevels(41400, 60000)).toBe(true); // 69%
    expect(clearOfLevels(60000 * 1.13, 60000)).toBe(true);
  });
});

describe('commission statement: bookings, accelerator and clawbacks', () => {
  const M = COMMISSION_MONTHS[1];
  const d: CommissionData = {
    reps: [
      { rep: 'Ava Patel', quota: 60000 },
      { rep: 'Ben Ortiz', quota: 50000 },
      { rep: 'Chloe Kim', quota: 80000 },
    ],
    deals: [
      { id: 'D-1', rep: 'Ava Patel', date: M - 1, amount: 9999, status: 'Won' }, // last day of the month before
      { id: 'D-2', rep: 'Ava Patel', date: M, amount: 40000, status: 'Won' }, // first day
      { id: 'D-3', rep: 'Ava Patel', date: eomonth(M), amount: 35000, status: 'Won' }, // last day
      { id: 'D-4', rep: 'Ava Patel', date: M + 9, amount: 5000, status: 'Cancelled' }, // clawed back
      { id: 'D-5', rep: 'Ava Patel', date: eomonth(M) + 1, amount: 7000, status: 'Cancelled' }, // next month's statement
      { id: 'D-6', rep: 'Ben Ortiz', date: M + 3, amount: 35000, status: 'Won' }, // exactly 70%
      { id: 'D-7', rep: 'Ben Ortiz', date: M + 4, amount: 2000, status: 'Cancelled' },
      { id: 'D-8', rep: 'Chloe Kim', date: M + 5, amount: 30000, status: 'Won' }, // 37.5%
      { id: 'D-9', rep: 'Chloe Kim', date: M - 3, amount: 4000, status: 'Cancelled' }, // last month's statement
    ],
    tiers: PLAN,
    month: M,
    accelerator: 1.5,
    flagBelow: 0.7,
  };

  it('counts Won deals in the month, first and last day included', () => {
    expect(statement(d).map((s) => s.bookings)).toEqual([75000, 35000, 30000]);
    expect(statement(d).map((s) => s.cancelled)).toEqual([5000, 2000, 0]);
  });

  it('pays the accelerator only above quota and claws back this month’s cancellations at the rep’s rate', () => {
    const [ava, ben, chloe] = statement(d);
    // Ava: 125% → 13%. $60,000 at 13%, $15,000 at 1.5 × 13%, minus $5,000 at 13%.
    expect(ava.rate).toBe(0.13);
    expect(ava.commission).toBeCloseTo(0.13 * 60000 + 1.5 * 0.13 * 15000 - 0.13 * 5000, 6);
    // Ben: exactly 70% earns the 70% tier, no accelerator.
    expect(ben.rate).toBe(0.08);
    expect(ben.commission).toBeCloseTo(0.08 * (35000 - 2000), 6);
    // Chloe: last month's cancellation stays on last month's statement.
    expect(chloe.commission).toBeCloseTo(0.04 * 30000, 6);
    expect(statement({ ...d, accelerator: 2 })[0].commission).toBeCloseTo(0.13 * (60000 + 2 * 15000 - 5000), 6);
  });

  it('flags only reps strictly below the level', () => {
    expect(statement(d).map((s) => s.flag)).toEqual(['', '', AT_RISK]);
    expect(statement({ ...d, flagBelow: 0.75 }).map((s) => s.flag)).toEqual(['', AT_RISK, AT_RISK]);
  });

  it('moves the clawback with the month', () => {
    const sept = statement({ ...d, month: COMMISSION_MONTHS[2] });
    expect(sept[0].cancelled).toBe(7000);
    expect(sept[0].bookings).toBe(0);
  });
});

describe('commission statement data', () => {
  it('every state keeps whole-dollar deals, readable attainment and a positive commission', () => {
    for (const seed of MANY_SEEDS) {
      for (const d of allStates<CommissionData>(commissionStatement, seed)) {
        const ids = d.deals.map((x) => x.id);
        expect(new Set(ids).size, 'deal ids are unique').toBe(ids.length);
        expect(new Set(d.reps.map((r) => r.rep)).size).toBe(d.reps.length);
        for (const x of d.deals) {
          expect(Number.isInteger(x.amount) && x.amount % 50 === 0 && x.amount >= 1000, `${x.id} ${x.amount}`).toBe(true);
          expect(x.date).toBeGreaterThanOrEqual(COMMISSION_MONTHS[0]);
          expect(x.date).toBeLessThanOrEqual(eomonth(COMMISSION_MONTHS[COMMISSION_MONTHS.length - 1]));
        }
        const lines = statement(d);
        d.reps.forEach((r, i) => {
          const won = d.deals.filter((x) => x.rep === r.rep && x.status === 'Won' && x.date >= d.month && x.date <= eomonth(d.month));
          expect(won.length, `${r.rep} has wins this month`).toBeGreaterThanOrEqual(2);
          expect(clearOfLevels(lines[i].bookings, r.quota, true), `${r.rep} attainment ${lines[i].attainment}`).toBe(true);
          expect(lines[i].commission).toBeGreaterThan(0);
        });
        const rates = d.tiers.map((t) => t.rate);
        expect(rates).toEqual([...rates].sort((a, b) => a - b));
        expect(new Set(rates).size).toBe(rates.length);
      }
    }
  });

  it.each(MANY_SEEDS.slice(0, 80))('seed %i: every month has a rep exactly on 70%, one over quota, one at risk and clawbacks', (seed) => {
    const d = commissionStatement.make(new Rng(seed));
    expect(d.accelerator).toBe(PLAN_ACCELERATOR);
    expect(d.flagBelow).toBe(PLAN_FLAG_BELOW);
    for (const month of COMMISSION_MONTHS) {
      const lines = statement({ ...d, month });
      expect(d.reps.some((r, i) => 100 * lines[i].bookings === 70 * r.quota), 'exactly 70%').toBe(true);
      expect(d.reps.some((r, i) => lines[i].bookings > r.quota), 'over quota').toBe(true);
      expect(lines.some((l) => l.flag === AT_RISK), 'at risk').toBe(true);
      expect(lines.filter((l) => l.cancelled > 0).length).toBeGreaterThanOrEqual(3);
      expect(d.deals.some((x) => x.date === month && x.status === 'Won'), 'a win on the first day').toBe(true);
      expect(d.deals.some((x) => x.date === eomonth(month) && x.status === 'Won'), 'a win on the last day').toBe(true);
      // The rep on 70% isn't flagged and earns the 70% tier.
      const edge = d.reps.findIndex((r, i) => 100 * lines[i].bookings === 70 * r.quota);
      expect(lines[edge].flag).toBe('');
      expect(lines[edge].rate).toBe(d.tiers[TIER_MINS_PCT.indexOf(70)].rate);
    }
  });

  it.each(MANY_SEEDS.slice(0, 60))('seed %i: each variant changes what it claims to', (seed) => {
    const base = commissionStatement.make(new Rng(seed));
    const before = statement(base);
    const apply = (step: number, label: RegExp) => {
      const v = commissionStatement.steps[step].variants.find((x) => label.test(x.label))!;
      return v.apply(base, new Rng(seed + 1));
    };
    const changed = (a: unknown[], b: unknown[]) => JSON.stringify(a) !== JSON.stringify(b);
    expect(changed(statement(apply(2, /tier/)).map((s) => s.rate), before.map((s) => s.rate)), 'tier rate').toBe(true);
    expect(changed(statement(apply(3, /accelerator/)).map((s) => s.commission), before.map((s) => s.commission)), 'accelerator').toBe(true);
    expect(changed(statement(apply(4, /flag level/)).map((s) => s.flag), before.map((s) => s.flag)), 'flag level').toBe(true);
    expect(changed(statement(apply(0, /month/)).map((s) => s.bookings), before.map((s) => s.bookings)), 'month').toBe(true);
    expect(changed(statement(apply(1, /quotas/)).map((s) => s.attainment), before.map((s) => s.attainment)), 'quotas').toBe(true);
    const added = apply(0, /new rep/);
    expect(added.reps.length).toBe(base.reps.length + 1);
    expect(statement(added).at(-1)!.bookings).toBeGreaterThan(0);
  });
});

// ---------- pipeline by stage ----------

describe('pipeline: slipped deal boundaries', () => {
  const asOf = AS_OF_DATES[1];
  const d: PipelineData = {
    opps: [
      { id: 'OPP-1', rep: 'Ava Patel', stage: 'Proposal', date: asOf, amount: 10000 }, // closes on the as-of date: not slipped
      { id: 'OPP-2', rep: 'Ava Patel', stage: 'Discovery', date: asOf - 1, amount: 20000 }, // a day late: slipped
      { id: 'OPP-3', rep: 'Ava Patel', stage: 'Closed Won', date: asOf - 20, amount: 30000 }, // closed, so not slipped
      { id: 'OPP-4', rep: 'Ava Patel', stage: 'Closed Lost', date: asOf - 2, amount: 40000 },
      { id: 'OPP-5', rep: 'Ava Patel', stage: 'Negotiation', date: asOf + 30, amount: 50000 },
      { id: 'OPP-6', rep: 'Ben Ortiz', stage: 'Prospecting', date: asOf - 40, amount: 60000 },
    ],
    stages: OPEN_STAGES.map((stage, i) => ({ stage, probability: [0.1, 0.25, 0.5, 0.75][i] })),
    asOf,
    stageRows: [...OPEN_STAGES],
    gridStages: [...OPEN_STAGES],
    reps: [
      { rep: 'Ava Patel', quota: 200000 },
      { rep: 'Ben Ortiz', quota: 150000 },
    ],
  };

  it('counts open deals dated before the as-of date, never on it, never closed ones', () => {
    expect(slippedCount(d, 'Ava Patel')).toBe(1);
    expect(slippedCount(d, 'Ben Ortiz')).toBe(1);
    expect(slippedCount({ ...d, asOf: asOf + 1 }, 'Ava Patel')).toBe(2); // OPP-1 slips the next day
    expect(slippedCount({ ...d, asOf: asOf - 1 }, 'Ava Patel')).toBe(0);
  });

  it('weighs closed deals at 0 because they aren’t in the Stages Table', () => {
    expect(weightedAmounts(d)).toEqual([10000 * 0.5, 20000 * 0.25, 0, 0, 50000 * 0.75, 60000 * 0.1]);
  });
});

describe('pipeline data', () => {
  it('every state keeps unique ids, round amounts and closed deals in the past', () => {
    for (const seed of MANY_SEEDS) {
      for (const d of allStates<PipelineData>(pipelineByStage, seed)) {
        const ids = d.opps.map((o) => o.id);
        expect(new Set(ids).size).toBe(ids.length);
        expect(d.opps.length).toBeLessThanOrEqual(120);
        expect(d.reps.length).toBe(6);
        for (const o of d.opps) {
          expect([...OPEN_STAGES, ...CLOSED_STAGES]).toContain(o.stage);
          expect(o.amount % 500).toBe(0);
          expect(d.reps.map((r) => r.rep)).toContain(o.rep);
        }
        for (const r of d.reps) expect(r.quota).toBeGreaterThan(0);
      }
    }
  });

  it.each(MANY_SEEDS.slice(0, 80))('seed %i: base data has boundary deals for every as-of date and every stage in use', (seed) => {
    const d = pipelineByStage.make(new Rng(seed));
    expect(AS_OF_DATES).toContain(d.asOf);
    for (const day of AS_OF_DATES) {
      expect(d.opps.some((o) => isOpenStage(o.stage) && o.date === day), 'open deal on the as-of date').toBe(true);
      expect(d.opps.some((o) => isOpenStage(o.stage) && o.date === day - 1), 'open deal the day before').toBe(true);
    }
    for (const o of d.opps) if (!isOpenStage(o.stage)) expect(o.date).toBeLessThan(Math.min(...AS_OF_DATES));
    expect(d.opps.some((o) => !isOpenStage(o.stage))).toBe(true);
    for (const s of OPEN_STAGES) expect(d.opps.some((o) => o.stage === s), s).toBe(true);
    for (const r of d.reps) expect(d.opps.some((o) => o.rep === r.rep && isOpenStage(o.stage)), r.rep).toBe(true);
  });

  it.each(MANY_SEEDS.slice(0, 60))('seed %i: moving the as-of date or closing deals changes the slipped counts', (seed) => {
    const base = pipelineByStage.make(new Rng(seed));
    const counts = (d: PipelineData) => d.reps.map((r) => slippedCount(d, r.rep));
    const step = pipelineByStage.steps[4];
    for (const label of [/as-of/, /new stages/, /close dates/]) {
      const v = step.variants.find((x) => label.test(x.label))!;
      expect(counts(v.apply(base, new Rng(seed + 3))), v.label).not.toEqual(counts(base));
    }
  });
});

// ---------- answer keys re-derived from the sheet ----------

describe('answer keys re-derived from the sheet, the way the intended formulas compute them', () => {
  it('commission statement: SUMIFS window, division, XLOOKUP -1, accelerator with clawback, flag', () => {
    for (const seed of MANY_SEEDS.slice(0, 60)) {
      for (const d of allStates<CommissionData>(commissionStatement, seed)) {
        const blocks = commissionStatement.blocks(d);
        const deals = tableRows(blocks, 'Deals');
        const reps = tableRows(blocks, 'Reps');
        const tiers = tableRows(blocks, 'Tiers');
        const [[month], [accelerator], [level]] = [cellsAt(blocks, 'H1')[0], cellsAt(blocks, 'H2')[0], cellsAt(blocks, 'H3')[0]] as number[][];
        const end = excelEomonth(month);
        const sumifs = (rep: Cell, status: string) =>
          total(
            deals.filter((x) => same(x.Rep, rep) && same(x.Status, status) && (x['Close date'] as number) >= month && (x['Close date'] as number) <= end),
            'Amount',
          );
        const rows = reps.map((r) => {
          const quota = r.Quota as number;
          const bookings = sumifs(r.Rep, 'Won');
          const attainment = bookings / quota;
          // XLOOKUP(…, -1) on doubles, as Excel compares them: the largest threshold at or below.
          const tier = tiers.filter((t) => (t['Min attainment'] as number) <= attainment).sort((a, b) => (b['Min attainment'] as number) - (a['Min attainment'] as number))[0];
          const rate = tier.Rate as number;
          const cancelled = sumifs(r.Rep, 'Cancelled');
          const commission = rate * (Math.min(bookings, quota) + accelerator * Math.max(0, bookings - quota) - cancelled);
          return [bookings, attainment, rate, commission, attainment < level ? 'At risk' : ''];
        });
        commissionStatement.steps.forEach((s, i) =>
          expectClose(
            s.expected(d),
            rows.map((r) => [r[i]]),
            `seed ${seed} step ${i + 1}`,
          ),
        );
      }
    }
  });

  it('pipeline: SUMIFS and COUNTIFS by stage, weighted lookup, grid, coverage and slipped COUNTIFS', () => {
    for (const seed of MANY_SEEDS.slice(0, 60)) {
      for (const d of allStates<PipelineData>(pipelineByStage, seed)) {
        const blocks = pipelineByStage.blocks(d);
        const opps = tableRows(blocks, 'Opps');
        const stages = tableRows(blocks, 'Stages');
        const asOf = cellsAt(blocks, 'I1')[0][0] as number;
        const stageRows = cellsAt(blocks, 'K2').map((r) => r[0]);
        const gridStages = cellsAt(blocks, 'L8')[0];
        const reps = cellsAt(blocks, 'K9').map((r) => r[0]);
        const quotas = cellsAt(blocks, 'P9').map((r) => r[0] as number);

        const summary = stageRows.map((s) => {
          const mine = opps.filter((o) => same(o.Stage, s));
          return [total(mine, 'Amount'), mine.length];
        });
        const weighted = opps.map((o) => {
          const hit = stages.find((s) => same(s.Stage, o.Stage));
          return [(o.Amount as number) * (hit ? (hit.Probability as number) : 0)];
        });
        const grid = reps.map((r) => gridStages.map((s) => total(opps.filter((o) => same(o.Rep, r) && same(o.Stage, s)), 'Amount')));
        const coverage = grid.map((row, i) => [row.reduce((a, v) => a + v, 0) / quotas[i]]);
        const slipped = reps.map((r) => [
          opps.filter((o) => same(o.Rep, r) && (o['Close date'] as number) < asOf && !same(o.Stage, 'Closed Won') && !same(o.Stage, 'Closed Lost')).length,
        ]);

        const [e1, e2, e3, e4, e5] = pipelineByStage.steps.map((s) => s.expected(d));
        expect(e1).toEqual(summary);
        expect(e2).toEqual(weighted);
        expect(e3).toEqual(grid);
        expectClose(e4, coverage, `seed ${seed} coverage`);
        expect(e5).toEqual(slipped);
      }
    }
  });
});

describe('rules accept realistic correct formulas', () => {
  const window = 'Deals[Close date],">="&$H$1,Deals[Close date],"<="&EOMONTH($H$1,0)';
  const good: Record<string, string[]> = {
    'mission-commission-statement#1': [
      `=SUMIFS(Deals[Amount],Deals[Rep],Reps[@Rep],Deals[Status],"Won",${window})`,
      '=SUMIFS(Deals[Amount],Deals[Rep],[@Rep],Deals[Close date],">="&$H$1,Deals[Close date],"<"&EDATE($H$1,1),Deals[Status],"Won")',
    ],
    'mission-commission-statement#2': ['=Reps[@Bookings]/Reps[@Quota]'],
    'mission-commission-statement#3': [
      '=_xlfn.XLOOKUP([@Attainment],Tiers[Min attainment],Tiers[Rate],,-1)',
      '=VLOOKUP([@Attainment],Tiers,2,TRUE)',
      '=INDEX(Tiers[Rate],MATCH([@Attainment],Tiers[Min attainment],1))',
      '=LOOKUP([@Attainment],Tiers[Min attainment],Tiers[Rate])',
    ],
    'mission-commission-statement#4': [
      `=[@Rate]*(MIN([@Bookings],[@Quota])+$H$2*MAX(0,[@Bookings]-[@Quota]))-[@Rate]*SUMIFS(Deals[Amount],Deals[Rep],[@Rep],Deals[Status],"Cancelled",${window})`,
      `=[@Rate]*(IF([@Bookings]>[@Quota],[@Quota]+$H$2*([@Bookings]-[@Quota]),[@Bookings])-SUMIFS(Deals[Amount],Deals[Rep],[@Rep],Deals[Status],"Cancelled",${window}))`,
    ],
    'mission-commission-statement#5': ['=IF(Reps[@Attainment]<$H$3,"At risk","")'],
    'mission-pipeline-by-stage#1': ['=SUMIFS(Opps[Amount],Opps[Stage],K2)', '=COUNTIFS(Opps[Stage],K2)', '=SUMIF(Opps[Stage],$K2,Opps[Amount])'],
    'mission-pipeline-by-stage#2': [
      '=IFERROR(XLOOKUP([@Stage],Stages[Stage],Stages[Probability]),0)*[@Amount]',
      '=[@Amount]*IFNA(VLOOKUP([@Stage],Stages,2,FALSE),0)',
      '=[@Amount]*_xlfn.XLOOKUP([@Stage],Stages[Stage],Stages[Probability],0)',
    ],
    'mission-pipeline-by-stage#3': [
      // Filled down, then copied and pasted across: paste keeps Table column names.
      '=SUMIFS(Opps[Amount],Opps[Stage],L$8,Opps[Rep],$K9)',
      '=SUMIFS(Opps[[#Data],[Amount]],Opps[[#Data],[Rep]],$K9,Opps[[#Data],[Stage]],L$8)',
    ],
    'mission-pipeline-by-stage#4': ['=SUM($L9:$O9)/$P9', '=SUMIFS(Opps[Amount],Opps[Rep],K9,Opps[Stage],"<>Closed Won",Opps[Stage],"<>Closed Lost")/P9'],
    'mission-pipeline-by-stage#5': [
      '=COUNTIFS(Opps[Rep],K9,Opps[Stage],"<>Closed*",Opps[Close date],"<"&$I$1)',
      '=SUMPRODUCT((Opps[Rep]=$K9)*(Opps[Close date]<$I$1)*ISNUMBER(MATCH(Opps[Stage],Stages[Stage],0)))',
    ],
  };

  it('no rule fails on the solution or common equivalents', () => {
    const failures: string[] = [];
    for (const m of SALES_MISSIONS) {
      const d = m.make(new Rng(1));
      m.steps.forEach((s, i) => {
        const id = `${m.id}#${i + 1}`;
        const solutions = s
          .solution(d)
          .split(/\s{2,}/)
          .map((f) => f.replace(/^[A-Z]+\d+:\s*/, ''));
        for (const f of [...solutions, ...(good[id] ?? [])]) {
          for (const item of gradeRules(s.rules, [f])) if (item.status === 'fail') failures.push(`${id} ${f}: ${item.label}`);
        }
      });
    }
    expect(failures).toEqual([]);
  });

  it('typed-in plan numbers and names are caught', () => {
    const caught = (id: string, f: string) => {
      const [mid, n] = id.split('#');
      const m = SALES_MISSIONS.find((x) => x.id === mid)!;
      return gradeRules(m.steps[Number(n) - 1].rules, [f]).some((item) => item.status === 'fail');
    };
    expect(caught('mission-commission-statement#1', `=SUMIFS(Deals[Amount],Deals[Rep],"Ava Patel",Deals[Status],"Won",${window})`)).toBe(true);
    expect(caught('mission-commission-statement#4', '=[@Rate]*(MIN([@Bookings],[@Quota])+1.5*MAX(0,[@Bookings]-[@Quota]))')).toBe(true);
    expect(caught('mission-commission-statement#5', '=IF([@Attainment]<0.7,"At risk","")')).toBe(true);
    expect(caught('mission-commission-statement#5', '=IF([@Attainment]<70%,"At risk","")')).toBe(true);
    expect(caught('mission-pipeline-by-stage#2', '=[@Amount]*IF([@Stage]="Proposal",0.5,0)')).toBe(true);
    // L9's plain formula dragged one column right: every Table column has moved over one.
    expect(caught('mission-pipeline-by-stage#3', '=SUMIFS(Opps[Weighted],Opps[Stage],$K9,Opps[Close date],M$8)')).toBe(true);
    expect(caught('mission-pipeline-by-stage#5', '=COUNTIFS(Opps[Rep],K9,Opps[Close date],"<"&DATE(2026,9,21))')).toBe(true);
  });
});

describe('solutions survive the fill the task asks for', () => {
  it('the rep × stage grid locks every Table column, since filling right shifts plain ones', () => {
    const f = pipelineByStage.steps[2].solution(pipelineByStage.make(new Rng(1)));
    expect(f).toMatch(/\$K9/);
    expect(f).toMatch(/L\$8/);
    expect(f).not.toMatch(/Opps\[(?!\[)/);
  });

  it('no solution filled through a range draws the sliding-range advice', () => {
    for (const m of SALES_MISSIONS) {
      const d = m.make(new Rng(1));
      m.steps.forEach((s, i) => {
        const e = s.expected(d);
        const formulas = s
          .solution(d)
          .split(/\s{2,}/)
          .map((f) => f.replace(/^[A-Z]+\d+:\s*/, ''));
        expect(slidingRangeHint(formulas, e.length * (e[0]?.length ?? 1)), `${m.id} step ${i + 1}`).toBeUndefined();
      });
    }
  });
});

describe('levels the data is built around', () => {
  it('include every tier threshold and every flag level', () => {
    for (const p of TIER_MINS_PCT.filter((x) => x > 0)) expect(LEVELS_PCT).toContain(p);
    for (const l of [PLAN_FLAG_BELOW, ...FLAG_LEVELS]) expect(LEVELS_PCT).toContain(Math.round(l * 100));
  });
});

import { describe, expect, it } from 'vitest';
import { eomonth, monthStart } from '../src/engine/data';
import { Rng } from '../src/engine/rng';
import { EXERCISES } from '../src/exercises';
import { missionStepExercise } from '../src/missions/compile';
import {
  BVA_MONTHS,
  DOLLAR_MARGIN,
  FINANCE_MISSIONS,
  PCT_MARGIN,
  REVIEW_LIMITS,
  actuals,
  budgetVsActual,
  flagOutcome,
  onTimeRate,
  ranking,
  ratesAreDistinct,
  reviewFlags,
  vendorScorecard,
  vendorTally,
  type BvaData,
  type ScoreData,
} from '../src/missions/finance';
import { missionSuite } from './helpers/suite';

missionSuite(FINANCE_MISSIONS);

const MANY_SEEDS = Array.from({ length: 120 }, (_, i) => i * 7919 + 13);

describe('finance missions', () => {
  it('have four steps each and only list skills that exist', () => {
    const ids = new Set(EXERCISES.map((e) => e.id));
    for (const m of FINANCE_MISSIONS) {
      expect(m.steps.length).toBe(4);
      for (const s of m.skills) expect(ids.has(s), `${m.id} skill ${s}`).toBe(true);
    }
  });
});

// ---------- budget vs actual ----------

/** Every data set a step of the budget mission can be graded on: the base data and each variant. */
function bvaSets(seed: number): BvaData[] {
  const base = budgetVsActual.make(new Rng(seed));
  const out = [base];
  budgetVsActual.steps.forEach((s, i) => s.variants.forEach((v, k) => out.push(v.apply(base, new Rng(seed * 31 + i * 7 + k)))));
  return out;
}

/** The flag step's data sets, where the margin guarantee has to hold. */
function flagSets(seed: number): BvaData[] {
  const base = budgetVsActual.make(new Rng(seed));
  const step = budgetVsActual.steps[3];
  return [base, ...step.variants.map((v, k) => v.apply(base, new Rng(seed * 31 + k)))];
}

describe('budget vs actual data', () => {
  it.each(MANY_SEEDS)('seed %i: GL holds 15–120 lines across three months and every budget is positive', (seed) => {
    for (const d of bvaSets(seed)) {
      expect(d.lines.length).toBeGreaterThanOrEqual(15);
      expect(d.lines.length).toBeLessThanOrEqual(120);
      for (const l of d.lines) expect(BVA_MONTHS).toContain(monthStart(l.date));
      for (const a of d.accounts) expect(a.budget).toBeGreaterThan(0);
      expect(BVA_MONTHS).toContain(d.month);
    }
  });

  it.each(MANY_SEEDS)('seed %i: every account-month variance stays clear of every review limit', (seed) => {
    for (const d of flagSets(seed)) {
      for (const month of BVA_MONTHS) {
        const act = actuals({ ...d, month });
        d.accounts.forEach((a, i) => {
          const dollars = Math.abs(act[i] - a.budget);
          const pct = dollars / a.budget;
          for (const l of REVIEW_LIMITS) {
            expect(Math.abs(pct - l.pct), `${a.name} pct vs ${l.pct}`).toBeGreaterThan(PCT_MARGIN - 1e-6);
            expect(Math.abs(dollars - l.amount), `${a.name} $ vs ${l.amount}`).toBeGreaterThan(DOLLAR_MARGIN - 1e-6);
          }
        });
      }
    }
  });

  it.each(MANY_SEEDS)('seed %i: margins also hold for the seeds the checker uses (seed + 7919 × (variant + 1))', (seed) => {
    const base = budgetVsActual.make(new Rng(seed));
    const sets = [base, ...budgetVsActual.steps[3].variants.map((v, i) => v.apply(base, new Rng(seed + 7919 * (i + 1))))];
    for (const d of sets) {
      for (const month of BVA_MONTHS) {
        const act = actuals({ ...d, month });
        d.accounts.forEach((a, i) => {
          const dollars = Math.abs(act[i] - a.budget);
          for (const l of REVIEW_LIMITS) {
            expect(Math.abs(dollars / a.budget - l.pct)).toBeGreaterThan(PCT_MARGIN - 1e-6);
            expect(Math.abs(dollars - l.amount)).toBeGreaterThan(DOLLAR_MARGIN - 1e-6);
          }
        });
      }
    }
  });

  it.each(MANY_SEEDS)('seed %i: every month has a GL line on its first day and on its last day', (seed) => {
    const d = budgetVsActual.make(new Rng(seed));
    for (const m of BVA_MONTHS) {
      expect(d.lines.some((l) => l.date === m)).toBe(true);
      expect(d.lines.some((l) => l.date === eomonth(m))).toBe(true);
    }
  });

  it.each(MANY_SEEDS)('seed %i: each month mixes flagged over, flagged under and single-limit accounts', (seed) => {
    const d = budgetVsActual.make(new Rng(seed));
    for (const month of BVA_MONTHS) {
      const act = actuals({ ...d, month });
      const seen = new Set(d.accounts.map((a, i) => flagOutcome(a.budget, act[i] - a.budget, REVIEW_LIMITS[0])));
      for (const o of ['over', 'under', 'pctOnly', 'dollarOnly']) expect(seen.has(o as never), `${month} has ${o}`).toBe(true);
    }
  });

  it('computes actuals the way SUMIFS with an EOMONTH window does', () => {
    const d: BvaData = {
      accounts: [
        { name: 'Rent', budget: 40000 },
        { name: 'Travel', budget: 5000 },
      ],
      lines: [
        { date: BVA_MONTHS[1] - 1, account: 'Rent', dept: 'Operations', amount: 999 }, // last day of the month before
        { date: BVA_MONTHS[1], account: 'Rent', dept: 'Operations', amount: 30000 }, // first day
        { date: eomonth(BVA_MONTHS[1]), account: 'Rent', dept: 'Warehouse', amount: 15000.5 }, // last day
        { date: eomonth(BVA_MONTHS[1]) + 1, account: 'Rent', dept: 'Operations', amount: 777 }, // first day of next month
        { date: BVA_MONTHS[1] + 3, account: 'Travel', dept: 'Sales', amount: 4200 },
        { date: BVA_MONTHS[1] + 4, account: 'Utilities', dept: 'Sales', amount: 123 }, // not budgeted
      ],
      month: BVA_MONTHS[1],
      pct: 0.1,
      limit: 5000,
    };
    expect(actuals(d)).toEqual([45000.5, 4200]);
    // Rent: +5,000.50 is 12.5% and over $5,000; Travel: -800 is 16% but under $5,000.
    expect(reviewFlags(d)).toEqual(['Review', '']);
    expect(reviewFlags({ ...d, accounts: [{ name: 'Rent', budget: 50000 }, d.accounts[1]] })).toEqual(['', '']); // -9.999% stays below 10%
  });

  it('flags under-budget accounts too, and needs both limits', () => {
    expect(flagOutcome(200000, -24000, REVIEW_LIMITS[0])).toBe('under');
    expect(flagOutcome(200000, 8000, REVIEW_LIMITS[0])).toBe('dollarOnly');
    expect(flagOutcome(3000, 900, REVIEW_LIMITS[0])).toBe('pctOnly');
    expect(flagOutcome(3000, 100, REVIEW_LIMITS[0])).toBe('within');
  });

  it.each(MANY_SEEDS.slice(0, 40))('seed %i: the flag step shows both Review and blank, and the limits variant changes a flag', (seed) => {
    const ex = missionStepExercise(budgetVsActual, 3);
    const base = budgetVsActual.make(new Rng(seed));
    const flags = ex.expected(base).map((r) => r[0]);
    expect(flags).toContain('Review');
    expect(flags).toContain('');
    const limits = budgetVsActual.steps[3].variants.find((v) => /limits/.test(v.label))!;
    const changed = limits.apply(base, new Rng(seed));
    expect(changed.pct === base.pct && changed.limit === base.limit).toBe(false);
  });
});

// ---------- vendor scorecard ----------

function scoreSets(seed: number): ScoreData[] {
  const base = vendorScorecard.make(new Rng(seed));
  const out = [base];
  vendorScorecard.steps.forEach((s, i) => s.variants.forEach((v, k) => out.push(v.apply(base, new Rng(seed * 31 + i * 7 + k)))));
  return out;
}

describe('vendor scorecard data', () => {
  it.each(MANY_SEEDS)('seed %i: no two vendors tie on on-time rate, in the base data or any variant', (seed) => {
    for (const d of scoreSets(seed)) {
      expect(ratesAreDistinct(d.rows)).toBe(true);
      const rates = d.vendors.map((v) => onTimeRate(d.rows, v));
      expect(new Set(rates).size).toBe(rates.length);
    }
  });

  it.each(MANY_SEEDS)('seed %i: every vendor has deliveries, and an on-the-day delivery makes < differ from <=', (seed) => {
    const d = vendorScorecard.make(new Rng(seed));
    expect(d.rows.length).toBeGreaterThanOrEqual(15);
    expect(d.rows.length).toBeLessThanOrEqual(120);
    for (const v of d.vendors) {
      const t = vendorTally(d.rows, v);
      expect(t.count).toBeGreaterThan(0);
      expect(t.qty).toBeGreaterThan(0);
      if (t.onTime > 0) expect(d.rows.some((r) => r.vendor === v && r.delivered === r.promised)).toBe(true);
    }
    expect(d.rows.some((r) => r.delivered === r.promised)).toBe(true);
  });

  it('ranks by on-time rate, highest first', () => {
    const d: ScoreData = {
      profiles: [],
      vendors: ['A', 'B', 'C'],
      rows: [
        { vendor: 'A', promised: 10, delivered: 11, qty: 100, defects: 1 },
        { vendor: 'A', promised: 10, delivered: 10, qty: 100, defects: 0 },
        { vendor: 'B', promised: 10, delivered: 9, qty: 50, defects: 2 },
        { vendor: 'C', promised: 10, delivered: 12, qty: 80, defects: 0 },
        { vendor: 'C', promised: 10, delivered: 13, qty: 80, defects: 0 },
        { vendor: 'C', promised: 10, delivered: 10, qty: 80, defects: 1 },
      ],
    };
    expect(onTimeRate(d.rows, 'A')).toBe(0.5); // delivered on the promised day counts
    expect(ranking(d)).toEqual([
      ['B', 1],
      ['A', 0.5],
      ['C', 1 / 3],
    ]);
  });
});

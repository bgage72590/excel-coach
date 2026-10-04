import { describe, expect, it } from 'vitest';
import { gradeRules } from '../src/engine/grade';
import { Rng } from '../src/engine/rng';
import { gradeSheetCheck } from '../src/engine/sheetChecks';
import type { Inspection, Rules } from '../src/engine/types';
import { MODELING, SCENARIOS } from '../src/exercises/modeling';
import { missionStepExercise } from '../src/missions/compile';
import { FORECAST_MISSIONS, opsForecast, opsForecastReview, opsSummary, type OpsForecastData } from '../src/missions/forecast';
import { SEEDS, missionSuite } from './helpers/suite';

missionSuite(FORECAST_MISSIONS);

/** True when every rule passes for the formula, as gradeRules reports it. */
const passesRules = (rules: Rules | undefined, formula: string) => gradeRules(rules, [formula]).every((i) => i.status === 'pass');

const MANY_SEEDS = [...SEEDS, ...Array.from({ length: 150 }, (_, k) => ((k + 1) * 2654435761) >>> 0)];

/** The state a step's variant puts the sheet in, seeded as ExcelHost.check seeds it. */
function variantState(step: number, label: RegExp, seed: number): OpsForecastData {
  const base = opsForecastReview.make(new Rng(seed));
  const i = opsForecastReview.steps[step].variants.findIndex((v) => label.test(v.label));
  expect(i, `step ${step + 1} has a variant matching ${label}`).toBeGreaterThanOrEqual(0);
  return opsForecastReview.steps[step].variants[i].apply(base, new Rng(seed + 7919 * (i + 1)));
}

describe('mission-ops-forecast', () => {
  it('is a four-step finance mission whose skills are modeling exercises', () => {
    const ids = new Set(MODELING.map((e) => e.id));
    expect(opsForecastReview.role).toBe('finance');
    expect(opsForecastReview.steps.length).toBe(4);
    for (const s of opsForecastReview.skills) expect(ids.has(s), s).toBe(true);
    expect(opsForecastReview.brief(opsForecastReview.make(new Rng(1))).from).toMatch(/^[A-Z][a-z]+ [A-Z][a-z]+, /);
  });

  it('starts on Base and asks for the dropdown in step 1', () => {
    const d = opsForecastReview.make(new Rng(7));
    expect(d.scenario).toBe('Base');
    expect(d.columns).toEqual([...SCENARIOS]);
    expect(opsForecastReview.steps[0].inspections?.(d)).toEqual([
      { kind: 'validationList', cell: 'B1', options: [...SCENARIOS], label: expect.any(String) },
      { kind: 'sheet', check: { kind: 'values', range: 'B1', expected: [['Base']], describe: 'the scenario' }, label: 'B1 is left on Base', advice: expect.any(String) },
    ]);
  });

  it('checks B1 is back on Base in every step whose answer depends on it', () => {
    const d = opsForecastReview.make(new Rng(7));
    const onBase = (s: number) => (opsForecastReview.steps[s].inspections?.(d) ?? []).some((i) => i.kind === 'sheet' && i.check.kind === 'values' && i.check.range === 'B1');
    expect([0, 1, 2, 3].map(onBase)).toEqual([true, true, true, false]);
  });

  it('tells step 2 not to round, as the forecast exercise does', () => {
    expect(opsForecastReview.steps[1].task(opsForecastReview.make(new Rng(1)))).toMatch(/Don’t round\.$/);
  });

  it('step 3 checks the margin’s percentage format and allows the common IFERROR margin', () => {
    const step = opsForecastReview.steps[2];
    const percent = step.inspections?.(opsForecastReview.make(new Rng(1))).find((i): i is Extract<Inspection, { kind: 'sheet' }> => i.kind === 'sheet' && i.check.kind === 'numberFormat');
    if (!percent) throw new Error('Expected a numberFormat inspection');
    expect(percent.check).toMatchObject({ range: 'B24' });
    expect(gradeSheetCheck(percent, { kind: 'numberFormat', address: 'B24', formats: [['0%']] }).status).toBe('pass');
    const left = gradeSheetCheck(percent, { kind: 'numberFormat', address: 'B24', formats: [['#,##0']] });
    expect(left.status).toBe('fail');
    expect(left.detail).toContain('Home › Percent Style');
    expect(passesRules(step.rules, '=IFERROR(B23/B22,0)')).toBe(true);
  });

  it('step 4 accepts a check that reads both sides, rounded or not, and rejects one that can never fire', () => {
    const { rules } = opsForecastReview.steps[3];
    for (const f of [
      '=SUM(B18:B21)-SUM(C11:N11)',
      '=ROUND(SUM(B18:B21)-SUM(C11:N11),2)',
      '=SUM($B$18:$B$21)-SUM($C$11:$N$11)',
      '=B18+B19+B20+B21-SUM(C11:N11)',
      '=SUM(B18:B21) - SUM(C11:N11)',
    ]) {
      expect(passesRules(rules, f), f).toBe(true);
    }
    // Always 0: a stray SUM, and annual revenue minus the Revenue row it is built from.
    for (const f of ['=SUM(A2)', '=B22-SUM(C11:N11)']) expect(passesRules(rules, f), f).toBe(false);
    // The task ties the quarters to the forecast's Revenue row, not to another summary figure.
    expect(passesRules(rules, '=SUM(B18:B21)-B22')).toBe(false);
  });

  it('builds the forecast with one formula per row and the summary from it', () => {
    const answers = opsForecastReview.steps.map((s) => s.answer(opsForecastReview.make(new Rng(1))));
    expect(answers[1]).toMatchObject({ kind: 'cells', range: 'C10:N15', consistency: 'rows' });
    expect(answers.map((a) => (a.kind === 'cells' ? a.range : ''))).toEqual(['E4:E7', 'C10:N15', 'B18:B24', 'B26']);
  });

  it.each(SEEDS)('seed %i: switching to Upside or Downside moves every output of steps 1 to 3', (seed) => {
    const base = opsForecastReview.make(new Rng(seed));
    for (const step of [0, 1, 2]) {
      const ex = missionStepExercise(opsForecastReview, step);
      const before = ex.expected(base).flat() as number[];
      for (const scenario of [/Upside$/, /Downside$/]) {
        const after = ex.expected(variantState(step, scenario, seed)).flat() as number[];
        after.forEach((v, k) => expect(v, `step ${step + 1}, ${scenario}, cell ${k}`).not.toBe(before[k]));
      }
    }
  });

  it.each(SEEDS)('seed %i: the column swap still selects Upside, so the Live column must look it up by name', (seed) => {
    const base = opsForecastReview.make(new Rng(seed));
    const d = variantState(0, /swap/, seed);
    expect(d.columns).toEqual(['Base', 'Downside', 'Upside']);
    expect(opsForecastReview.steps[0].expected(d)).toEqual(base.values.Upside.map((v) => [v]));
  });

  it.each(MANY_SEEDS)('seed %i: the summary adds up and the margin is EBITDA over revenue in every scenario', (seed) => {
    const base = opsForecastReview.make(new Rng(seed));
    for (const scenario of SCENARIOS) {
      const d = { ...base, scenario };
      const [q1, q2, q3, q4, revenue, ebitda, margin] = opsSummary(d);
      const lines = opsForecast(d);
      expect(q1 + q2 + q3 + q4).toBeCloseTo(revenue, 6);
      expect(revenue).toBeCloseTo(lines[1].reduce((a, b) => a + b, 0), 6);
      expect(margin).toBeCloseTo(ebitda / revenue, 12);
      expect(revenue).toBeGreaterThan(0);
      // Each quarter is three months: never the year, never one month.
      for (const q of [q1, q2, q3, q4]) expect(q).toBeLessThan(revenue / 2);
    }
  });

  it.each(MANY_SEEDS)('seed %i: the Base case is profitable in month 1 and Upside beats Downside', (seed) => {
    const base = opsForecastReview.make(new Rng(seed));
    expect(opsForecast(base)[5][0]).toBeGreaterThan(0);
    const annual = (scenario: (typeof SCENARIOS)[number]) => opsSummary({ ...base, scenario })[5];
    expect(annual('Upside')).toBeGreaterThan(annual('Base'));
    expect(annual('Base')).toBeGreaterThan(annual('Downside'));
  });

  it('expects the check cell to read 0 on every run', () => {
    const step = opsForecastReview.steps[3];
    const base = opsForecastReview.make(new Rng(42));
    expect(step.expected(base)).toEqual([[0]]);
    for (const v of step.variants) expect(step.expected(v.apply(base, new Rng(1)))).toEqual([[0]]);
  });
});

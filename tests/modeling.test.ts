import { describe, expect, it } from 'vitest';
import { colToNumber, numberToCol, parseCell, parseRange } from '../src/engine/address';
import { compareGrids } from '../src/engine/compare';
import { DEPARTMENTS, sum } from '../src/engine/data';
import { isFormula } from '../src/engine/formula';
import { gradeStructure, type AnswerRead } from '../src/engine/grade';
import { gradeSheetCheck, gradeValidationList } from '../src/engine/sheetChecks';
import { Rng } from '../src/engine/rng';
import type { AnswerArea, Cell, Exercise, ExpectedGrid, Grid, InputWrite, Layout } from '../src/engine/types';
import { missionStepExercise } from '../src/missions/compile';
import { opsForecastReview, type OpsForecastData } from '../src/missions/forecast';
import {
  AUDIT_RANGE,
  LIVE_RULES,
  MODELING,
  MONTH_COLS,
  SCENARIOS,
  auditBugs,
  auditGrid,
  checkResults,
  modelAuditTrace,
  modelCheckCells,
  modelDriverForecast,
  modelScenarioSwitch,
  scenarioLeftOn,
  type AuditData,
  type CheckData,
  type ForecastData,
  type ScenarioData,
} from '../src/exercises/modeling';
import { SEEDS, blockRange, exerciseSuite, overlaps } from './helpers/suite';

exerciseSuite(MODELING);

/** Session seeds are random 32-bit numbers, so invariants are checked over many of them. */
const MANY_SEEDS = [...SEEDS, ...Array.from({ length: 150 }, (_, k) => ((k + 1) * 2654435761) >>> 0)];

/**
 * The base data and every variant, seeded both the way ExcelHost.check seeds them
 * (seed + 7919 × (index + 1)) and the way the shared suite does (seed × 31 + index).
 */
function runs<D>(ex: Exercise<D>, seed: number): { label: string; d: D }[] {
  const base = ex.make(new Rng(seed));
  return [
    { label: 'base', d: base },
    ...ex.variants.flatMap((v, i) => [
      { label: v.label, d: v.apply(base, new Rng(seed + 7919 * (i + 1))) },
      { label: v.label, d: v.apply(base, new Rng(seed * 31 + i)) },
    ]),
  ];
}

// =====================================================================
// A tiny Excel: enough of the formula language to evaluate the formulas these exercises use
// =====================================================================

type Val = number | string | boolean;
type Arg = Val | Val[][];

const TOKEN = /\s*(?:(\d+(?:\.\d+)?)|("(?:[^"]|"")*")|(\$?[A-Z]{1,3}\$?\d+)|([A-Z][A-Z0-9.]*)(?=\()|(<>|<=|>=|[-+*/(),:=<>]))/y;

const num = (v: Arg): number => {
  if (Array.isArray(v)) throw new Error('A range where a single value was expected');
  if (typeof v === 'number') return v;
  if (typeof v === 'boolean') return v ? 1 : 0;
  if (v === '') return 0;
  throw new Error(`Not a number: ${v}`);
};
const same = (a: Val, b: Val) => (typeof a === 'string' && typeof b === 'string' ? a.toLowerCase() === b.toLowerCase() : a === b);

/** A sheet of values and formulas. Formulas are evaluated on read, the way Excel recalculates. */
class Sheet {
  private cells = new Map<string, Cell>();
  private cache = new Map<string, Val>();

  set(address: string, value: Cell) {
    this.cells.set(address, value);
    this.cache.clear();
  }

  write(address: string, values: Grid) {
    const { start } = parseRange(address);
    values.forEach((row, r) => row.forEach((v, c) => this.set(`${numberToCol(start.col + c)}${start.row + r}`, v)));
  }

  raw(address: string): Cell | undefined {
    return this.cells.get(address);
  }

  value(address: string): Val {
    const hit = this.cache.get(address);
    if (hit !== undefined) return hit;
    const raw = this.cells.get(address) ?? '';
    const v = isFormula(raw) ? this.evaluate(raw.slice(1)) : raw === null ? '' : raw;
    this.cache.set(address, v);
    return v;
  }

  range(address: string): Val[][] {
    const { start, end } = parseRange(address);
    return Array.from({ length: end.row - start.row + 1 }, (_, r) =>
      Array.from({ length: end.col - start.col + 1 }, (_, c) => this.value(`${numberToCol(start.col + c)}${start.row + r}`)),
    );
  }

  private evaluate(src: string): Val {
    const text = src.trimEnd();
    const tokens: string[] = [];
    TOKEN.lastIndex = 0;
    while (TOKEN.lastIndex < text.length) {
      const m = TOKEN.exec(text);
      if (!m) throw new Error(`Can’t read the formula =${src}`);
      tokens.push(m[0].trim());
    }
    let i = 0;
    const peek = () => tokens[i];
    const take = (t?: string) => {
      if (t && tokens[i] !== t) throw new Error(`Expected ${t} in =${src}`);
      return tokens[i++];
    };
    const ref = (t: string) => t.replace(/\$/g, '');

    const primary = (): Arg => {
      const t = take();
      if (t === '(') {
        const v = compare();
        take(')');
        return v;
      }
      if (t === '-') return -num(primary());
      if (/^\d/.test(t)) return Number(t);
      if (t.startsWith('"')) return t.slice(1, -1).replace(/""/g, '"');
      if (/^\$?[A-Z]{1,3}\$?\d+$/.test(t)) {
        if (peek() === ':') {
          take(':');
          return this.range(`${ref(t)}:${ref(take())}`);
        }
        return this.value(ref(t));
      }
      take('(');
      const args: Arg[] = [];
      while (peek() !== ')') {
        args.push(compare());
        if (peek() === ',') take(',');
      }
      take(')');
      return call(t, args);
    };
    const term = (): Arg => {
      let v = primary();
      while (peek() === '*' || peek() === '/') v = take() === '*' ? num(v) * num(primary()) : num(v) / num(primary());
      return v;
    };
    const additive = (): Arg => {
      let v = term();
      while (peek() === '+' || peek() === '-') v = take() === '+' ? num(v) + num(term()) : num(v) - num(term());
      return v;
    };
    const compare = (): Arg => {
      const a = additive();
      if (!['=', '<>', '<', '>', '<=', '>='].includes(peek())) return a;
      const op = take();
      const b = additive() as Val;
      const x = a as Val;
      if (op === '=') return same(x, b);
      if (op === '<>') return !same(x, b);
      return op === '<' ? num(x) < num(b) : op === '>' ? num(x) > num(b) : op === '<=' ? num(x) <= num(b) : num(x) >= num(b);
    };
    const v = compare();
    if (i !== tokens.length) throw new Error(`Unread tokens in =${src}`);
    if (Array.isArray(v)) throw new Error(`=${src} returns a range`);
    return v;
  }
}

function call(name: string, args: Arg[]): Val {
  const flat = (a: Arg): Val[] => (Array.isArray(a) ? a.flat() : [a]);
  switch (name) {
    case 'SUM':
      return args.reduce<number>((s, a) => s + (Array.isArray(a) ? flat(a).filter((v): v is number => typeof v === 'number').reduce((x, y) => x + y, 0) : num(a)), 0);
    case 'XLOOKUP': {
      const at = flat(args[1]).findIndex((v) => same(v, args[0] as Val));
      return at < 0 ? '#N/A' : flat(args[2])[at];
    }
    case 'IF':
      return (args[0] ? args[1] : (args[2] ?? false)) as Val;
    case 'AND':
      return args.every((a) => !!a);
    default:
      throw new Error(`The test sheet doesn’t know ${name}`);
  }
}

/** Shifts the relative parts of every A1 reference, like Excel's fill handle. */
function shifted(formula: string, cols: number, rows: number): string {
  return formula.replace(/(\$?)([A-Z]{1,3})(\$?)(\d+)/g, (_, c$: string, col: string, r$: string, row: string) =>
    `${c$}${c$ ? col : numberToCol(colToNumber(col) + cols)}${r$}${r$ ? row : Number(row) + rows}`,
  );
}

/** Formulas written once in `first` and filled right across `cols` columns. */
function fillRight(first: string, formulas: string[], cols: number): [string, string][] {
  const start = parseCell(first);
  return formulas.flatMap((f, r) => Array.from({ length: cols }, (_, c) => [`${numberToCol(start.col + c)}${start.row + r}`, shifted(f, c, 0)] as [string, string]));
}

/** The sheet setup writes for a layout, with the inputs a check run writes on top. */
function sheetFor(layout: Layout, writes: InputWrite[] = []): Sheet {
  const sheet = new Sheet();
  for (const b of layout.blocks) {
    if (b.kind === 'cells') sheet.write(b.at, b.values);
    else sheet.write(b.at, [b.columns.map((c) => c.header), ...b.rows]);
  }
  for (const w of writes) if (w.kind === 'range') sheet.write(w.address, w.values);
  return sheet;
}

const answerRange = (a: AnswerArea): string => (a.kind === 'cells' || a.kind === 'bugHunt' ? a.range : '');

function expectGrid(actual: Val[][], expected: ExpectedGrid, what: string) {
  const mismatches = compareGrids(actual, expected);
  expect(mismatches, `${what}: ${JSON.stringify(mismatches[0])}`).toEqual([]);
}

// =====================================================================
// gradeStructure: one formula filled across each row
// =====================================================================

describe("consistency 'rows'", () => {
  const area: AnswerArea = { kind: 'cells', range: 'C8:E9', consistency: 'rows' };
  const read = (formulas: string[][], r1c1: string[][]): AnswerRead => ({ address: 'C8:E9', values: formulas.map((r) => r.map(() => 0)), formulas, r1c1 });
  const units = ['=B8*(1+$B$2)', '=C8*(1+$B$2)', '=D8*(1+$B$2)'];
  const unitsR1C1 = ['=RC[-1]*(1+R2C2)', '=RC[-1]*(1+R2C2)', '=RC[-1]*(1+R2C2)'];
  const revenue = ['=C8*$B$3', '=D8*$B$3', '=E8*$B$3'];
  const revenueR1C1 = ['=R[-1]C*R3C2', '=R[-1]C*R3C2', '=R[-1]C*R3C2'];
  const consistent = (s: ReturnType<typeof gradeStructure>) => s.items.find((i) => i.id === 'consistent');

  it('passes when each row holds one formula, though the rows differ from each other', () => {
    const s = gradeStructure(area, read([units, revenue], [unitsR1C1, revenueR1C1]), 2, 3);
    expect(consistent(s)).toEqual({ id: 'consistent', label: 'One formula filled across each row', status: 'pass' });
    expect(s.formulas).toEqual([...units, ...revenue]);
  });

  it('flags the cell that breaks its row and points at it', () => {
    const s = gradeStructure(area, read([units, [revenue[0], '=D8*42', revenue[2]]], [unitsR1C1, [revenueR1C1[0], '=R[-1]C*42', revenueR1C1[2]]]), 2, 3);
    expect(consistent(s)).toMatchObject({ status: 'fail', label: 'One formula filled across each row', focus: 'D9' });
    expect(consistent(s)?.detail).toContain('D9 doesn’t match the formula in the cells around it');
  });

  it('catches a first month written differently from the rest of its row', () => {
    const s = gradeStructure(area, read([['=$B$8*(1+$B$2)', units[1], units[2]], revenue], [['=R8C2*(1+R2C2)', unitsR1C1[1], unitsR1C1[2]], revenueR1C1]), 2, 3);
    expect(consistent(s)).toMatchObject({ status: 'fail', focus: 'C8' });
  });

  it('groups by row only: the same grid fails as columns or as one formula', () => {
    const r = read([units, revenue], [unitsR1C1, revenueR1C1]);
    expect(consistent(gradeStructure({ ...area, consistency: 'columns' }, r, 2, 3))).toMatchObject({ status: 'fail', label: 'One formula filled down each column' });
    expect(consistent(gradeStructure({ ...area, consistency: 'all' }, r, 2, 3))).toMatchObject({ status: 'fail', label: 'One formula filled across and down' });
  });
});

// =====================================================================
// model-scenario-switch
// =====================================================================

const XLOOKUP_LIVE = '=XLOOKUP($B$1,$B$3:$D$3,B4:D4)';
/** A switch that ties each scenario to a column position, the bug the swap variant exists to catch. */
const POSITIONAL_LIVE = '=IF($B$1="Base",B4,IF($B$1="Upside",C4,D4))';

function liveColumn(d: ScenarioData, base: ScenarioData, formula: string): Val[][] {
  const sheet = sheetFor(modelScenarioSwitch.layout(base), modelScenarioSwitch.inputs(d));
  for (let r = 0; r < 4; r++) sheet.set(`E${4 + r}`, shifted(formula, 0, r));
  return sheet.range('E4:E7');
}

describe('model-scenario-switch', () => {
  it.each(MANY_SEEDS)('seed %i: every case differs from the others on every assumption', (seed) => {
    const d = modelScenarioSwitch.make(new Rng(seed));
    for (let i = 0; i < 4; i++) expect(new Set(SCENARIOS.map((s) => d.values[s][i])).size, `row ${i}`).toBe(3);
    expect(d.values.Base[2]).toBeGreaterThan(0.5);
    expect(d.values.Upside[2]).toBeLessThan(d.values.Base[2]);
    expect(d.values.Downside[0]).toBeLessThan(d.values.Base[0]);
  });

  it.each(SEEDS)('seed %i: every variant changes the Live values', (seed) => {
    const base = modelScenarioSwitch.expected(modelScenarioSwitch.make(new Rng(seed)));
    for (const { label, d } of runs(modelScenarioSwitch, seed).slice(1)) {
      expect(modelScenarioSwitch.expected(d), label).not.toEqual(base);
    }
  });

  it.each(SEEDS)('seed %i: the hinted XLOOKUP matches the answer key on every run', (seed) => {
    const base = modelScenarioSwitch.make(new Rng(seed));
    for (const { label, d } of runs(modelScenarioSwitch, seed)) expectGrid(liveColumn(d, base, XLOOKUP_LIVE), modelScenarioSwitch.expected(d), label);
  });

  it.each(SEEDS)('seed %i: only the column swap catches a switch that picks columns by position', (seed) => {
    const base = modelScenarioSwitch.make(new Rng(seed));
    const swap = modelScenarioSwitch.variants.findIndex((v) => /swap/.test(v.label));
    modelScenarioSwitch.variants.forEach((v, i) => {
      const d = v.apply(base, new Rng(seed + 7919 * (i + 1)));
      const wrong = compareGrids(liveColumn(d, base, POSITIONAL_LIVE), modelScenarioSwitch.expected(d)).length;
      if (i === swap) expect(wrong, v.label).toBe(4);
      else expect(wrong, v.label).toBe(0);
    });
  });

  it('the swap keeps Upside selected and moves its header with its values', () => {
    const base = modelScenarioSwitch.make(new Rng(42));
    const d = modelScenarioSwitch.variants[3].apply(base, new Rng(1));
    expect(d.scenario).toBe('Upside');
    expect(d.columns).toEqual(['Base', 'Downside', 'Upside']);
    expect(modelScenarioSwitch.expected(d)).toEqual(base.values.Upside.map((v) => [v]));
  });

  it('the dropdown the hints describe passes the validation check, from the header row or typed', () => {
    const [insp] = modelScenarioSwitch.inspections?.(modelScenarioSwitch.make(new Rng(1))) ?? [];
    if (insp?.kind !== 'validationList') throw new Error('Expected a validationList inspection');
    const fromHeader = gradeValidationList(insp, { type: 'List', source: '=$B$3:$D$3', sourceValues: [['Base', 'Upside', 'Downside']], inCellDropDown: true });
    const typed = gradeValidationList(insp, { type: 'List', source: 'Base,Upside,Downside', inCellDropDown: true });
    const missing = gradeValidationList(insp, { type: 'List', source: 'Base,Upside', inCellDropDown: true });
    expect([fromHeader.status, typed.status, missing.status]).toEqual(['pass', 'pass', 'fail']);
  });

  it('declares the dropdown on B1, a check that B1 is back on Base, and the lookup rules', () => {
    const d = modelScenarioSwitch.make(new Rng(1));
    expect(modelScenarioSwitch.inspections?.(d)).toEqual([
      { kind: 'validationList', cell: 'B1', options: ['Base', 'Upside', 'Downside'], label: expect.any(String) },
      { kind: 'sheet', check: { kind: 'values', range: 'B1', expected: [['Base']], describe: 'the scenario' }, label: 'B1 is left on Base', advice: expect.any(String) },
    ]);
    expect(modelScenarioSwitch.rules?.require?.[0].pattern.test(XLOOKUP_LIVE)).toBe(true);
    expect(modelScenarioSwitch.rules?.forbidText?.values).toEqual([...SCENARIOS]);
  });

  it('names B1 as the cause when the learner tried the dropdown and left it on another case', () => {
    const insp = scenarioLeftOn(modelScenarioSwitch.make(new Rng(1)));
    if (insp.kind !== 'sheet') throw new Error('Expected a sheet inspection');
    expect(gradeSheetCheck(insp, { kind: 'values', address: 'B1', values: [['Base']] }).status).toBe('pass');
    const left = gradeSheetCheck(insp, { kind: 'values', address: 'B1', values: [['Upside']] });
    expect(left).toMatchObject({ status: 'fail', focus: 'B1' });
    expect(left.detail).toContain('Pick Base in the B1 dropdown');
  });

  it('accepts exactly the lookups the task names: XLOOKUP, or INDEX and MATCH', () => {
    const [rule] = LIVE_RULES.require ?? [];
    for (const f of [XLOOKUP_LIVE, '=INDEX(B4:D4,MATCH($B$1,$B$3:$D$3,0))', '=INDEX(B4:D4,XMATCH($B$1,$B$3:$D$3))']) expect(rule.pattern.test(f), f).toBe(true);
    for (const f of ['=INDEX(B4:D4,1)', '=CHOOSE(MATCH($B$1,$B$3:$D$3,0),B4,C4,D4)', '=HLOOKUP($B$1,$B$3:$D$7,ROWS($A$3:A4),FALSE)']) expect(rule.pattern.test(f), f).toBe(false);
    expect(rule.label).toMatch(/XLOOKUP, or INDEX and MATCH/);
    const d = modelScenarioSwitch.make(new Rng(1));
    expect(modelScenarioSwitch.task(d)).toContain('one XLOOKUP (or INDEX and MATCH) in `E4`');
    expect(opsForecastReview.steps[0].task(opsForecastReview.make(new Rng(1)))).toContain('one XLOOKUP (or INDEX and MATCH) in `E4`');
  });
});

// =====================================================================
// model-driver-forecast
// =====================================================================

/** The hinted formulas for column C, filled right to N. */
const FORECAST_FORMULAS = ['=B8*(1+$B$2)', '=C8*$B$3', '=C9*$B$4', '=C9-C10', '=$B$5', '=C11-C12'];

function forecastSheet(d: ForecastData, base: ForecastData): Sheet {
  const sheet = sheetFor(modelDriverForecast.layout(base), modelDriverForecast.inputs(d));
  for (const [address, f] of fillRight('C8', FORECAST_FORMULAS, 12)) sheet.set(address, f);
  return sheet;
}

describe('model-driver-forecast', () => {
  it.each(SEEDS)('seed %i: the hinted formulas, filled right, match the answer key on every run', (seed) => {
    const base = modelDriverForecast.make(new Rng(seed));
    for (const { label, d } of runs(modelDriverForecast, seed)) expectGrid(forecastSheet(d, base).range('C8:N13'), modelDriverForecast.expected(d), label);
  });

  it.each(SEEDS)('seed %i: every variant changes the forecast', (seed) => {
    const base = modelDriverForecast.expected(modelDriverForecast.make(new Rng(seed)));
    for (const { label, d } of runs(modelDriverForecast, seed).slice(1)) expect(modelDriverForecast.expected(d), label).not.toEqual(base);
  });

  it('month 1 reads the Last actual column with the same formula as every later month', () => {
    const fills = fillRight('C8', FORECAST_FORMULAS, 12);
    expect(fills.find(([a]) => a === 'C8')?.[1]).toBe('=B8*(1+$B$2)');
    expect(fills.find(([a]) => a === 'N8')?.[1]).toBe('=M8*(1+$B$2)');
    const d = modelDriverForecast.make(new Rng(7));
    const units = modelDriverForecast.expected(d)[0] as number[];
    expect(units[0]).toBe(d.actual.units * (1 + d.drivers.growth));
    expect(units[11]).toBe(units[10] * (1 + d.drivers.growth));
  });

  it.each(MANY_SEEDS)('seed %i: realistic drivers, a profitable first month and actuals that differ from the forecast price', (seed) => {
    const d = modelDriverForecast.make(new Rng(seed));
    const [, revenue, , , , ebitda] = modelDriverForecast.expected(d) as number[][];
    expect(d.drivers.growth).toBeGreaterThan(0);
    expect(d.drivers.cogsPct).toBeGreaterThan(0.5);
    expect(d.drivers.cogsPct).toBeLessThan(0.7);
    expect(ebitda[0]).toBeGreaterThan(0);
    // Growing last month's revenue instead of Units × price gives a different answer.
    expect(Math.abs(d.actual.revenue * (1 + d.drivers.growth) - revenue[0])).toBeGreaterThan(100);
  });
});

// =====================================================================
// model-check-cells
// =====================================================================

/**
 * G13:G15 on one run: the summary's totals and the ledger as Excel shows them, the hinted checks,
 * and `master` in G15. The test sheet can't read Tables, so the Table's row count stands in for
 * ROWS(Expenses).
 */
function checkColumn(d: CheckData, master: string): Val[][] {
  const listed = d.lines.filter((l) => d.departments.includes(l.dept));
  const sheet = new Sheet();
  sheet.write('G8', [[listed.length, sum(listed.map((l) => l.amount))]]);
  sheet.write('H10', [[d.ledger]]);
  sheet.write('G13', [['=H8-H10'], [`=${d.lines.length}-G8`], [master]]);
  return sheet.range('G13:G15');
}

describe('model-check-cells', () => {
  const variant = (name: RegExp) => modelCheckCells.variants.findIndex((v) => name.test(v.label));

  it.each(MANY_SEEDS)('seed %i: the starting data ties, with every summary department in use', (seed) => {
    const d = modelCheckCells.make(new Rng(seed));
    expect(checkResults(d)).toEqual([0, 0, 'OK']);
    for (const dept of d.departments) expect(d.lines.filter((l) => l.dept === dept).length, dept).toBeGreaterThanOrEqual(2);
    for (const l of d.lines) expect(Number.isInteger(l.amount), 'whole dollars keep the sums exact').toBe(true);
  });

  it.each(SEEDS)('seed %i: data that changes but still ties keeps every check at 0 and OK', (seed) => {
    const base = modelCheckCells.make(new Rng(seed));
    for (const i of [variant(/amounts change/), variant(/lines are added/)]) {
      const d = modelCheckCells.variants[i].apply(base, new Rng(seed + 7919 * (i + 1)));
      expect(JSON.stringify(d.lines)).not.toBe(JSON.stringify(base.lines));
      expect(modelCheckCells.expected(d)).toEqual([[0], [0], ['OK']]);
    }
  });

  it.each(MANY_SEEDS)('seed %i: a ledger that stops matching flips the master check', (seed) => {
    const base = modelCheckCells.make(new Rng(seed));
    const i = variant(/ledger total stops matching/);
    const d = modelCheckCells.variants[i].apply(base, new Rng(seed + 7919 * (i + 1)));
    expect(d.ledger).not.toBe(base.ledger);
    expect(modelCheckCells.expected(d)).toEqual([[base.ledger - d.ledger], [0], ['Check']]);
  });

  it.each(MANY_SEEDS)('seed %i: a line coded outside the summary flips both checks and the master check', (seed) => {
    const base = modelCheckCells.make(new Rng(seed));
    const i = variant(/department the summary doesn’t list/);
    const d: CheckData = modelCheckCells.variants[i].apply(base, new Rng(seed + 7919 * (i + 1)));
    const stray = d.lines.filter((l) => !d.departments.includes(l.dept));
    expect(stray.length).toBe(1);
    expect(DEPARTMENTS).toContain(stray[0].dept);
    expect(d.ledger).toBe(base.ledger);
    expect(modelCheckCells.expected(d)).toEqual([[-stray[0].amount], [1], ['Check']]);
  });

  it.each(MANY_SEEDS)('seed %i: a line that drops out with its amount ties G13 but flips the row count and the master check', (seed) => {
    const base = modelCheckCells.make(new Rng(seed));
    const i = variant(/drops out of the summary/);
    const d: CheckData = modelCheckCells.variants[i].apply(base, new Rng(seed + 7919 * (i + 1)));
    const stray = d.lines.filter((l) => !d.departments.includes(l.dept));
    expect(stray.length).toBe(1);
    expect(d.ledger).toBe(base.ledger - stray[0].amount);
    expect(modelCheckCells.expected(d)).toEqual([[0], [1], ['Check']]);
  });

  it.each(SEEDS)('seed %i: the hinted checks are right on every run, and a master check that reads only one side is caught', (seed) => {
    const all = runs(modelCheckCells, seed);
    const wrongOn = (master: string) => all.filter(({ d }) => compareGrids(checkColumn(d, master), modelCheckCells.expected(d)).length).map((r) => r.label);
    expect(wrongOn('=IF(AND(G13=0,G14=0),"OK","Check")')).toEqual([]);
    // The two shortcuts a learner reaches for both ignore the row count.
    for (const master of ['=IF(G13=0,"OK","Check")', '=IF(H8=H10,"OK","Check")']) {
      expect(wrongOn(master), master).toContainEqual(expect.stringMatching(/drops out of the summary/));
    }
    expect(wrongOn('=IF(G14=0,"OK","Check")')).toContainEqual(expect.stringMatching(/ledger total stops matching/));
  });

  it('prewrites the summary formulas against the Table and leaves the check cells to the learner', () => {
    const d = modelCheckCells.make(new Rng(3));
    const layout = modelCheckCells.layout(d);
    const sheet = sheetFor(layout);
    expect(sheet.raw('G2')).toBe('=COUNTIFS(Expenses[Department],F2)');
    expect(sheet.raw('H7')).toBe('=SUMIFS(Expenses[Amount],Expenses[Department],F7)');
    expect(sheet.raw('H8')).toBe('=SUM(H2:H7)');
    expect(sheet.raw('H10')).toBe(d.ledger);
    expect(layout.answer).toEqual({ kind: 'cells', range: 'G13:G15', format: '#,##0', consistency: 'none' });
  });
});

// =====================================================================
// model-audit-trace
// =====================================================================

const rangeCells = (address: string) => {
  const { start, end } = parseRange(address);
  const out: string[] = [];
  for (let r = start.row; r <= end.row; r++) for (let c = start.col; c <= end.col; c++) out.push(`${numberToCol(c)}${r}`);
  return out;
};

/** The model's range under one run: as handed over, or with both cells repaired. */
function auditRun(d: AuditData, base: AuditData, repaired: boolean): Val[][] {
  const sheet = sheetFor(modelAuditTrace.layout(base), modelAuditTrace.inputs(d));
  if (repaired) {
    const p = MONTH_COLS[base.pasted];
    const w = MONTH_COLS[base.wrongLink];
    sheet.set(`${p}11`, `=${p}9-${p}10`);
    sheet.set(`${w}10`, `=${w}9*$B$4`);
  }
  return sheet.range(AUDIT_RANGE);
}

describe('model-audit-trace', () => {
  it.each(MANY_SEEDS)('seed %i: the model is prewritten to cover the range exactly, all formulas but the pasted cell', (seed) => {
    const d = modelAuditTrace.make(new Rng(seed));
    const layout = modelAuditTrace.layout(d);
    const range = parseRange(AUDIT_RANGE);
    const inside = layout.blocks.filter((b) => overlaps(blockRange(b), range));
    for (const b of inside) {
      const r = blockRange(b);
      expect(r.start.row >= range.start.row && r.end.row <= range.end.row && r.start.col >= range.start.col && r.end.col <= range.end.col, `block at ${b.at} stays inside`).toBe(true);
    }
    const covered = inside.flatMap((b) => rangeCells(`${b.at}:${numberToCol(blockRange(b).end.col)}${blockRange(b).end.row}`));
    expect(covered.sort()).toEqual(rangeCells(AUDIT_RANGE).sort());

    const sheet = sheetFor(layout);
    const pastedCell = `${MONTH_COLS[d.pasted]}11`;
    for (const a of rangeCells(AUDIT_RANGE)) expect(isFormula(sheet.raw(a)), a).toBe(a !== pastedCell);
    expect(sheet.raw(pastedCell)).toBe(d.pastedValue);
    expect(sheet.raw(`${MONTH_COLS[d.wrongLink]}10`)).toBe(`=${MONTH_COLS[d.wrongLink]}9*$C$4`);
  });

  it.each(MANY_SEEDS)('seed %i: the two bugs throw off separate cells inside the range', (seed) => {
    const d = modelAuditTrace.make(new Rng(seed));
    const [pasted, wrong] = auditBugs(d);
    expect(d.pasted).not.toBe(d.wrongLink);
    expect(pasted.cells.filter((c) => wrong.cells.includes(c))).toEqual([]);
    const all = new Set(rangeCells(AUDIT_RANGE));
    for (const c of [...pasted.cells, ...wrong.cells]) expect(all.has(c), c).toBe(true);
    expect(d.lastYear.cogsPct).not.toBe(d.forecast.cogsPct);
  });

  it.each(SEEDS)('seed %i: every bug cell is wrong on some run, every other cell is right on every run, and the repair is right everywhere', (seed) => {
    const base = modelAuditTrace.make(new Rng(seed));
    const bugCells = new Map(auditBugs(base).flatMap((b) => b.cells.map((c) => [c, b.id] as const)));
    const caught = new Set<string>();
    const { start } = parseRange(AUDIT_RANGE);
    for (const { label, d } of runs(modelAuditTrace, seed)) {
      const expected = modelAuditTrace.expected(d);
      for (const m of compareGrids(auditRun(d, base, false), expected)) {
        const address = `${numberToCol(start.col + m.col)}${start.row + m.row}`;
        expect(bugCells.has(address), `${label}: ${address} is wrong but belongs to no bug`).toBe(true);
        caught.add(address);
      }
      expectGrid(auditRun(d, base, true), expected, `${label}, repaired`);
    }
    expect([...caught].sort()).toEqual([...bugCells.keys()].sort());
  });

  it.each(SEEDS)('seed %i: the pasted number hides on the starting data and shows once the drivers move', (seed) => {
    const base = modelAuditTrace.make(new Rng(seed));
    const correct = auditGrid(base);
    expect(base.pastedValue).toBe(correct[3][base.pasted]);
    const asBuilt = auditRun(base, base, false);
    expect(asBuilt[3][base.pasted]).toBe(correct[3][base.pasted]);
    expect(asBuilt[2][base.wrongLink]).not.toBeCloseTo(correct[2][base.wrongLink] as number, 2);
    const moved = modelAuditTrace.variants[0].apply(base, new Rng(seed + 7919));
    expect(auditRun(moved, base, false)[3][base.pasted]).not.toBeCloseTo(auditGrid(moved)[3][base.pasted] as number, 2);
  });

  it('names the broken cells in the solution and the months in the bug labels', () => {
    const d = modelAuditTrace.make(new Rng(42));
    const [p, w] = [MONTH_COLS[d.pasted], MONTH_COLS[d.wrongLink]];
    expect(modelAuditTrace.solution(d)).toContain(`=${p}9-${p}10`);
    expect(modelAuditTrace.solution(d)).toContain(`=${w}9*$B$4`);
    for (const bug of auditBugs(d)) {
      expect(bug.label).toMatch(/ (was|read) /);
      expect(bug.label).toMatch(/[A-Z][a-z]{2} 20\d\d/);
    }
  });
});

// =====================================================================
// The forecast mission's hinted formulas, evaluated step by step
// =====================================================================

/** Every hinted formula for the mission's four steps, on one sheet. */
function missionSheet(d: OpsForecastData, base: OpsForecastData): Sheet {
  const sheet = sheetFor(missionStepExercise(opsForecastReview, 0).layout(base), opsForecastReview.inputs(d));
  for (let r = 0; r < 4; r++) sheet.set(`E${4 + r}`, shifted(XLOOKUP_LIVE, 0, r));
  for (const [address, f] of fillRight('C10', ['=B10*(1+$E$4)', '=C10*$H$4*(1+$E$5)', '=C11*$E$6', '=C11-C12', '=$H$5*(1+$E$7)', '=C13-C14'], 12)) sheet.set(address, f);
  const summary = ['=SUM(C11:E11)', '=SUM(F11:H11)', '=SUM(I11:K11)', '=SUM(L11:N11)', '=SUM(C11:N11)', '=SUM(C15:N15)', '=B23/B22'];
  summary.forEach((f, r) => sheet.set(`B${18 + r}`, f));
  sheet.set('B26', '=SUM(B18:B21)-SUM(C11:N11)');
  return sheet;
}

describe('mission-ops-forecast formulas', () => {
  it.each(SEEDS)('seed %i: the hinted formulas match every step’s answer key under every variant of every step', (seed) => {
    const base = opsForecastReview.make(new Rng(seed));
    opsForecastReview.steps.forEach((_, s) => {
      const ex = missionStepExercise(opsForecastReview, s);
      for (const { label, d } of runs(ex, seed)) {
        expectGrid(missionSheet(d, base).range(answerRange(ex.layout(d).answer)), ex.expected(d), `step ${s + 1}, ${label}`);
      }
    });
  });

  it('the step 3 solution text is the formulas the evaluator checks', () => {
    const text = opsForecastReview.steps[2].solution(opsForecastReview.make(new Rng(1)));
    for (const f of ['=SUM(C11:E11)', '=SUM(L11:N11)', '=SUM(C11:N11)', '=SUM(C15:N15)', '=B23/B22']) expect(text).toContain(f);
  });
});

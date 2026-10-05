import { describe, expect, it } from 'vitest';
import { parseRange } from '../src/engine/address';
import { cellInRange, parseSpot, partsFormula, probeAnswer } from '../src/engine/guide';
import type { AnswerRead } from '../src/engine/grade';
import { localize } from '../src/engine/platform';
import { Rng } from '../src/engine/rng';
import type { AnswerArea, Exercise, GuideStep, Inspection } from '../src/engine/types';
import { EXERCISES } from '../src/exercises';
import { cellList } from '../src/exercises/guides';
import { SEEDS } from './helpers/suite';

/** Words the copy guidelines rule out. */
const BANNED = /\b(please|simply|just|easy|easily|successfully|leverage|seamless)\b|!/i;
const PLACEHOLDER = /\{[A-Za-z]+(:[^}]*)?\}/;
const ERROR_NAME = /#(NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A|SPILL!|CALC!)/g;
/** Excel's limit for an address list handed to getRanges. */
const MAX_SPOT = 255;

/** A formula with the spaces outside quoted text removed, uppercased: how Excel would compare them. */
const squash = (f: string) => f.replace(/("(?:[^"]|"")*")|\s+/g, (_, quoted?: string) => quoted ?? '').toUpperCase();

function copyOf(step: GuideStep): string[] {
  return [
    step.do,
    step.why,
    ...(step.formula ?? []).map((p) => p.means),
    ...(step.show ?? []).flatMap((p) => [p.label, p.note]),
  ].filter((t): t is string => !!t);
}

/** Table names and their columns: the layout's Tables plus the ones the learner is asked to build. */
function knownTables(ex: Exercise<any>, d: unknown): Map<string, Set<string>> {
  const tables = new Map<string, Set<string>>();
  const layout = ex.layout(d);
  for (const b of layout.blocks) if (b.kind === 'data') tables.set(b.table, new Set(b.columns.map((c) => c.header)));
  const add = (table: string, column?: string) => {
    if (!tables.has(table)) tables.set(table, new Set(layout.blocks.flatMap((b) => (b.kind === 'data' && b.table === table ? b.columns.map((c) => c.header) : []))));
    if (column) tables.get(table)!.add(column);
  };
  const area = layout.answer;
  if (area.kind === 'tableColumn') add(area.table, area.column);
  for (const i of (ex.inspections?.(d) ?? []) as Inspection[]) {
    if (i.kind === 'tableExists') add(i.table);
    if (i.kind === 'tableColumn') add(i.table, i.column);
  }
  return tables;
}

function checkSpot(spot: string, tables: Map<string, Set<string>>, where: string) {
  expect(spot.length, `${where}: ${spot} is too long to select`).toBeLessThanOrEqual(MAX_SPOT);
  const p = parseSpot(spot);
  if (p.kind === 'range') {
    for (const area of p.address.split(',')) expect(() => parseRange(area), `${where}: ${area}`).not.toThrow();
    return;
  }
  // Tables made from a plain range (asTable: false) still carry their name in the layout.
  expect(tables.has(p.table), `${where}: unknown Table ${p.table}`).toBe(true);
  if (p.kind === 'column') expect(tables.get(p.table)!.has(p.column), `${where}: ${p.table} has no column ${p.column}`).toBe(true);
}

function checkAnswerStep(area: AnswerArea, cells: string | undefined, where: string) {
  expect(['cells', 'spill', 'tableColumn'], `${where}: answer steps need a cells, spill or tableColumn answer`).toContain(area.kind);
  if (cells === undefined) return;
  expect(area.kind, `${where}: only a cells answer can be judged in part`).toBe('cells');
  if (area.kind !== 'cells') return;
  const outer = parseRange(area.range);
  const inner = parseRange(cells);
  const inside = outer.start.row <= inner.start.row && inner.end.row <= outer.end.row && outer.start.col <= inner.start.col && inner.end.col <= outer.end.col;
  expect(inside, `${where}: ${cells} is outside the answer ${area.range}`).toBe(true);
}

const guided = EXERCISES.filter((e) => e.guide);

describe('walkthroughs', () => {
  it('exist for every skill', () => {
    expect(EXERCISES.filter((e) => !e.guide).map((e) => e.id)).toEqual([]);
  });

  for (const ex of guided) {
    describe(ex.id, () => {
      for (const seed of SEEDS) {
        it(`is well formed (seed ${seed})`, () => {
          const d = ex.make(new Rng(seed));
          const steps = ex.guide!(d);
          const area = ex.layout(d).answer;
          const tables = knownTables(ex, d);

          expect(steps.length).toBeGreaterThanOrEqual(2);
          expect(steps.at(-1)!.done?.kind, 'the last step is the check').toBe('check');
          expect(steps.slice(0, -1).some((s) => s.done?.kind === 'check'), 'only the last step is the check').toBe(false);

          steps.forEach((step, i) => {
            const where = `step ${i + 1}`;
            for (const text of copyOf(step)) {
              for (const platform of ['mac', 'windows', 'web'] as const) expect(localize(text, platform), `${where}: placeholder left in “${text}”`).not.toMatch(PLACEHOLDER);
              // Excel's error names (#SPILL!, #CALC!) are the one place a "!" belongs.
              expect(text.replace(ERROR_NAME, ''), `${where}: copy guidelines`).not.toMatch(BANNED);
              expect((text.match(/`/g) ?? []).length % 2, `${where}: unbalanced backticks in “${text}”`).toBe(0);
            }
            for (const p of step.formula ?? []) if (p.at) checkSpot(p.at, tables, where);
            for (const p of step.show ?? []) checkSpot(p.at, tables, where);
            if (step.done?.kind === 'answer') checkAnswerStep(area, step.done.cells, where);
            if (step.done?.kind === 'select') expect(() => step.done?.kind === 'select' && step.done.range.split(',').forEach((r) => parseRange(r))).not.toThrow();
          });

          // When the answer is one formula, the walkthrough types exactly that formula.
          const solution = localize(ex.solution(d), 'windows').trim();
          const typed = steps.filter((s) => s.formula).map((s) => squash(partsFormula(s.formula!)));
          if (typed.length && solution.startsWith('=') && !solution.includes('\n')) {
            expect(typed, 'a step types the solution').toContain(squash(solution));
          }
        });
      }
    });
  }
});

// ---------- engine ----------

const read = (address: string, formulas: unknown[][], values: unknown[][]): AnswerRead => ({ address, formulas: formulas as AnswerRead['formulas'], values: values as AnswerRead['values'], r1c1: formulas as AnswerRead['r1c1'] });

describe('probeAnswer', () => {
  const area: AnswerArea = { kind: 'cells', range: 'I2:I4', consistency: 'all' };
  const expected = [[10], [20], [30]];

  it('waits on blank cells without a note', () => {
    expect(probeAnswer(area, expected, read('I2:I4', [[''], [''], ['']], [[''], [''], ['']]), 'I2')).toEqual({ done: false });
  });

  it('ticks off the first cell before the fill', () => {
    const r = read('I2:I4', [['=A'], [''], ['']], [[10], [''], ['']]);
    expect(probeAnswer(area, expected, r, 'I2').done).toBe(true);
    expect(probeAnswer(area, expected, r)).toEqual({ done: false });
  });

  it('notes a wrong value with the learner’s formula', () => {
    const p = probeAnswer(area, expected, read('I2:I4', [['=B'], [''], ['']], [[12], [''], ['']]), 'I2');
    expect(p.done).toBe(false);
    expect(p.note).toMatch(/I2 shows 12, but it should show 10/);
    expect(p).toMatchObject({ focus: 'I2', formula: '=B' });
  });

  it('explains error values', () => {
    expect(probeAnswer(area, expected, read('I2:I4', [['=C'], [''], ['']], [['#CALC!'], [''], ['']]), 'I2').note).toMatch(/^I2 shows #CALC!\. Excel couldn’t calculate/);
  });

  it('notes a typed number', () => {
    expect(probeAnswer(area, expected, read('I2:I4', [[10], [''], ['']], [[10], [''], ['']]), 'I2').note).toMatch(/typed value/);
  });

  it('passes the whole area once filled', () => {
    expect(probeAnswer(area, expected, read('I2:I4', [['=A'], ['=A'], ['=A']], [[10], [20], [30]])).done).toBe(true);
  });

  it('judges a spill from its anchor', () => {
    const spill: AnswerArea = { kind: 'spill', anchor: 'H2' };
    const grid = [['a'], ['b']];
    expect(probeAnswer(spill, grid, read('H2:I4', [[''], [''], ['']], [[''], [''], ['']]))).toEqual({ done: false });
    expect(probeAnswer(spill, grid, read('H2:I4', [['=F', ''], ['a', ''], ['', '']], [['a', ''], ['b', ''], ['', '']])).done).toBe(true);
    expect(probeAnswer(spill, grid, read('H2:I4', [['=F', ''], ['', ''], ['', '']], [['a', ''], ['b', ''], ['c', '']])).note).toMatch(/bigger than expected: H4/);
  });

  it('judges a Table column', () => {
    const col: AnswerArea = { kind: 'tableColumn', table: 'T', column: 'Total' };
    expect(probeAnswer(col, [[1], [2]], read('E2:E3', [['=x'], ['=x']], [[1], [2]])).done).toBe(true);
    expect(probeAnswer(col, [[1], [2]], read('E2:E3', [['=x'], ['=x']], [[1], [5]])).note).toMatch(/E3 shows 5/);
  });
});

describe('spots', () => {
  it('parses Table references and range lists', () => {
    expect(parseSpot('Inventory[Value]')).toEqual({ kind: 'column', table: 'Inventory', column: 'Value' });
    expect(parseSpot('Spend[[Unit cost]]')).toEqual({ kind: 'column', table: 'Spend', column: 'Unit cost' });
    expect(parseSpot('Inventory[#Data]')).toEqual({ kind: 'table', table: 'Inventory', part: 'data' });
    expect(parseSpot('F6:F8, F12')).toEqual({ kind: 'range', address: 'F6:F8,F12', areas: 2 });
  });

  it('finds a cell in a range list', () => {
    expect(cellInRange('F7', 'F6:F8,F12')).toBe(true);
    expect(cellInRange('F12', 'F6:F8,F12')).toBe(true);
    expect(cellInRange('F9', 'F6:F8,F12')).toBe(false);
  });

  it('merges runs of rows', () => {
    expect(cellList('F', [12, 6, 7, 8])).toBe('F6:F8,F12');
    expect(cellList('D', [3])).toBe('D3');
  });
});

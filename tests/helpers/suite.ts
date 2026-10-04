import { describe, expect, it } from 'vitest';
import { cellAddress, parseCell, parseRange, rangeSize, type RangeRef } from '../../src/engine/address';
import { isMatcher } from '../../src/engine/compare';
import { isFormula } from '../../src/engine/formula';
import { Rng } from '../../src/engine/rng';
import type { AnswerArea, Block, CellsBlock, Exercise, ExpectedGrid, Layout } from '../../src/engine/types';
import { missionStepExercise } from '../../src/missions/compile';
import type { Mission } from '../../src/missions/types';

export const SEEDS = [1, 7, 42, 1234, 987654];

/** Words the copy guidelines rule out. */
const BANNED = /\b(please|simply|just|easy|easily|successfully|leverage|seamless)\b|!/i;

export function blockRange(b: Block): RangeRef {
  const start = parseCell(b.at);
  const rows = b.kind === 'data' ? b.rows.length + 1 : b.values.length;
  const cols = b.kind === 'data' ? b.columns.length : Math.max(...b.values.map((r) => r.length));
  return { start, end: { row: start.row + rows - 1, col: start.col + cols - 1 } };
}

export function overlaps(a: RangeRef, b: RangeRef): boolean {
  return a.start.row <= b.end.row && b.start.row <= a.end.row && a.start.col <= b.end.col && b.start.col <= a.end.col;
}

export function answerRange(a: AnswerArea, expected: ExpectedGrid): RangeRef | undefined {
  if (a.kind === 'cells' || a.kind === 'bugHunt') return parseRange(a.range);
  if (a.kind === 'spill') {
    const start = parseCell(a.anchor);
    // The checker reads one extra row and column past the expected result.
    return { start, end: { row: start.row + expected.length, col: start.col + (expected[0]?.length ?? 1) } };
  }
  return undefined;
}

export function wellFormed(grid: ExpectedGrid): boolean {
  return grid.every((row) => row.every((c) => isMatcher(c) || typeof c === 'string' || typeof c === 'boolean' || (typeof c === 'number' && Number.isFinite(c))));
}

const usesExpected = (a: AnswerArea) => a.kind === 'cells' || a.kind === 'spill' || a.kind === 'tableColumn' || a.kind === 'query' || a.kind === 'bugHunt';

const contains = (outer: RangeRef, inner: RangeRef) =>
  outer.start.row <= inner.start.row && inner.end.row <= outer.end.row && outer.start.col <= inner.start.col && inner.end.col <= outer.end.col;

/** The cells blocks that lie inside a bug hunt's range: the report the coach writes. */
function huntBlocks(a: Extract<AnswerArea, { kind: 'bugHunt' }>, layout: Layout): CellsBlock[] {
  const range = parseRange(a.range);
  return layout.blocks.filter((b): b is CellsBlock => b.kind === 'cells' && contains(range, blockRange(b)));
}

/**
 * A bug hunt's range is written exactly once by the coach's formula blocks. A block inside it that
 * isn't a formula can only be a planted mistake (a pasted value). Bug cells sit inside the range and
 * bug ids are unique.
 */
function checkBugHunt(a: Extract<AnswerArea, { kind: 'bugHunt' }>, layout: Layout, expected: ExpectedGrid) {
  expect(rangeSize(a.range), 'bug hunt range matches the expected grid').toEqual({ rows: expected.length, cols: expected[0]?.length ?? 0 });
  const range = parseRange(a.range);
  const inRange = (address: string) => contains(range, { start: parseCell(address), end: parseCell(address) });
  const bugCells = new Set(a.bugs.flatMap((bug) => bug.cells.map((c) => cellAddress(parseCell(c)))));
  const written = new Map<string, number>();
  for (const b of huntBlocks(a, layout)) {
    const { start, end } = blockRange(b);
    for (let row = start.row; row <= end.row; row++) {
      for (let col = start.col; col <= end.col; col++) {
        const address = cellAddress({ row, col });
        written.set(address, (written.get(address) ?? 0) + 1);
        const value = b.values[row - start.row]?.[col - start.col] ?? '';
        if (b.role === 'formula') expect(isFormula(value), `${address} in a formula block holds a formula`).toBe(true);
        else expect(bugCells.has(address), `${address} in the bug hunt range is a formula or a planted mistake`).toBe(true);
      }
    }
  }
  for (let row = range.start.row; row <= range.end.row; row++) {
    for (let col = range.start.col; col <= range.end.col; col++) {
      const address = cellAddress({ row, col });
      expect(written.get(address) ?? 0, `${address} is written once by the report`).toBe(1);
    }
  }
  expect(a.bugs.length, 'planted bugs').toBeGreaterThan(0);
  expect(new Set(a.bugs.map((bug) => bug.id)).size, 'bug ids are unique').toBe(a.bugs.length);
  for (const bug of a.bugs) {
    expect(bug.cells.length, `bug ${bug.id} has cells`).toBeGreaterThan(0);
    for (const c of bug.cells) expect(inRange(c), `bug ${bug.id} cell ${c} is inside ${a.range}`).toBe(true);
  }
}

/** Shape checks shared by exercises and mission steps. */
export function checkLayout(ex: Exercise<any>, d: unknown, layout: Layout) {
  const expected = usesExpected(layout.answer) ? ex.expected(d) : [];

  for (const b of layout.blocks) {
    blockRange(b); // addresses must parse
    if (b.kind === 'data') for (const row of b.rows) expect(row.length, `row width in ${b.table}`).toBe(b.columns.length);
  }

  const a = layout.answer;
  if (a.kind === 'pivot') expect(ex.inspections?.(d).some((i) => i.kind === 'pivot')).toBe(true);
  if (a.kind === 'objects') expect((ex.inspections?.(d) ?? []).length).toBeGreaterThan(0);
  if (a.kind === 'cellChecks') {
    expect(a.checks.length).toBeGreaterThan(0);
    for (const c of a.checks) parseCell(c.cell);
  }

  if (usesExpected(a)) {
    expect(expected.length, 'expected rows').toBeGreaterThan(0);
    expect(wellFormed(expected), 'expected values are finite numbers, text, booleans or matchers').toBe(true);
    const width = expected[0].length;
    for (const row of expected) expect(row.length).toBe(width);
    if (a.kind === 'cells') expect(rangeSize(a.range)).toEqual({ rows: expected.length, cols: width });
    if (a.kind === 'query') expect(width).toBe(a.columns.length);
    if (a.kind === 'tableColumn') {
      const table = layout.blocks.find((b) => b.kind === 'data' && b.table === a.table);
      expect(table && table.kind === 'data' ? table.rows.length : -1).toBe(expected.length);
    }
  }
  if (a.kind === 'query') expect(layout.blocks.some((b) => b.kind === 'data' && b.asTable), 'Power Query needs a source Table').toBe(true);
  if (a.kind === 'bugHunt') checkBugHunt(a, layout, expected);

  const ans = answerRange(a, expected);
  // A bug hunt's range is filled by its own formula blocks; nothing else may sit in an answer.
  const own: Block[] = a.kind === 'bugHunt' ? huntBlocks(a, layout) : [];
  if (ans) for (const b of layout.blocks) if (!own.includes(b)) expect(overlaps(ans, blockRange(b)), `answer overlaps block at ${b.at}`).toBe(false);

  // Excel grows a Table over cells written directly beside or below it, so keep a gap.
  for (const t of layout.blocks) {
    if (t.kind !== 'data' || !t.asTable) continue;
    const r = blockRange(t);
    const halo = { start: r.start, end: { row: r.end.row + 1, col: r.end.col + 1 } };
    const others = [...layout.blocks.filter((b) => b !== t).map(blockRange), ...(ans && a.kind !== 'tableColumn' ? [ans] : [])];
    for (const o of others) expect(overlaps(halo, o) && !overlaps(r, o), `something sits right next to Table ${t.table}`).toBe(false);
  }

  for (const w of ex.inputs(d)) {
    if (w.kind === 'range') expect(rangeSize(w.address), `input ${w.address} size`).toEqual({ rows: w.values.length, cols: w.values[0].length });
    else for (const row of w.rows) expect(row.length).toBe(w.columns.length);
  }
}

export function checkVariants(ex: Exercise<any>, seed: number) {
  const base = ex.make(new Rng(seed));
  const baseInputs = JSON.stringify(ex.inputs(base));
  ex.variants.forEach((v, i) => {
    const d = v.apply(base, new Rng(seed * 31 + i));
    expect(JSON.stringify(ex.inputs(d)), `variant "${v.label}" changes nothing`).not.toBe(baseInputs);
    const layout = ex.layout(d);
    if (usesExpected(layout.answer)) {
      const expected = ex.expected(d);
      expect(expected.length).toBeGreaterThan(0);
      expect(wellFormed(expected)).toBe(true);
      if (layout.answer.kind === 'cells' || layout.answer.kind === 'bugHunt') expect(rangeSize(layout.answer.range).rows).toBe(expected.length);
    }
  });
}

/** Runs the standard checks on a list of exercises. */
export function exerciseSuite(exercises: Exercise<any>[]) {
  describe.each(exercises.map((e) => [e.id, e] as [string, Exercise<any>]))('%s', (_id, ex) => {
    it('has teaching content that follows the copy rules', () => {
      expect(ex.title.length).toBeGreaterThan(5);
      expect(ex.replaces.length).toBeGreaterThan(5);
      expect(ex.hints.length).toBeGreaterThanOrEqual(2);
      expect(ex.concept.summary.length).toBeGreaterThan(20);
      const d = ex.make(new Rng(1));
      for (const text of [ex.title, ex.replaces, ex.task(d), ex.concept.summary, ...ex.hints]) expect(text, `copy: ${text}`).not.toMatch(BANNED);
    });

    it.each(SEEDS)('seed %i: deterministic, consistent layout and answer key', (seed) => {
      const d1 = ex.make(new Rng(seed));
      expect(JSON.stringify(d1)).toBe(JSON.stringify(ex.make(new Rng(seed))));
      expect(ex.task(d1).length).toBeGreaterThan(20);
      expect(ex.solution(d1).length).toBeGreaterThan(3);
      checkLayout(ex, d1, ex.layout(d1));
    });

    it.each(SEEDS)('seed %i: every variant changes the inputs and keeps a valid answer key', (seed) => checkVariants(ex, seed));
  });
}

/** Runs the standard checks on missions: every step as an exercise, plus mission-level rules. */
export function missionSuite(missions: Mission<any>[]) {
  describe.each(missions.map((m) => [m.id, m] as [string, Mission<any>]))('%s', (_id, m) => {
    it('has a brief, 3+ steps and follows the copy rules', () => {
      const d = m.make(new Rng(1));
      const brief = m.brief(d);
      expect(brief.body.length).toBeGreaterThan(60);
      expect(m.steps.length).toBeGreaterThanOrEqual(3);
      expect(m.skills.length).toBeGreaterThan(0);
      for (const text of [m.title, m.summary, brief.subject, ...m.steps.flatMap((s) => [s.title, s.task(d), ...s.hints])]) expect(text, `copy: ${text}`).not.toMatch(BANNED);
    });

    it.each(SEEDS)('seed %i: steps have valid, non-overlapping answers', (seed) => {
      const d = m.make(new Rng(seed));
      expect(JSON.stringify(d)).toBe(JSON.stringify(m.make(new Rng(seed))));
      const ranges: RangeRef[] = [];
      m.steps.forEach((_, i) => {
        const ex = missionStepExercise(m, i);
        const layout = ex.layout(d);
        checkLayout(ex, d, layout);
        const r = answerRange(layout.answer, usesExpected(layout.answer) ? ex.expected(d) : []);
        if (r) {
          for (const other of ranges) expect(overlaps(r, other), `step ${i + 1} answer overlaps an earlier step`).toBe(false);
          ranges.push(r);
        }
      });
    });

    it.each(SEEDS)('seed %i: step variants change inputs', (seed) => {
      m.steps.forEach((_, i) => checkVariants(missionStepExercise(m, i), seed));
    });
  });
}

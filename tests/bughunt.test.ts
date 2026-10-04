import { describe, expect, it } from 'vitest';
import { cellAddress, parseCell, parseRange, type CellRef } from '../src/engine/address';
import { gradeBugHunt, type BugHuntRun } from '../src/engine/bughunt';
import { cellMatches, isErrorValue } from '../src/engine/compare';
import { Rng } from '../src/engine/rng';
import type { AnswerArea, Block, Cell, CellMark, CheckItem, Exercise, Grid, InputWrite, Layout, PlantedBug } from '../src/engine/types';
import { BUG_HUNTS, HUNT_REPORTS, reportValues, type HuntData } from '../src/exercises/bughunt';
import { SEEDS, checkLayout, exerciseSuite } from './helpers/suite';

exerciseSuite(BUG_HUNTS);

type Hunt = Extract<AnswerArea, { kind: 'bugHunt' }>;

// =====================================================================
// gradeBugHunt
// =====================================================================

describe('gradeBugHunt', () => {
  // B2:C4. "typed-fee" throws off C3 and the total under it in C4; "sign" throws off B4.
  // B2, C2 and B3 belong to no mistake.
  const AREA: Hunt = {
    kind: 'bugHunt',
    range: 'B2:C4',
    bugs: [
      { id: 'typed-fee', label: 'The fee in C3 was typed in', cells: ['C3', 'C4'] },
      { id: 'sign', label: 'B4 subtracted the wrong way round', cells: ['$B$4'] },
    ],
  };
  const FORMULAS: Grid = [
    ['=A2*2', '=A2*3'],
    ['=A3*2', '=A3*$F$1'],
    ['=A4-B3', '=SUM(C2:C3)'],
  ];
  const NOW: Grid = [
    [1, 2],
    [3, 4],
    [5, 6],
  ];
  const LATER: Grid = [
    [10, 20],
    [30, 40],
    [50, 60],
  ];
  const run = (label: string, values: Grid, expected: Grid, formulas: Grid = FORMULAS): BugHuntRun => ({
    label,
    read: { address: AREA.range, values, formulas, r1c1: formulas },
    expected,
  });
  const edit = (grid: Grid, address: string, value: Cell): Grid => {
    const at = parseCell(address);
    return grid.map((row, r) => row.map((v, c) => (r === at.row - 2 && c === at.col - 2 ? value : v)));
  };
  const item = (items: CheckItem[], id: string) => items.find((i) => i.id === id);

  it('passes and paints every cell green when every mistake is fixed', () => {
    const g = gradeBugHunt(AREA, [run('', NOW, NOW), run('the fee changes', LATER, LATER)]);
    expect(g.items).toEqual([
      { id: 'formulas', label: 'Every report cell still has a formula', status: 'pass' },
      { id: 'bug-typed-fee', label: 'Fixed: The fee in C3 was typed in', status: 'pass' },
      { id: 'bug-sign', label: 'Fixed: B4 subtracted the wrong way round', status: 'pass' },
    ]);
    expect(g.marks).toEqual(['B2', 'C2', 'B3', 'C3', 'B4', 'C4'].map((address) => ({ address, ok: true })));
  });

  it('counts a mistake that shows on the current data without saying where it is', () => {
    const g = gradeBugHunt(AREA, [run('', edit(NOW, 'B4', -5), NOW), run('the fee changes', edit(LATER, 'B4', -50), LATER)]);
    expect(item(g.items, 'bug-typed-fee')?.status).toBe('pass');
    expect(item(g.items, 'bug-sign')).toBeUndefined();
    expect(item(g.items, 'bugs-left')).toEqual({
      id: 'bugs-left',
      label: '1 of 2 mistakes fixed',
      status: 'fail',
      detail: '1 mistake is still in the report. It shows on the current data.',
    });
    expect(g.marks).toEqual([]);
  });

  it('counts a mistake that only shows once the data changes, and suggests how to find it', () => {
    // A typed fee matches today: C3 and its total are right now and wrong once the fee changes.
    const g = gradeBugHunt(AREA, [run('', NOW, NOW), run('the fee changes', edit(edit(LATER, 'C3', 4), 'C4', 24), LATER)]);
    expect(item(g.items, 'bug-sign')?.status).toBe('pass');
    expect(item(g.items, 'bugs-left')).toEqual({
      id: 'bugs-left',
      label: '1 of 2 mistakes fixed',
      status: 'fail',
      detail: '1 mistake is still in the report. It shows only when the data changes. Change an input and watch which results don’t follow it. The coach puts the original inputs back when you check.',
    });
    expect(g.marks).toEqual([]);
    expect(g.items.some((i) => i.focus)).toBe(false);
  });

  it('flags a cell that was right and broke, naming the variant, and paints only that cell', () => {
    const g = gradeBugHunt(AREA, [run('', NOW, NOW), run('the fee changes', edit(LATER, 'B2', 11), LATER)]);
    expect(item(g.items, 'regression')).toEqual({
      id: 'regression',
      label: 'The rest of the report still matches',
      status: 'fail',
      detail: 'B2 was right and now shows 11; expected 10 when the fee changes.',
      focus: 'B2',
    });
    expect(item(g.items, 'bug-typed-fee')?.status).toBe('pass');
    expect(item(g.items, 'bug-sign')?.status).toBe('pass');
    expect(item(g.items, 'bugs-left')).toBeUndefined();
    expect(g.marks).toEqual([{ address: 'B2', ok: false }]);
  });

  it('suggests the learner left an input changed when cells are wrong only on the sheet as it stands', () => {
    // The learner raised the input B2 and C2 read and didn't change it back. Every variant writes
    // the coach's inputs first, so those cells match there.
    const g = gradeBugHunt(AREA, [run('', edit(edit(NOW, 'B2', 9), 'C2', 9), NOW), run('the fee changes', LATER, LATER)]);
    expect(item(g.items, 'regression')?.detail).toBe(
      '2 cells that were right no longer match. B2 was right and now shows 9; expected 1. If you changed an input to test the report, the coach has put it back. Check again.',
    );
    // Wrong under a variant too: a real regression, so no suggestion.
    const real = gradeBugHunt(AREA, [run('', edit(NOW, 'B2', 9), NOW), run('the fee changes', edit(LATER, 'B2', 90), LATER)]);
    expect(item(real.items, 'regression')?.detail).toBe('B2 was right and now shows 9; expected 1.');
  });

  it('explains an error a broken cell shows, and counts every broken cell', () => {
    const g = gradeBugHunt(AREA, [run('', edit(edit(NOW, 'B3', '#REF!'), 'C2', 9), NOW)]);
    expect(item(g.items, 'regression')?.detail).toBe(
      '2 cells that were right no longer match. C2 was right and now shows 9; expected 2.',
    );
    const only = gradeBugHunt(AREA, [run('', edit(NOW, 'B3', '#REF!'), NOW)]);
    expect(item(only.items, 'regression')?.detail).toBe('B3 was right and now shows #REF!; expected 3. The formula points at a cell that no longer exists.');
    expect(only.marks).toEqual([{ address: 'B3', ok: false }]);
  });

  it('names a typed-over cell outside the mistakes', () => {
    const typed = edit(FORMULAS, 'B3', 3);
    const g = gradeBugHunt(AREA, [run('', NOW, NOW, typed), run('the fee changes', edit(LATER, 'B3', 3), LATER)]);
    expect(item(g.items, 'formulas')).toEqual({
      id: 'formulas',
      label: 'Every report cell still has a formula',
      status: 'fail',
      detail: 'B3 is empty or has a typed value. It needs a formula that keeps up when the data changes.',
      focus: 'B3',
    });
    expect(item(g.items, 'regression')?.focus).toBe('B3');
  });

  it('keeps a typed-over mistake unfixed, even when its numbers match, and never names it', () => {
    // The learner typed today's right value into C3, and no variant happens to move it.
    const typed = edit(FORMULAS, 'C3', 4);
    const g = gradeBugHunt(AREA, [run('', NOW, NOW, typed), run('the fee changes', edit(LATER, 'C3', 4), edit(LATER, 'C3', 4))]);
    expect(item(g.items, 'formulas')).toEqual({
      id: 'formulas',
      label: 'Every report cell still has a formula',
      status: 'fail',
      detail: 'One cell in the report is empty or has a typed value. It needs a formula that keeps up when the data changes.',
    });
    expect(item(g.items, 'bug-typed-fee')).toBeUndefined();
    expect(item(g.items, 'bugs-left')?.label).toBe('1 of 2 mistakes fixed');
    expect(g.marks).toEqual([]);
  });

  it('counts typed-over cells inside and outside the mistakes together, naming only the outside one', () => {
    const typed = edit(edit(FORMULAS, 'C3', 4), 'B2', 1);
    const g = gradeBugHunt(AREA, [run('', NOW, NOW, typed)]);
    expect(item(g.items, 'formulas')?.detail).toBe('2 cells are empty or have typed values, including B2. Each one needs a formula that keeps up when the data changes.');
    expect(item(g.items, 'formulas')?.focus).toBe('B2');
  });

  it('describes how many mistakes are left and how many hide, in every combination', () => {
    const area: Hunt = {
      kind: 'bugHunt',
      range: 'B2:C4',
      bugs: [
        { id: 'a', label: 'A', cells: ['B2'] },
        { id: 'b', label: 'B', cells: ['C2'] },
        { id: 'c', label: 'C', cells: ['B3'] },
      ],
    };
    // Each mistake is fixed, visible (wrong now) or hidden (wrong only under the variant).
    const detail = (states: ('fixed' | 'visible' | 'hidden')[]) => {
      let now = NOW;
      let later = LATER;
      states.forEach((s, i) => {
        const address = area.bugs[i].cells[0];
        if (s === 'visible') now = edit(now, address, -1);
        if (s !== 'fixed') later = edit(later, address, -1);
      });
      return gradeBugHunt(area, [run('', now, NOW), run('the fee changes', later, LATER)]).items.find((i) => i.id === 'bugs-left')?.detail;
    };
    const tip = ' Change an input and watch which results don’t follow it. The coach puts the original inputs back when you check.';
    expect(detail(['fixed', 'fixed', 'fixed'])).toBeUndefined();
    expect(detail(['visible', 'visible', 'fixed'])).toBe('2 mistakes are still in the report. All of them show on the current data.');
    expect(detail(['hidden', 'hidden', 'fixed'])).toBe(`2 mistakes are still in the report. All of them show only when the data changes.${tip}`);
    expect(detail(['visible', 'hidden', 'visible'])).toBe(`3 mistakes are still in the report. 1 of them shows only when the data changes.${tip}`);
    expect(detail(['hidden', 'visible', 'hidden'])).toBe(`3 mistakes are still in the report. 2 of them show only when the data changes.${tip}`);
  });

  it('never points at, paints or describes the cells of a mistake that is still there', () => {
    const states = ['fixed', 'visible', 'hidden', 'typed'] as const;
    const breaks = ['none', 'now', 'later'] as const;
    for (const fee of states) {
      for (const sign of states) {
        for (const broken of breaks) {
          let formulas = FORMULAS;
          let now = NOW;
          let later = LATER;
          const plant = (state: (typeof states)[number], address: string) => {
            if (state === 'visible') now = edit(now, address, -7);
            if (state === 'visible' || state === 'hidden') later = edit(later, address, -7);
            if (state === 'typed') formulas = edit(formulas, address, 123);
          };
          plant(fee, 'C3');
          plant(sign, 'B4');
          if (broken === 'now') now = edit(now, 'C2', 99);
          if (broken === 'later') later = edit(later, 'C2', 99);
          const g = gradeBugHunt(AREA, [run('', now, NOW, formulas), run('the fee changes', later, LATER, formulas)]);

          const hidden = new Set<string>();
          if (fee !== 'fixed') ['C3', 'C4'].forEach((c) => hidden.add(c));
          if (sign !== 'fixed') hidden.add('B4');
          const where = `${fee}/${sign}/${broken}`;
          for (const m of g.marks) expect(hidden.has(m.address), `${where}: mark on ${m.address}`).toBe(false);
          for (const i of g.items) {
            if (i.focus) expect(hidden.has(i.focus), `${where}: focus on ${i.focus}`).toBe(false);
            for (const c of hidden) expect(i.detail ?? '', `${where}: ${i.id} names ${c}`).not.toContain(c);
          }
          const passed = g.items.every((i) => i.status !== 'fail');
          expect(passed, where).toBe(fee === 'fixed' && sign === 'fixed' && broken === 'none');
          if (!passed) expect(g.marks.every((m) => !m.ok), `${where}: only red marks before the end`).toBe(true);
        }
      }
    }
  });
});

// =====================================================================
// An independent Excel evaluator for the formulas setup writes
// =====================================================================

/** How an approximate-match VLOOKUP walks the list. 'floor' is the JS model; the rest prove a mistake shows however Excel searches. */
type Approx = 'floor' | 'ceil' | 'floor-exit' | 'ceil-exit' | 'scan';
const SEARCHES: Approx[] = ['floor', 'ceil', 'floor-exit', 'ceil-exit', 'scan'];

interface Area {
  area: Cell[][];
}
type Val = Cell | Area;
const isArea = (v: Val): v is Area => typeof v === 'object' && v !== null && 'area' in v;

type Node =
  | { t: 'num'; v: number }
  | { t: 'str'; v: string }
  | { t: 'bool'; v: boolean }
  | { t: 'ref'; a: string }
  | { t: 'range'; from: string; to: string }
  | { t: 'col'; table: string; column: string }
  | { t: 'table'; table: string }
  | { t: 'call'; fn: string; args: Node[] }
  | { t: 'neg'; x: Node }
  | { t: 'bin'; op: string; l: Node; r: Node };

const CELL_REF = /^\$?[A-Z]{1,3}\$?\d+/;
const bare = (a: string) => a.replace(/\$/g, '');
const parsed = new Map<string, Node>();

/** Parses the subset of Excel formulas the reports use: references, ranges, structured references, + - * /, and function calls. */
function parse(formula: string): Node {
  const hit = parsed.get(formula);
  if (hit) return hit;
  const src = formula.replace(/^=/, '');
  let i = 0;
  const fail = (what: string): never => {
    throw new Error(`Can't parse ${formula}: ${what} at ${i}`);
  };
  const eat = (ch: string) => (src[i] === ch ? i++ : fail(`expected ${ch}`));
  const expression = (): Node => {
    let l = term();
    while (src[i] === '+' || src[i] === '-') {
      const op = src[i++];
      l = { t: 'bin', op, l, r: term() };
    }
    return l;
  };
  const term = (): Node => {
    let l = unary();
    while (src[i] === '*' || src[i] === '/') {
      const op = src[i++];
      l = { t: 'bin', op, l, r: unary() };
    }
    return l;
  };
  const unary = (): Node => {
    if (src[i] !== '-') return primary();
    i++;
    return { t: 'neg', x: unary() };
  };
  const primary = (): Node => {
    const rest = src.slice(i);
    if (src[i] === '(') {
      i++;
      const e = expression();
      eat(')');
      return e;
    }
    if (src[i] === '"') {
      let text = '';
      i++;
      while (i < src.length && !(src[i] === '"' && src[i + 1] !== '"')) {
        text += src[i];
        i += src[i] === '"' ? 2 : 1;
      }
      eat('"');
      return { t: 'str', v: text };
    }
    const num = /^\d+(\.\d+)?/.exec(rest);
    if (num) {
      i += num[0].length;
      return { t: 'num', v: Number(num[0]) };
    }
    const cell = CELL_REF.exec(rest);
    if (cell && !/[A-Za-z0-9_[(]/.test(src[i + cell[0].length] ?? '')) {
      i += cell[0].length;
      if (src[i] !== ':') return { t: 'ref', a: bare(cell[0]) };
      i++;
      const end = CELL_REF.exec(src.slice(i)) ?? fail('range end');
      i += end[0].length;
      return { t: 'range', from: bare(cell[0]), to: bare(end[0]) };
    }
    const name = /^[A-Za-z_][A-Za-z0-9_.]*/.exec(rest) ?? fail('unexpected character');
    i += name[0].length;
    if (src[i] === '(') {
      i++;
      const args: Node[] = [];
      while (src[i] !== ')') {
        args.push(expression());
        if (src[i] === ',') i++;
        else if (src[i] !== ')') fail('expected , or )');
      }
      i++;
      return { t: 'call', fn: name[0].toUpperCase(), args };
    }
    if (src[i] === '[') {
      const close = src.indexOf(']', i);
      const column = src.slice(i + 1, close);
      i = close + 1;
      return { t: 'col', table: name[0], column };
    }
    if (/^(TRUE|FALSE)$/i.test(name[0])) return { t: 'bool', v: name[0].toUpperCase() === 'TRUE' };
    return { t: 'table', table: name[0] };
  };
  const node = expression();
  if (i !== src.length) fail('trailing text');
  parsed.set(formula, node);
  return node;
}

interface TableInfo {
  at: CellRef;
  headers: string[];
  rows: number;
}

/** The practice sheet as setup leaves it, with a check run's input writes on top. */
class Sheet {
  readonly literals = new Map<string, Cell>();
  readonly formulas = new Map<string, string>();
  readonly tables = new Map<string, TableInfo>();

  constructor(layout: Layout, writes: InputWrite[] = []) {
    for (const b of layout.blocks) {
      const start = parseCell(b.at);
      if (b.kind === 'data') {
        this.put(start, [b.columns.map((c) => c.header), ...b.rows]);
        this.tables.set(b.table.toLowerCase(), { at: start, headers: b.columns.map((c) => c.header), rows: b.rows.length });
      } else if (b.role === 'formula') {
        b.values.forEach((row, r) => row.forEach((v, c) => this.formulas.set(cellAddress({ row: start.row + r, col: start.col + c }), String(v))));
      } else {
        this.put(start, b.values);
      }
    }
    for (const w of writes) {
      if (w.kind === 'range') {
        this.put(parseRange(w.address).start, w.values);
        continue;
      }
      const t = this.tables.get(w.table.toLowerCase())!;
      expect(w.rows.length, `the variant keeps ${w.table} at ${t.rows} rows or more`).toBeGreaterThanOrEqual(t.rows);
      // The host resizes the Table over the rows below it, which must be empty.
      for (let r = t.rows; r < w.rows.length; r++) {
        for (let c = 0; c < t.headers.length; c++) {
          const address = cellAddress({ row: t.at.row + 1 + r, col: t.at.col + c });
          expect(this.literals.has(address) || this.formulas.has(address), `${w.table} grows over ${address}`).toBe(false);
        }
      }
      t.rows = w.rows.length;
      w.columns.forEach((name, c) => {
        const col = t.headers.indexOf(name);
        expect(col, `${w.table} has a ${name} column`).toBeGreaterThanOrEqual(0);
        w.rows.forEach((row, r) => this.literals.set(cellAddress({ row: t.at.row + 1 + r, col: t.at.col + col }), row[c]));
      });
    }
  }

  private put(start: CellRef, grid: Grid) {
    grid.forEach((row, r) => row.forEach((v, c) => this.literals.set(cellAddress({ row: start.row + r, col: start.col + c }), v)));
  }
}

function numberOf(v: Val): number | string {
  if (isArea(v)) return '#VALUE!';
  if (v === '' || v === null) return 0;
  if (typeof v === 'number') return v;
  if (typeof v === 'boolean') return v ? 1 : 0;
  return isErrorValue(v) ? v : '#VALUE!';
}

const textOf = (v: Cell) => String(v ?? '').toLowerCase();
const compareKeys = (a: Cell, b: Cell) => (textOf(a) < textOf(b) ? -1 : textOf(a) > textOf(b) ? 1 : 0);

/** Excel's criteria: an optional comparison operator, then a number or text. */
function criteriaTest(criterion: Cell): (v: Cell) => boolean {
  if (typeof criterion === 'number') return (v) => v === criterion;
  const [, op = '=', rest] = /^(>=|<=|<>|>|<|=)?(.*)$/.exec(String(criterion ?? ''))!;
  const n = rest.trim() !== '' && Number.isFinite(Number(rest)) ? Number(rest) : undefined;
  if (n !== undefined) {
    return (v) => {
      if (typeof v !== 'number') return op === '<>';
      return op === '=' ? v === n : op === '<>' ? v !== n : op === '>' ? v > n : op === '<' ? v < n : op === '>=' ? v >= n : v <= n;
    };
  }
  return (v) => {
    const same = typeof v === 'string' && v.toLowerCase() === rest.toLowerCase();
    return op === '<>' ? !same : op === '=' && same;
  };
}

function approxIndex(keys: Cell[], x: Cell, how: Approx): number {
  if (how === 'scan') {
    let i = 0;
    while (i < keys.length && compareKeys(keys[i], x) <= 0) i++;
    return i - 1;
  }
  let lo = 0;
  let hi = keys.length - 1;
  while (lo <= hi) {
    const mid = how.startsWith('floor') ? Math.floor((lo + hi) / 2) : Math.ceil((lo + hi) / 2);
    if (how.endsWith('exit') && compareKeys(keys[mid], x) === 0) return mid;
    if (compareKeys(keys[mid], x) <= 0) lo = mid + 1;
    else hi = mid - 1;
  }
  return hi;
}

/** Works the sheet's formulas out, recording which formula cells each one reads. */
class Evaluator {
  private readonly memo = new Map<string, Cell>();
  /** Table areas by structured reference. Tables hold only literals, so nothing in them is read as a formula. */
  private readonly tableAreas = new Map<string, Area>();
  readonly reads = new Map<string, Set<string>>();

  constructor(
    private readonly sheet: Sheet,
    private readonly search: Approx = 'floor',
  ) {}

  cell(address: string): Cell {
    const formula = this.sheet.formulas.get(address);
    if (formula === undefined) return this.sheet.literals.get(address) ?? '';
    if (!this.memo.has(address)) {
      const reads = new Set<string>();
      const v = this.eval(parse(formula), reads);
      if (isArea(v)) throw new Error(`${address} returned a range`);
      this.memo.set(address, v);
      this.reads.set(address, reads);
    }
    return this.memo.get(address)!;
  }

  private read(address: string, reads: Set<string>): Cell {
    if (this.sheet.formulas.has(address)) reads.add(address);
    return this.cell(address);
  }

  private areaOf(from: CellRef, rows: number, cols: number, reads: Set<string>): Area {
    return { area: Array.from({ length: rows }, (_, r) => Array.from({ length: cols }, (_, c) => this.read(cellAddress({ row: from.row + r, col: from.col + c }), reads))) };
  }

  private eval(node: Node, reads: Set<string>): Val {
    switch (node.t) {
      case 'num':
      case 'str':
      case 'bool':
        return node.v;
      case 'ref':
        return this.read(node.a, reads);
      case 'range': {
        const r = parseRange(`${node.from}:${node.to}`);
        return this.areaOf(r.start, r.end.row - r.start.row + 1, r.end.col - r.start.col + 1, reads);
      }
      case 'col':
      case 'table': {
        const key = node.t === 'col' ? `${node.table}[${node.column}]`.toLowerCase() : node.table.toLowerCase();
        if (!this.tableAreas.has(key)) {
          const t = this.sheet.tables.get(node.table.toLowerCase());
          if (!t) throw new Error(`No Table named ${node.table}`);
          const col = node.t === 'col' ? t.headers.findIndex((h) => h.toLowerCase() === node.column.toLowerCase()) : 0;
          if (col < 0) throw new Error(`${node.table} has no ${node.t === 'col' ? node.column : ''} column`);
          this.tableAreas.set(key, this.areaOf({ row: t.at.row + 1, col: t.at.col + col }, t.rows, node.t === 'col' ? 1 : t.headers.length, reads));
        }
        return this.tableAreas.get(key)!;
      }
      case 'neg': {
        const x = numberOf(this.eval(node.x, reads));
        return typeof x === 'number' ? -x : x;
      }
      case 'bin': {
        const l = numberOf(this.eval(node.l, reads));
        const r = numberOf(this.eval(node.r, reads));
        if (typeof l === 'string') return l;
        if (typeof r === 'string') return r;
        if (node.op === '+') return l + r;
        if (node.op === '-') return l - r;
        if (node.op === '*') return l * r;
        return r === 0 ? '#DIV/0!' : l / r;
      }
      case 'call':
        return this.call(node.fn, node.args.map((a) => this.eval(a, reads)));
    }
  }

  private call(fn: string, args: Val[]): Cell {
    const area = (v: Val) => (isArea(v) ? v.area : [[v]]);
    const flat = (v: Val) => area(v).flat();
    const scalar = (v: Val): Cell => (isArea(v) ? '#VALUE!' : v);
    switch (fn) {
      case 'SUM':
      case 'AVERAGE': {
        let total = 0;
        let count = 0;
        for (const v of args.flatMap(flat)) {
          if (isErrorValue(v)) return v;
          if (typeof v === 'number') {
            total += v;
            count++;
          }
        }
        return fn === 'SUM' ? total : count ? total / count : '#DIV/0!';
      }
      case 'SUMIFS':
      case 'COUNTIFS': {
        const values = fn === 'SUMIFS' ? flat(args[0]) : undefined;
        const pairs = fn === 'SUMIFS' ? args.slice(1) : args;
        const tests: { range: Cell[]; test: (v: Cell) => boolean }[] = [];
        for (let k = 0; k < pairs.length; k += 2) tests.push({ range: flat(pairs[k]), test: criteriaTest(scalar(pairs[k + 1])) });
        const length = tests[0].range.length;
        for (const t of tests) if (t.range.length !== length || (values && values.length !== length)) return '#VALUE!';
        let total = 0;
        for (let r = 0; r < length; r++) {
          if (!tests.every((t) => t.test(t.range[r]))) continue;
          if (!values) total++;
          else if (typeof values[r] === 'number') total += values[r] as number;
        }
        return total;
      }
      case 'XLOOKUP': {
        const want = scalar(args[0]);
        const keys = flat(args[1]);
        const hit = keys.findIndex((k) => compareKeys(k, want) === 0 && typeof k === typeof want);
        return hit < 0 ? '#N/A' : flat(args[2])[hit];
      }
      case 'VLOOKUP': {
        const want = scalar(args[0]);
        const rows = area(args[1]);
        const col = Number(scalar(args[2])) - 1;
        const keys = rows.map((r) => r[0]);
        const hit = args[3] === true ? approxIndex(keys, want, this.search) : keys.findIndex((k) => compareKeys(k, want) === 0);
        return hit < 0 ? '#N/A' : rows[hit][col];
      }
      default:
        throw new Error(`The test evaluator doesn't know ${fn}`);
    }
  }
}

// =====================================================================
// Checking the hunts the way the host does
// =====================================================================

/** The host's randomness for variant i (ExcelHost.checkBugHunt). */
const hostRng = (seed: number, i: number) => new Rng(seed + 7919 * (i + 1));

interface Run<D> {
  label: string;
  d: D;
  writes: InputWrite[];
}

function runsFor<D>(ex: Exercise<D>, d: D, seed: number, rng = hostRng): Run<D>[] {
  return [
    { label: '', d, writes: [] },
    ...ex.variants.map((v, i) => {
      const dv = v.apply(d, rng(seed, i));
      return { label: v.label, d: dv, writes: ex.inputs(dv) };
    }),
  ];
}

const huntOf = (layout: Layout) => layout.answer as Hunt;

/** Reads the range the way checkBugHunt does: the setup layout, each run's inputs written on top, Excel recalculating. */
function check<D>(ex: Exercise<D>, layout: Layout, runs: Run<D>[], search: Approx = 'floor', edits: Map<string, string> = new Map()): BugHuntRun[] {
  const range = huntOf(layout).range;
  return runs.map((run) => {
    const sheet = new Sheet(layout, run.writes);
    for (const [a, f] of edits) sheet.formulas.set(a, f);
    const excel = new Evaluator(sheet, search);
    const cellsIn = rangeGrid(range);
    const values = cellsIn.map((row) => row.map((a) => excel.cell(a)));
    const formulas = cellsIn.map((row) => row.map((a) => sheet.formulas.get(a) ?? sheet.literals.get(a) ?? ''));
    return { label: run.label, read: { address: range, values, formulas, r1c1: formulas }, expected: ex.expected(run.d) };
  });
}

function rangeGrid(range: string): string[][] {
  const r = parseRange(range);
  return Array.from({ length: r.end.row - r.start.row + 1 }, (_, i) =>
    Array.from({ length: r.end.col - r.start.col + 1 }, (_, j) => cellAddress({ row: r.start.row + i, col: r.start.col + j })),
  );
}

/** The formula setup writes in each cell of the range. */
function formulasOf(layout: Layout): Map<string, string> {
  return new Sheet(layout).formulas;
}

/** Formula edits that fix one planted mistake: the correct formula back in every cell it changed. */
function fixOf(buggy: Layout, correct: Layout, bug: PlantedBug): Map<string, string> {
  const now = formulasOf(buggy);
  const right = formulasOf(correct);
  return new Map(bug.cells.filter((c) => now.get(c) !== right.get(c)).map((c) => [c, right.get(c)!]));
}

const misses = (run: BugHuntRun, address: string, range: string) => {
  const start = parseRange(range).start;
  const at = parseCell(address);
  const r = at.row - start.row;
  const c = at.col - start.col;
  return !cellMatches(run.read.values[r][c], run.expected[r][c]);
};

const fixedCount = (items: CheckItem[]) => items.filter((i) => i.id.startsWith('bug-')).length;

/** The Table the report tallies: every hunt puts it at A1. */
const sourceTable = (layout: Layout) => layout.blocks.find((b): b is Extract<Block, { kind: 'data' }> => b.kind === 'data' && b.at === 'A1')!;
const greenAll = (marks: CellMark[], range: string) => marks.length === rangeGrid(range).flat().length && marks.every((m) => m.ok);

describe.each(BUG_HUNTS.map((ex) => [ex.id, ex] as [string, Exercise<HuntData>]))('%s', (_id, ex) => {
  const report = HUNT_REPORTS[ex.id];

  describe.each(SEEDS)('seed %i', (seed) => {
    const d = ex.make(new Rng(seed));
    const layout = ex.layout(d);
    const hunt = huntOf(layout);
    const correct = ex.layout({ ...d, bugs: [] });
    const runs = runsFor(ex, d, seed);

    it('plants 5 or 6 mistakes of different kinds, at least one hidden on the current data', () => {
      expect(d.bugs.length).toBeGreaterThanOrEqual(5);
      expect(d.bugs.length).toBeLessThanOrEqual(6);
      expect(new Set(d.bugs.map((b) => b.kind)).size).toBe(d.bugs.length);
      for (const search of SEARCHES) {
        const [now] = check(ex, layout, runs.slice(0, 1), search);
        expect(hunt.bugs.some((bug) => bug.cells.every((c) => !misses(now, c, hunt.range))), search).toBe(true);
      }
      expect(ex.task(d)).toContain(`${d.bugs.length} mistakes`);
    });

    it('the correct formulas show the answer key on the current data and under every variant', () => {
      for (const r of check(ex, correct, runs)) {
        const bad = rangeGrid(hunt.range).flat().filter((a) => misses(r, a, hunt.range));
        expect(bad, `wrong when ${r.label || 'nothing changes'}`).toEqual([]);
        for (const row of r.expected) for (const v of row) expect(typeof v === 'number' && Number.isFinite(v), `answer key holds a number, not ${String(v)}`).toBe(true);
      }
    });

    it('each formula setup writes shows what its JS twin says, mistakes and all', () => {
      const evaluated = check(ex, layout, runs);
      runs.forEach((run, i) => {
        const twin = reportValues(report, run.d);
        evaluated[i].read.values.forEach((row, r) =>
          row.forEach((v, c) => {
            const js = twin[r][c];
            const what = `${rangeGrid(hunt.range)[r][c]} when ${run.label || 'nothing changes'}`;
            if (typeof v === 'number' && typeof js === 'number') expect(Math.abs(v - js), what).toBeLessThanOrEqual(1e-9 * Math.max(1, Math.abs(js)));
            else expect(v, what).toBe(js);
          }),
        );
      });
    });

    it('every planted mistake shows a wrong value, on the current data or under a variant, however VLOOKUP searches', () => {
      for (const search of SEARCHES) {
        const evaluated = check(ex, layout, runs, search);
        for (const bug of hunt.bugs) {
          const caught = evaluated.some((r) => bug.cells.some((c) => misses(r, c, hunt.range)));
          expect(caught, `${bug.id} (${search} search)`).toBe(true);
        }
        const g = gradeBugHunt(hunt, evaluated);
        expect(g.items.find((i) => i.id === 'bugs-left')?.label, search).toBe(`0 of ${hunt.bugs.length} mistakes fixed`);
        expect(g.items.find((i) => i.id === 'regression'), search).toBeUndefined();
        expect(g.marks, search).toEqual([]);
      }
    });

    it('each mistake throws off exactly its cells, and no two mistakes share one', () => {
      const traced = new Map<string, Set<string>>();
      for (const l of [correct, layout]) {
        const excel = new Evaluator(new Sheet(l));
        for (const a of rangeGrid(hunt.range).flat()) excel.cell(a);
        for (const [reader, from] of excel.reads) for (const f of from) traced.set(f, (traced.get(f) ?? new Set()).add(reader));
      }
      const changed = formulasOf(layout);
      const right = formulasOf(correct);
      const edited = rangeGrid(hunt.range).flat().filter((a) => changed.get(a) !== right.get(a));
      const seen = new Set<string>();
      for (const bug of hunt.bugs) {
        const own = edited.filter((a) => bug.cells.includes(a));
        expect(own.length, `${bug.id} changes at least one formula`).toBeGreaterThan(0);
        const reach = new Set(own);
        const queue = [...own];
        while (queue.length) {
          for (const reader of traced.get(queue.pop()!) ?? []) {
            if (reach.has(reader)) continue;
            reach.add(reader);
            queue.push(reader);
          }
        }
        expect([...reach].sort(), bug.id).toEqual([...bug.cells].sort());
        for (const c of bug.cells) {
          expect(seen.has(c), `${c} belongs to two mistakes`).toBe(false);
          seen.add(c);
        }
      }
      for (const a of edited) expect(seen.has(a), `${a} was changed by no mistake`).toBe(true);
    });

    it('fixing the mistakes one at a time counts each fix, with nothing else breaking', () => {
      let edits = new Map<string, string>();
      hunt.bugs.forEach((bug, k) => {
        edits = new Map([...edits, ...fixOf(layout, correct, bug)]);
        const g = gradeBugHunt(hunt, check(ex, layout, runs, 'floor', edits));
        expect(g.items.find((i) => i.id === `bug-${bug.id}`)?.status, bug.id).toBe('pass');
        expect(fixedCount(g.items)).toBe(k + 1);
        expect(g.items.find((i) => i.id === 'regression')).toBeUndefined();
        const done = k === hunt.bugs.length - 1;
        expect(g.items.every((i) => i.status !== 'fail')).toBe(done);
        if (done) expect(greenAll(g.marks, hunt.range)).toBe(true);
        else expect(g.marks).toEqual([]);
      });
    });

    it('a variant changes only the inputs it writes, never the report or its mistakes', () => {
      const setup = new Sheet(layout);
      runs.slice(1).forEach((run) => {
        const after = ex.layout(run.d);
        expect(formulasOf(after), run.label).toEqual(setup.formulas);
        expect(huntOf(after).bugs, run.label).toEqual(hunt.bugs);
        // What the host's writes leave on the sheet is exactly what the variant's own layout says.
        expect(new Sheet(layout, run.writes).literals, run.label).toEqual(new Sheet(after).literals);
        // A Table that grows keeps its gap from everything else.
        checkLayout(ex, run.d, after);
      });
    });

    it('a variant adds rows to the source Table', () => {
      const source = sourceTable(layout);
      const grown = runs.slice(1).filter((run) => run.writes.some((w) => w.kind === 'table' && w.table === source.table && w.rows.length > source.rows.length));
      expect(grown.length).toBeGreaterThan(0);
    });

    it('today’s numbers typed into formulas never count as a fix', () => {
      // =6 for a count, or =1234.5 for a total: right on the current data, wrong once it changes.
      const right = check(ex, correct, runs.slice(0, 1))[0];
      const start = parseRange(hunt.range).start;
      for (const bug of hunt.bugs) {
        const typed = new Map(
          [...fixOf(layout, correct, bug).keys()].map((a) => {
            const at = parseCell(a);
            return [a, `=${(right.read.values[at.row - start.row][at.col - start.col] as number).toFixed(12)}`] as const;
          }),
        );
        const evaluated = check(ex, layout, runs, 'floor', typed);
        for (const c of bug.cells) expect(misses(evaluated[0], c, hunt.range), `${bug.id}: ${c} matches today`).toBe(false);
        expect(gradeBugHunt(hunt, evaluated).items.find((i) => i.id === `bug-${bug.id}`), bug.id).toBeUndefined();
      }
    });

    it('a fixed range stretched to today’s last row is still a mistake, wherever it goes', () => {
      const kind = report.kinds.find((k) => k.id === 'short-range')!;
      const end = (d as HuntData & { templateEnd: number }).templateEnd;
      const last = sourceTable(layout).rows.length + 1;
      kind.slots(d).forEach((_, slot) => {
        const one = ex.layout({ ...d, bugs: [{ kind: 'short-range', slot }] });
        const [bug] = huntOf(one).bugs;
        const planted = formulasOf(one);
        const stretched = new Map(
          [...fixOf(one, correct, bug).keys()].map((a) => [a, planted.get(a)!.replace(new RegExp(`\\$${end}(?!\\d)`, 'g'), () => `$${last}`)] as const),
        );
        for (const [a, f] of stretched) expect(f, a).toContain(`$${last}`);
        const evaluated = check(ex, one, runsFor(ex, { ...d, bugs: [{ kind: 'short-range', slot }] }, seed), 'floor', stretched);
        for (const c of bug.cells) expect(misses(evaluated[0], c, hunt.range), `${bug.label}: ${c} matches today`).toBe(false);
        expect(evaluated.some((r) => bug.cells.some((c) => misses(r, c, hunt.range))), bug.label).toBe(true);
      });
    });

    it('the answer lists every mistake with its cells and the corrected formula', () => {
      const answer = ex.solution(d);
      const right = formulasOf(correct);
      const lines = answer.split('\n');
      expect(lines.length).toBe(hunt.bugs.length);
      hunt.bugs.forEach((bug, i) => {
        expect(lines[i]).toContain(bug.label);
        const first = [...fixOf(layout, correct, bug).keys()].sort((a, b) => parseCell(a).row - parseCell(b).row || parseCell(a).col - parseCell(b).col)[0];
        expect(lines[i]).toContain(right.get(first)!);
        expect(lines[i]).toContain(first);
      });
    });
  });

  it('stays deterministic and catchable across many more seeds, with every kind of mistake in use', () => {
    const kinds = new Set<string>();
    const positions = new Set<string>();
    for (let seed = 100; seed < 160; seed++) {
      const d = ex.make(new Rng(seed));
      expect(JSON.stringify(ex.make(new Rng(seed)))).toBe(JSON.stringify(d));
      const layout = ex.layout(d);
      const hunt = huntOf(layout);
      d.bugs.forEach((b) => kinds.add(b.kind));
      hunt.bugs.forEach((b) => positions.add(`${b.id}@${b.cells[0]}`));
      // Different draws for the variants than the host's, too: no mistake depends on one. An
      // approximate lookup must show however Excel walks the list.
      const searches: Approx[] = d.bugs.some((b) => b.kind === 'approx-lookup') ? SEARCHES : ['floor'];
      for (const rng of [hostRng, (s: number, i: number) => new Rng(s * 131 + i + 1)]) {
        for (const search of searches) {
          const g = gradeBugHunt(hunt, check(ex, layout, runsFor(ex, d, seed, rng), search));
          expect(g.items.find((i) => i.id === 'bugs-left')?.label, `seed ${seed}, ${search}`).toBe(`0 of ${hunt.bugs.length} mistakes fixed`);
          expect(g.items.find((i) => i.id === 'regression'), `seed ${seed}, ${search}`).toBeUndefined();
        }
      }
    }
    expect([...kinds].sort()).toEqual(report.kinds.map((k) => k.id).sort());
    expect(positions.size, 'mistakes land in different places from rep to rep').toBeGreaterThan(30);
  }, 30_000);
});

describe('bug hunt pools', () => {
  // The third hint names the areas to check, never cells. Every place a mistake can go must be among them.
  const SUMMARY_ROWS: Record<string, Record<number, string>> = {
    'bughunt-inventory': { 7: 'Total', 8: 'Average' },
    'bughunt-commission': { 5: 'subtotals', 9: 'subtotals', 10: 'Total', 11: 'Average' },
  };
  const BUDGET_ROWS: Record<number, string> = {
    4: 'account rows (`4` and `7` to `10`)',
    5: 'row `5`',
    7: 'account rows (`4` and `7` to `10`)',
    8: 'account rows (`4` and `7` to `10`)',
    9: 'account rows (`4` and `7` to `10`)',
    10: 'account rows (`4` and `7` to `10`)',
    11: '`11`',
    12: '`12`',
    13: 'row `13`',
    14: 'row `14`',
  };

  it.each(BUG_HUNTS.map((ex) => [ex.id, ex] as [string, Exercise<HuntData>]))('%s: the third hint names every area a mistake can go', (id, ex) => {
    const d = ex.make(new Rng(1));
    const literals = new Sheet(ex.layout(d)).literals;
    const hint = ex.hints[2];
    for (const kind of HUNT_REPORTS[id].kinds) {
      for (const slot of kind.slots(d)) {
        const cells = Object.keys(slot.formulas).map(bare);
        const rows = [...new Set(cells.map((c) => parseCell(c).row))];
        const cols = [...new Set(cells.map((c) => c.replace(/\d+$/, '')))];
        let area: string;
        if (id === 'bughunt-pnl') {
          expect(rows.length, cells.join()).toBe(1);
          area = BUDGET_ROWS[rows[0]];
        } else if (rows.length === 1 && SUMMARY_ROWS[id][rows[0]]) {
          area = SUMMARY_ROWS[id][rows[0]];
        } else {
          expect(cols.length, cells.join()).toBe(1);
          area = String(literals.get(`${cols[0]}1`));
        }
        expect(area, `${kind.id} in ${cells.join()}`).toBeTruthy();
        expect(hint, `${kind.id} in ${cells.join()}`).toContain(area);
      }
    }
  });

  it('offer at least eight kinds of mistake per report', () => {
    for (const ex of BUG_HUNTS) expect(HUNT_REPORTS[ex.id].kinds.length, ex.id).toBeGreaterThanOrEqual(8);
  });

  it('use the bughunt module, unique ids and 8 to 10 minutes', () => {
    expect(new Set(BUG_HUNTS.map((e) => e.id)).size).toBe(BUG_HUNTS.length);
    for (const ex of BUG_HUNTS) {
      expect(ex.module).toBe('bughunt');
      expect(ex.minutes).toBeGreaterThanOrEqual(8);
      expect(ex.minutes).toBeLessThanOrEqual(10);
    }
  });

  it('carry the note about the AI assistant and keep the report free of yellow answer styling', () => {
    for (const ex of BUG_HUNTS) {
      const layout = ex.layout(ex.make(new Rng(3)));
      const notes = layout.blocks.filter((b) => b.kind === 'cells' && b.values.flat().some((v) => typeof v === 'string' && v.startsWith('Built by an AI assistant')));
      expect(notes.length, ex.id).toBe(1);
      expect(layout.alsoStyle, ex.id).toBeUndefined();
    }
  });

  it('write cell references in backticks and menu paths with ›, in the task and hints', () => {
    for (const ex of BUG_HUNTS) {
      const d = ex.make(new Rng(5));
      for (const text of [ex.task(d), ...ex.hints]) {
        // Quarters (Q3) aren't cell references.
        const bareRef = text.replace(/`[^`]*`/g, '').replace(/\bQ[1-4]\b/g, '').match(/\b[A-Z]{1,2}\d{1,3}(:[A-Z]{1,2}\d{1,3})?\b/);
        expect(bareRef, `${ex.id}: ${text}`).toBeNull();
        expect(text, ex.id).not.toMatch(/'/);
      }
      expect(ex.task(d)).toContain('Formulas › Trace Precedents');
      expect(ex.task(d)).toContain('Formulas › Show Formulas');
      expect(ex.task(d)).toContain('green error triangles');
    }
  });

  it('give mistake labels in sentence case, past tense and without banned words', () => {
    for (const ex of BUG_HUNTS) {
      for (const seed of [1, 2, 3]) {
        for (const bug of huntOf(ex.layout(ex.make(new Rng(seed)))).bugs) {
          expect(bug.label[0], bug.label).toMatch(/[A-Z]/);
          expect(bug.label, bug.label).not.toMatch(/\b(please|simply|just|easy|easily|successfully|leverage|seamless)\b|!|'/i);
        }
      }
    }
  });
});

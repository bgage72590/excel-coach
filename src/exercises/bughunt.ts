import { cellAddress, parseCell, parseRange } from '../engine/address';
import { cellMatches, isErrorValue } from '../engine/compare';
import { CATEGORIES, DEPARTMENTS, ITEMS, REGIONS, REPS, excelTextCompare, serial, sum, type ItemDef } from '../engine/data';
import { Rng, round } from '../engine/rng';
import type { Block, Cell, ColumnSpec, Concept, Exercise, Grid, InputWrite, Layout, PlantedBug, Variant } from '../engine/types';
import { FMT, cells, column, dataBlock, defineExercise, rangeWrite, tableWrite } from './common';

/**
 * Bug hunts. The coach writes a finished-looking report that "an AI assistant built", with five or
 * six mistakes planted in its formulas, and the learner finds and fixes them.
 *
 * Every report formula is written twice: as the text setup puts in the sheet, and as the value Excel
 * shows for it, worked out in JS (a Calc). Each mistake carries its own buggy Calcs, so the coach
 * knows exactly which cells it throws off (the cells it replaces and everything that reads them)
 * and can prove that Excel shows a wrong value somewhere, on the current data or once a variant
 * changes it. A mistake is planted only where it shows, and never where its cells would overlap
 * another's, so fixing one always counts on its own.
 */

/** Whole dollars, for budgets, quotas and sales. */
const DOLLARS = '$#,##0';

/**
 * Sits in its own column, one gap to the right of the report: setup autofits every column, and a
 * long note in a report column would stretch that column across the screen.
 */
const AI_NOTE = 'Built by an AI assistant from last month’s template. Check it before it goes out.';

/** Inclusive run of integers. */
const span = (from: number, to: number) => Array.from({ length: to - from + 1 }, (_, i) => from + i);

// =====================================================================================
// Excel's arithmetic, for the JS side of every formula
// =====================================================================================

const NA = '#N/A';
const VALUE = '#VALUE!';
const DIV0 = '#DIV/0!';

type Sign = '+' | '-' | '*' | '/';

/** A cell as an operand: blank is 0, TRUE is 1, text is #VALUE!, and errors pass through. */
function operand(v: Cell): number | string {
  if (v === '' || v === null) return 0;
  if (typeof v === 'number') return v;
  if (typeof v === 'boolean') return v ? 1 : 0;
  return isErrorValue(v) ? v : VALUE;
}

function arith(sign: Sign, a: Cell, b: Cell): Cell {
  const x = operand(a);
  if (typeof x === 'string') return x;
  const y = operand(b);
  if (typeof y === 'string') return y;
  if (sign === '+') return x + y;
  if (sign === '-') return x - y;
  if (sign === '*') return x * y;
  return y === 0 ? DIV0 : x / y;
}

/** SUM over references: numbers count, text and blanks are skipped, and the first error wins. */
function sumCells(values: Cell[]): Cell {
  let total = 0;
  for (const v of values) {
    if (typeof v === 'number') total += v;
    else if (isErrorValue(v)) return v;
  }
  return total;
}

/** AVERAGE over references: the same rules as SUM, and #DIV/0! when there are no numbers. */
function averageCells(values: Cell[]): Cell {
  let total = 0;
  let count = 0;
  for (const v of values) {
    if (typeof v === 'number') {
      total += v;
      count++;
    } else if (isErrorValue(v)) {
      return v;
    }
  }
  return count ? total / count : DIV0;
}

/** SUMIFS/COUNTIFS criteria: ">0"-style tests against numbers, otherwise a case-insensitive match. */
function meets(value: Cell, criterion: Cell): boolean {
  if (typeof criterion === 'number') return value === criterion;
  if (typeof criterion !== 'string') return false;
  const m = /^(<=|>=|<>|<|>|=)(-?\d+(?:\.\d+)?)$/.exec(criterion);
  if (!m) return typeof value === 'string' && value.toLowerCase() === criterion.toLowerCase();
  if (typeof value !== 'number') return m[1] === '<>';
  const n = Number(m[2]);
  if (m[1] === '<') return value < n;
  if (m[1] === '<=') return value <= n;
  if (m[1] === '>') return value > n;
  if (m[1] === '>=') return value >= n;
  if (m[1] === '=') return value === n;
  return value !== n;
}

/** Lookup keys match case-insensitively, as XLOOKUP and VLOOKUP do. */
function sameKey(a: Cell, b: Cell): boolean {
  return typeof a === 'string' && typeof b === 'string' ? a.toLowerCase() === b.toLowerCase() : a === b;
}

// =====================================================================================
// Formulas twice over: the text Excel gets and the value it shows
// =====================================================================================

/** Reads a cell on the practice sheet: report cells are calculated, everything else comes from the layout. */
export type Read = (address: string) => Cell;

/** A formula, or part of one: its text (no leading "=") and the value Excel shows for it. */
export interface Calc<D = unknown> {
  text: string;
  value(d: D, at: Read): Cell;
}

const bare = (address: string) => address.replace(/\$/g, '').toUpperCase();

/** A cell reference, written as given ($ and all). */
function ref(address: string): Calc {
  return { text: address, value: (_d, at) => at(bare(address)) };
}

/** A number typed into the formula. */
function typed(n: number): Calc {
  return { text: String(n), value: () => n };
}

/** `a sign b`. Both sides must be single terms (a reference, a number or a function call). */
function op<D>(a: Calc<D>, sign: Sign, b: Calc<D>): Calc<D> {
  return { text: `${a.text}${sign}${b.text}`, value: (d, at) => arith(sign, a.value(d, at), b.value(d, at)) };
}

/** Addresses in a range, row by row. */
function cellsIn(range: string): string[] {
  const r = parseRange(range);
  return span(r.start.row, r.end.row).flatMap((row) => span(r.start.col, r.end.col).map((col) => cellAddress({ row, col })));
}

function sumOf(range: string): Calc {
  return { text: `SUM(${range})`, value: (_d, at) => sumCells(cellsIn(range).map(at)) };
}

function averageOf(...ranges: string[]): Calc {
  return { text: `AVERAGE(${ranges.join(',')})`, value: (_d, at) => averageCells(ranges.flatMap(cellsIn).map(at)) };
}

/** A Table column: its name in structured references, its sheet column for A1 ranges, and its value in a row. */
interface Field<R> {
  name: string;
  letter: string;
  get(row: R): Cell;
}

/** A Table the report reads. Tables read with A1 ranges sit at A1, so their data starts on row 2. */
interface Source<D, R> {
  table: string;
  rows(d: D): readonly R[];
}

/** One SUMIFS/COUNTIFS condition: the column tested and the criteria as written (a cell like $J4, or a quoted string like ">0"). */
interface Cond<R> {
  field: Field<R>;
  criteria: string;
}

/**
 * SUMIFS, or COUNTIFS when `total` is null, over a Table. With `through`, the formula uses fixed A1
 * ranges that end at that sheet row, the way a template sized for last month’s data would.
 */
function ifs<D, R>(src: Source<D, R>, total: Field<R> | null, conds: Cond<R>[], through?: number): Calc<D> {
  const col = (f: Field<R>) => (through ? `$${f.letter}$2:$${f.letter}$${through}` : `${src.table}[${f.name}]`);
  const args = conds.flatMap((c) => [col(c.field), c.criteria]);
  return {
    text: total ? `SUMIFS(${[col(total), ...args].join(',')})` : `COUNTIFS(${args.join(',')})`,
    value: (d, at) => {
      const criteria = conds.map((c) => (c.criteria.startsWith('"') ? c.criteria.slice(1, -1) : at(bare(c.criteria))));
      const rows = through ? src.rows(d).slice(0, through - 1) : src.rows(d);
      const hits = rows.filter((r) => conds.every((c, i) => meets(c.field.get(r), criteria[i])));
      if (!total) return hits.length;
      return sum(hits.map((r) => total.get(r)).filter((v): v is number => typeof v === 'number'));
    },
  };
}

/** XLOOKUP with its default exact match: the first row whose key matches, else #N/A. */
function xlookup<D, R>(src: Source<D, R>, x: string, key: Field<R>, result: Field<R>): Calc<D> {
  return {
    text: `XLOOKUP(${x},${src.table}[${key.name}],${src.table}[${result.name}])`,
    value: (d, at) => {
      const want = at(bare(x));
      const hit = src.rows(d).find((r) => sameKey(key.get(r), want));
      return hit ? result.get(hit) : NA;
    },
  };
}

/**
 * VLOOKUP(…, TRUE), approximate match. Excel runs a binary search that assumes the first column is
 * sorted A to Z and returns the last row at or below the lookup value; on an unsorted list that can
 * be the wrong row, which is the mistake. `column` is the 1-based column `result` reads.
 */
function vlookupApprox<D, R>(src: Source<D, R>, x: string, key: Field<R>, column: number, result: Field<R>): Calc<D> {
  return {
    text: `VLOOKUP(${x},${src.table},${column},TRUE)`,
    value: (d, at) => {
      const want = String(at(bare(x)));
      const rows = src.rows(d);
      let lo = 0;
      let hi = rows.length - 1;
      while (lo <= hi) {
        const mid = (lo + hi) >> 1;
        if (excelTextCompare(String(key.get(rows[mid])), want) <= 0) lo = mid + 1;
        else hi = mid - 1;
      }
      return hi >= 0 ? result.get(rows[hi]) : NA;
    },
  };
}

// =====================================================================================
// A report with mistakes planted in it
// =====================================================================================

/** One place a mistake can go. */
interface Slot<D> {
  /** Shown once fixed: past tense, specific, with the cells. */
  label: string;
  /** The buggy formula for each cell it replaces. */
  formulas: Record<string, Calc<D>>;
  /** For the answer: the corrected formula and where it goes. */
  fix: string;
}

/** A kind of mistake and every place it could go in a rep, in a fixed order (a rep stores the index). */
interface BugKind<D> {
  id: string;
  /**
   * The mistake can match on the current data, and a variant built for it exposes it whatever that
   * variant draws. 'always': it matches by construction (a typed-in copy of today's input).
   * 'sometimes': it may (a loose lookup that can land on the right row, depending on how Excel
   * searches). Every other kind is planted only where it shows on the current data, since a random
   * variant may not expose it.
   */
  hides?: 'always' | 'sometimes';
  slots(d: D): Slot<D>[];
}

/** A planted mistake, stored as plain data so a rep regenerates from its seed. */
export interface Plant {
  kind: string;
  slot: number;
}

export interface HuntData {
  bugs: Plant[];
}

/**
 * The report the AI built: where it sits, its correct formulas and the mistakes it can carry.
 * `formulas` and the slots may only depend on what variants keep (labels, order, the as-built inputs).
 */
export interface Report<D extends HuntData> {
  range: string;
  /** Number format for each column of the range. */
  formats: string[];
  /** Everything else on the sheet: data, inputs, labels, headers and the note. */
  sheet(d: D): Block[];
  /** The correct formula for every cell in the range, row by row. */
  formulas(d: D): Calc<D>[][];
  kinds: BugKind<D>[];
}

function slotOf<D extends HuntData>(report: Report<D>, d: D, p: Plant): Slot<D> {
  const slot = report.kinds.find((k) => k.id === p.kind)?.slots(d)[p.slot];
  if (!slot) throw new Error(`No mistake ${p.kind} #${p.slot} in this report`);
  return slot;
}

function rangeCells(range: string): string[][] {
  const r = parseRange(range);
  return span(r.start.row, r.end.row).map((row) => span(r.start.col, r.end.col).map((col) => cellAddress({ row, col })));
}

/** Address to formula for the whole range, with the given mistakes written over the correct formulas. */
function formulaMap<D extends HuntData>(report: Report<D>, d: D, slots: readonly Slot<D>[]): Map<string, Calc<D>> {
  const grid = report.formulas(d);
  const map = new Map<string, Calc<D>>();
  rangeCells(report.range).forEach((row, r) => row.forEach((address, c) => map.set(address, grid[r][c])));
  for (const slot of slots) for (const [address, f] of Object.entries(slot.formulas)) map.set(bare(address), f);
  return map;
}

/** Every literal value setup writes, by address. */
function sheetValues(blocks: Block[]): Map<string, Cell> {
  const out = new Map<string, Cell>();
  for (const b of blocks) {
    const start = parseCell(b.at);
    const grid: Grid = b.kind === 'data' ? [b.columns.map((c) => c.header), ...b.rows] : b.values;
    grid.forEach((row, r) => row.forEach((v, c) => out.set(cellAddress({ row: start.row + r, col: start.col + c }), v)));
  }
  return out;
}

/**
 * Calculates `cells`, and whatever they read, the way Excel would. Cells already in `known` keep
 * their value. `reads`, when given, collects the report cells each calculated formula reads.
 */
function recalc<D>(
  d: D,
  formulas: Map<string, Calc<D>>,
  sheet: Map<string, Cell>,
  cells: Iterable<string>,
  known: Map<string, Cell> = new Map(),
  reads?: Map<string, Set<string>>,
): Map<string, Cell> {
  const todo = [...cells];
  const values = new Map(known);
  for (const a of todo) values.delete(a);
  const busy = new Set<string>();
  const value = (address: string): Cell => {
    if (values.has(address)) return values.get(address)!;
    if (busy.has(address)) throw new Error(`Circular reference at ${address}`);
    busy.add(address);
    const from = new Set<string>();
    const v = formulas.get(address)!.value(d, (a) => {
      const key = bare(a);
      if (!formulas.has(key)) return sheet.get(key) ?? '';
      from.add(key);
      return value(key);
    });
    busy.delete(address);
    values.set(address, v);
    reads?.set(address, from);
    return v;
  };
  for (const a of todo) value(a);
  return values;
}

/** The report worked out on one set of data, ready to try mistakes against. */
interface Baseline<D> {
  d: D;
  formulas: Map<string, Calc<D>>;
  sheet: Map<string, Cell>;
  values: Map<string, Cell>;
}

function baseline<D extends HuntData>(report: Report<D>, d: D, formulas: Map<string, Calc<D>>, reads?: Map<string, Set<string>>): Baseline<D> {
  const sheet = sheetValues(report.sheet(d));
  return { d, formulas, sheet, values: recalc(d, formulas, sheet, formulas.keys(), new Map(), reads) };
}

/** Report cells each of a mistake's formulas reads. */
function slotReads<D>(base: Baseline<D>, slot: Slot<D>): Map<string, Set<string>> {
  const out = new Map<string, Set<string>>();
  for (const [address, f] of Object.entries(slot.formulas)) {
    const from = new Set<string>();
    f.value(base.d, (a) => {
      const key = bare(a);
      if (!base.values.has(key)) return base.sheet.get(key) ?? '';
      from.add(key);
      return base.values.get(key)!;
    });
    out.set(bare(address), from);
  }
  return out;
}

const byPosition = (a: string, b: string) => {
  const x = parseCell(a);
  const y = parseCell(b);
  return x.row - y.row || x.col - y.col;
};

/**
 * Every cell each mistake throws off: the cells it replaces and everything that reads them,
 * following what both the correct and the buggy formulas read.
 */
function spread<D>(base: Baseline<D>, correctReads: Map<string, Set<string>>, slots: readonly Slot<D>[]): string[][] {
  const readers = new Map<string, Set<string>>();
  const link = (reader: string, from: Set<string>) => {
    for (const f of from) readers.set(f, (readers.get(f) ?? new Set()).add(reader));
  };
  for (const [reader, from] of correctReads) link(reader, from);
  for (const slot of slots) for (const [reader, from] of slotReads(base, slot)) link(reader, from);
  return slots.map((slot) => {
    const seen = new Set(Object.keys(slot.formulas).map(bare));
    const queue = [...seen];
    while (queue.length) {
      for (const reader of readers.get(queue.pop()!) ?? []) {
        if (seen.has(reader)) continue;
        seen.add(reader);
        queue.push(reader);
      }
    }
    return [...seen].sort(byPosition);
  });
}

/** What Excel shows across the report with the given mistakes planted, row by row. */
export function reportValues<D extends HuntData>(report: Report<D>, d: D, plants: readonly Plant[] = d.bugs): Cell[][] {
  const formulas = formulaMap(report, d, plants.map((p) => slotOf(report, d, p)));
  const values = recalc(d, formulas, sheetValues(report.sheet(d)), formulas.keys());
  return rangeCells(report.range).map((row) => row.map((a) => values.get(a)!));
}

/** The cells each planted mistake throws off, in plant order. */
export function bugCells<D extends HuntData>(report: Report<D>, d: D, plants: readonly Plant[] = d.bugs): string[][] {
  const reads = new Map<string, Set<string>>();
  const base = baseline(report, d, formulaMap(report, d, []), reads);
  return spread(base, reads, plants.map((p) => slotOf(report, d, p)));
}

/**
 * Randomness for the variant runs while planting. The host draws its own, so only kinds that hide
 * may count on a variant, and the variants built for them expose them whatever they draw.
 */
const PLANT_SALT = 0x5eed;

/**
 * Picks `count` mistakes of different kinds. Each must show on the current data (or, for a kind
 * that hides, under a variant), no two may share a cell, and at least one must hide on the current
 * data, so the rep teaches why the report has to stay right when the data changes.
 */
function plantBugs<D extends HuntData>(report: Report<D>, d: D, variants: Variant<D>[], rng: Rng, count: number): Plant[] {
  const formulas = formulaMap(report, d, []);
  const reads = new Map<string, Set<string>>();
  const runs = [d, ...variants.map((v, i) => v.apply(d, new Rng(PLANT_SALT + i)))].map((r, i) => baseline(report, r, formulas, i === 0 ? reads : undefined));
  const kinds = report.kinds.map((k) => ({ id: k.id, hides: k.hides, slots: k.slots(d) }));

  const memo = new Map<string, { now: boolean; ever: boolean }>();
  const shows = (kind: { id: string; hides?: string }, i: number, slot: Slot<D>) => {
    const key = `${kind.id}#${i}`;
    if (!memo.has(key)) {
      const [cells] = spread(runs[0], reads, [slot]);
      const changed = runs.map((run) => {
        const values = recalc(run.d, new Map([...run.formulas, ...Object.entries(slot.formulas).map(([a, f]) => [bare(a), f] as const)]), run.sheet, cells, run.values);
        return cells.some((c) => !cellMatches(values.get(c)!, run.values.get(c)!));
      });
      memo.set(key, { now: changed[0], ever: changed[0] || (kind.hides !== undefined && changed.some(Boolean)) });
    }
    return memo.get(key)!;
  };

  for (let attempt = 0; attempt < 60; attempt++) {
    const plants: Plant[] = [];
    const chosen: Slot<D>[] = [];
    const hiddenNow: boolean[] = [];
    for (const kind of rng.shuffle(kinds)) {
      if (plants.length === count) break;
      for (const i of rng.shuffle(span(0, kind.slots.length - 1))) {
        const slot = kind.slots[i];
        if (!shows(kind, i, slot).ever) continue;
        const footprints = spread(runs[0], reads, [...chosen, slot]).flat();
        if (new Set(footprints).size < footprints.length) continue;
        plants.push({ kind: kind.id, slot: i });
        chosen.push(slot);
        hiddenNow.push(kind.hides === 'always');
        break;
      }
    }
    if (plants.length === count && hiddenNow.some(Boolean)) {
      const first = (p: Plant) => Object.keys(slotOf(report, d, p).formulas).map(bare).sort(byPosition)[0];
      return plants.sort((a, b) => byPosition(first(a), first(b)));
    }
  }
  throw new Error(`Couldn’t plant ${count} mistakes in ${report.range}`);
}

// =====================================================================================
// Exercise wiring shared by every hunt
// =====================================================================================

interface HuntSpec<D extends HuntData> {
  id: string;
  title: string;
  replaces: string;
  minutes: number;
  task(d: D): string;
  concept: Concept;
  hints: string[];
  report: Report<D>;
  /** The data before any mistakes are planted (bugs: []). */
  data(rng: Rng): D;
  inputs(d: D): InputWrite[];
  variants: Variant<D>[];
}

function huntLayout<D extends HuntData>(report: Report<D>, d: D): Layout {
  const formulas = formulaMap(report, d, d.bugs.map((p) => slotOf(report, d, p)));
  const grid: Grid = rangeCells(report.range).map((row) => row.map((a) => `=${formulas.get(a)!.text}`));
  const footprints = bugCells(report, d);
  const bugs: PlantedBug[] = d.bugs.map((p, i) => ({ id: p.kind, label: slotOf(report, d, p).label, cells: footprints[i] }));
  const start = cellAddress(parseRange(report.range).start);
  return {
    blocks: [...report.sheet(d), cells(start, grid, 'formula', report.formats)],
    answer: { kind: 'bugHunt', range: report.range, bugs },
  };
}

function huntAnswer<D extends HuntData>(report: Report<D>, d: D): string {
  return d.bugs
    .map((p, i) => {
      const slot = slotOf(report, d, p);
      return `${i + 1}. ${slot.label}. ${slot.fix}`;
    })
    .join('\n');
}

function bugHunt<D extends HuntData>(spec: HuntSpec<D>): Exercise<D> {
  const { report, variants } = spec;
  return defineExercise<D>({
    id: spec.id,
    module: 'bughunt',
    title: spec.title,
    replaces: spec.replaces,
    minutes: spec.minutes,
    task: spec.task,
    concept: spec.concept,
    hints: spec.hints,
    solution: (d) => huntAnswer(report, d),
    make: (rng) => {
      const d = spec.data(rng);
      return { ...d, bugs: plantBugs(report, d, variants, rng, rng.int(5, 6)) };
    },
    layout: (d) => huntLayout(report, d),
    expected: (d) => reportValues(report, d, []),
    inputs: spec.inputs,
    variants,
  });
}

/** "Write =X in R4." or "Write =X in R4 and fill down to R6." */
function writeFix(formula: Calc<any>, from: string, to?: string): string {
  const fill = to && to !== from ? ` and fill down to ${to}` : '';
  return `Write =${formula.text} in ${from}${fill}.`;
}

const TOOLS =
  'Formulas › Show Formulas shows every formula. Formulas › Trace Precedents shows which cells a formula reads. Excel’s green error triangles flag formulas that look out of place, but not every triangle is a mistake.';

const HUNT_CONCEPT = {
  summary:
    'A workbook built in a hurry, by a person or an AI, tends to break in a few familiar ways: a number typed where a reference belongs, a range sized for last month, a lookup that matches loosely, a reference that slides when the formula is filled. Review the patterns rather than every number, then test the report by changing its inputs.',
  syntax: 'Formulas › Show Formulas · Formulas › Trace Precedents · Formulas › Error Checking',
  tip: 'A typed-in number matches today’s data and only goes wrong when an input changes. Change an input and watch which results don’t follow it. The coach puts the original inputs back when you check.',
};

/** The second hint in every hunt: the kinds of mistake in its pool, never where they are. */
const kindsHint = (kinds: string) => `The mistakes are the kind an AI assistant makes when it reuses a template: ${kinds}.`;

// =====================================================================================
// Reorder report (ops)
// =====================================================================================

interface StockLine {
  sku: string;
  item: string;
  category: string;
  onHand: number;
  reorder: number;
  /** Order-up-to level. Not on the sheet; drives Order qty. */
  target: number;
  qty: number;
  cost: number;
  orderCost: number;
}

interface FreightRate {
  category: string;
  rate: number;
}

export interface ReorderHuntData extends HuntData {
  stock: StockLine[];
  /** Handling fee per unit ordered, in K10. */
  fee: number;
  /** The freight rate list in sheet order, which isn't sorted. */
  freight: FreightRate[];
  /** Report row order, J2:J6. */
  categories: string[];
  /** The inputs on the day the AI built the report; a typed-in number comes from here. */
  asBuilt: { fee: number; rates: Record<string, number> };
  /** Last row of last month’s Stock export; a fixed range from the template stops here. */
  templateEnd: number;
}

type RO = ReorderHuntData;

const STOCK_COLS: ColumnSpec[] = [
  { header: 'SKU' },
  { header: 'Item' },
  { header: 'Category' },
  { header: 'On hand', format: FMT.int },
  { header: 'Reorder point', format: FMT.int },
  { header: 'Order qty', format: FMT.int },
  { header: 'Unit cost', format: FMT.currency },
  { header: 'Order cost', format: FMT.currency },
];
const FREIGHT_COLS: ColumnSpec[] = [{ header: 'Category' }, { header: 'Freight rate', format: FMT.pct }];

const STOCK = {
  category: { name: 'Category', letter: 'C', get: (s: StockLine) => s.category },
  onHand: { name: 'On hand', letter: 'D', get: (s: StockLine) => s.onHand },
  reorder: { name: 'Reorder point', letter: 'E', get: (s: StockLine) => s.reorder },
  qty: { name: 'Order qty', letter: 'F', get: (s: StockLine) => s.qty },
  orderCost: { name: 'Order cost', letter: 'H', get: (s: StockLine) => s.orderCost },
} satisfies Record<string, Field<StockLine>>;
const FREIGHT = {
  category: { name: 'Category', letter: 'J', get: (f: FreightRate) => f.category },
  rate: { name: 'Freight rate', letter: 'K', get: (f: FreightRate) => f.rate },
} satisfies Record<string, Field<FreightRate>>;

const stockTable: Source<RO, StockLine> = { table: 'Stock', rows: (d) => d.stock };
const freightTable: Source<RO, FreightRate> = { table: 'Freight', rows: (d) => d.freight };

/** Categories in J2:J6, then Total and Average per category. The fee sits in K10. */
const RO_ROWS = { first: 2, last: 6, total: 7, average: 8 } as const;
const RO_RANGE = 'K2:S8';
const RO_COLUMNS = [
  { col: 'K', header: 'SKUs', format: FMT.int },
  { col: 'L', header: 'On hand', format: FMT.int },
  { col: 'M', header: 'Units over reorder point', format: FMT.int },
  { col: 'N', header: 'SKUs to reorder', format: FMT.int },
  { col: 'O', header: 'Units to order', format: FMT.int },
  { col: 'P', header: 'Order cost', format: FMT.currency },
  { col: 'Q', header: 'Freight', format: FMT.currency },
  { col: 'R', header: 'Handling', format: FMT.currency },
  { col: 'S', header: 'Share of order cost', format: FMT.pct },
] as const;
/** Every column but Share, which divides instead of adding up. */
const RO_MEASURES = RO_COLUMNS.slice(0, 8).map((c) => c.col);
const roHeader = (col: string) => RO_COLUMNS.find((c) => c.col === col)!.header;
const roCategory = (d: RO, row: number) => d.categories[row - RO_ROWS.first];

const byCategory = (total: Field<StockLine> | null, label: string, through?: number) =>
  ifs(stockTable, total, [{ field: STOCK.category, criteria: label }], through);

/** Columns that tally one Stock field per category; off-by-one and short-range mistakes land here. */
const RO_TALLIES: Record<string, Field<StockLine> | null> = { K: null, L: STOCK.onHand, O: STOCK.qty, P: STOCK.orderCost };

function reorderFormulas(): Calc<RO>[][] {
  const rows: Calc<RO>[][] = span(RO_ROWS.first, RO_ROWS.last).map((r) => {
    const label = `$J${r}`;
    return [
      byCategory(null, label),
      byCategory(STOCK.onHand, label),
      op(byCategory(STOCK.onHand, label), '-', byCategory(STOCK.reorder, label)),
      ifs(stockTable, null, [{ field: STOCK.category, criteria: label }, { field: STOCK.qty, criteria: '">0"' }]),
      byCategory(STOCK.qty, label),
      byCategory(STOCK.orderCost, label),
      op(byCategory(STOCK.orderCost, label), '*', xlookup(freightTable, label, FREIGHT.category, FREIGHT.rate)),
      op(byCategory(STOCK.qty, label), '*', ref('$K$10')),
      op(ref(`P${r}`), '/', ref('P$7')),
    ];
  });
  rows.push([...RO_MEASURES.map((c) => sumOf(`${c}2:${c}6`)), op(ref('P7'), '/', ref('P$7'))]);
  rows.push([...RO_MEASURES.map((c) => averageOf(`${c}2:${c}6`)), op(ref('P8'), '/', ref('P$7'))]);
  return rows;
}

const RO_FORMULAS = reorderFormulas();
const roFormula = (address: string) => {
  const cell = parseCell(address);
  return RO_FORMULAS[cell.row - RO_ROWS.first][cell.col - parseCell('K1').col];
};

const reorderKinds: BugKind<RO>[] = [
  {
    id: 'typed-input',
    hides: 'always',
    slots: (d) => [
      ...span(RO_ROWS.first, RO_ROWS.last).map((r) => ({
        label:
          r === RO_ROWS.last
            ? `Handling for ${roCategory(d, r)} in R${r} had the fee typed in instead of pointing at K10`
            : `Handling in R${r}:R${RO_ROWS.last} had the fee typed in instead of pointing at K10`,
        formulas: Object.fromEntries(span(r, RO_ROWS.last).map((k) => [`R${k}`, op(byCategory(STOCK.qty, `$J${k}`), '*', typed(d.asBuilt.fee))])),
        fix: writeFix(roFormula(`R${r}`), `R${r}`, `R${RO_ROWS.last}`),
      })),
      ...span(RO_ROWS.first, RO_ROWS.last).map((r) => ({
        label: `Freight for ${roCategory(d, r)} in Q${r} had its rate typed in instead of looked up`,
        formulas: { [`Q${r}`]: op(byCategory(STOCK.orderCost, `$J${r}`), '*', typed(d.asBuilt.rates[roCategory(d, r)])) },
        fix: writeFix(roFormula(`Q${r}`), `Q${r}`),
      })),
    ],
  },
  {
    id: 'fill-drift',
    slots: () => [
      {
        label: 'Handling in R2:R6 pointed at K10 without $ signs, so the filled rows read the cells below it',
        formulas: Object.fromEntries(span(RO_ROWS.first, RO_ROWS.last).map((k) => [`R${k}`, op(byCategory(STOCK.qty, `$J${k}`), '*', ref(`K${8 + k}`))])),
        fix: writeFix(roFormula('R2'), 'R2', 'R6'),
      },
    ],
  },
  {
    id: 'approx-lookup',
    hides: 'sometimes',
    slots: () => [
      {
        label: 'Freight in Q2:Q6 looked up the rate with VLOOKUP’s approximate match (TRUE), which needs a sorted list',
        formulas: Object.fromEntries(
          span(RO_ROWS.first, RO_ROWS.last).map((k) => [`Q${k}`, op(byCategory(STOCK.orderCost, `$J${k}`), '*', vlookupApprox(freightTable, `$J${k}`, FREIGHT.category, 2, FREIGHT.rate))]),
        ),
        fix: `${writeFix(roFormula('Q2'), 'Q2', 'Q6')} VLOOKUP with FALSE as its last argument works too.`,
      },
    ],
  },
  {
    id: 'sign-flip',
    slots: (d) =>
      span(RO_ROWS.first, RO_ROWS.last).map((r) => ({
        label: `Units over reorder point for ${roCategory(d, r)} in M${r} subtracted the wrong way round`,
        formulas: { [`M${r}`]: op(byCategory(STOCK.reorder, `$J${r}`), '-', byCategory(STOCK.onHand, `$J${r}`)) },
        fix: writeFix(roFormula(`M${r}`), `M${r}`),
      })),
  },
  {
    id: 'wrong-criteria',
    slots: (d) =>
      span(RO_ROWS.first, RO_ROWS.last).map((r) => ({
        label: `SKUs to reorder for ${roCategory(d, r)} in N${r} tested On hand instead of Order qty`,
        formulas: { [`N${r}`]: ifs(stockTable, null, [{ field: STOCK.category, criteria: `$J${r}` }, { field: STOCK.onHand, criteria: '">0"' }]) },
        fix: writeFix(roFormula(`N${r}`), `N${r}`),
      })),
  },
  {
    id: 'off-by-one',
    slots: (d) =>
      Object.entries(RO_TALLIES).flatMap(([col, field]) =>
        span(RO_ROWS.first, RO_ROWS.last).flatMap((r) =>
          [-1, 1].map((step) => ({
            label: `${roHeader(col)} for ${roCategory(d, r)} in ${col}${r} read the category from the row ${step > 0 ? 'below' : 'above'}`,
            formulas: { [`${col}${r}`]: byCategory(field, `$J${r + step}`) },
            fix: writeFix(roFormula(`${col}${r}`), `${col}${r}`),
          })),
        ),
      ),
  },
  {
    id: 'short-range',
    slots: (d) =>
      Object.entries(RO_TALLIES).flatMap(([col, field]) =>
        span(RO_ROWS.first, RO_ROWS.last).map((r) => ({
          label: `${roHeader(col)} for ${roCategory(d, r)} in ${col}${r} used a fixed range that stops at row ${d.templateEnd}, so it missed the SKUs below it`,
          formulas: { [`${col}${r}`]: byCategory(field, `$J${r}`, d.templateEnd) },
          fix: writeFix(roFormula(`${col}${r}`), `${col}${r}`),
        })),
      ),
  },
  {
    id: 'short-total',
    slots: (d) =>
      RO_MEASURES.map((col) => ({
        label: `The ${roHeader(col)} total in ${col}7 left out ${roCategory(d, RO_ROWS.last)}`,
        formulas: { [`${col}7`]: sumOf(`${col}2:${col}5`) },
        fix: writeFix(roFormula(`${col}7`), `${col}7`),
      })),
  },
  {
    id: 'average-total',
    slots: () =>
      RO_MEASURES.map((col) => ({
        label: `The ${roHeader(col)} average in ${col}8 included the Total row`,
        formulas: { [`${col}8`]: averageOf(`${col}2:${col}7`) },
        fix: writeFix(roFormula(`${col}8`), `${col}8`),
      })),
  },
  {
    id: 'share-drift',
    slots: () => [
      {
        label: 'Share of order cost in S2:S8 divided by P7 without a $ sign, so the total slid down as the formula filled',
        formulas: Object.fromEntries(span(RO_ROWS.first, RO_ROWS.average).map((k) => [`S${k}`, op(ref(`P${k}`), '/', ref(`P${k + 5}`))])),
        fix: writeFix(roFormula('S2'), 'S2', 'S8'),
      },
    ],
  },
];

const reorderReport: Report<RO> = {
  range: RO_RANGE,
  formats: RO_COLUMNS.map((c) => c.format),
  sheet: (d) => [
    dataBlock('Stock', 'A1', STOCK_COLS, stockGrid(d.stock)),
    cells('J1', [['Category', ...RO_COLUMNS.map((c) => c.header)]], 'header'),
    cells('J2', column([...d.categories, 'Total', 'Average per category']), 'label'),
    cells('J10', [['Handling fee per unit']], 'label'),
    cells('K10', [[d.fee]], 'input', FMT.currency),
    dataBlock('Freight', 'J12', FREIGHT_COLS, freightGrid(d.freight)),
    cells('U1', [[AI_NOTE]], 'label'),
  ],
  formulas: () => RO_FORMULAS,
  kinds: reorderKinds,
};

function stockGrid(rows: StockLine[]): Grid {
  return rows.map((s) => [s.sku, s.item, s.category, s.onHand, s.reorder, s.qty, s.cost, s.orderCost]);
}

function freightGrid(rows: FreightRate[]): Grid {
  return rows.map((f) => [f.category, f.rate]);
}

const FEES = [0.25, 0.3, 0.35, 0.4, 0.45, 0.5];
const FREIGHT_RATES = [0.035, 0.04, 0.05, 0.06, 0.065, 0.075, 0.08, 0.09, 0.1, 0.12];

/** Cheap items are stocked by the thousand, expensive ones by the dozen. */
function stockTarget(rng: Rng, cost: number): number {
  if (cost < 1) return rng.int(150, 600) * 10;
  if (cost < 10) return rng.int(150, 800);
  if (cost < 40) return rng.int(40, 220);
  return rng.int(12, 70);
}

const withOrder = (s: StockLine): StockLine => {
  const qty = s.onHand <= s.reorder ? s.target - s.onHand : 0;
  return { ...s, qty, orderCost: round(qty * s.cost, 2) };
};

function onHandFor(rng: Rng, s: Pick<StockLine, 'reorder' | 'target'>): number {
  return rng.chance(0.35) ? rng.int(0, s.reorder) : rng.int(s.reorder + 1, s.target);
}

/** A Stock line for an item: an order-up-to level that suits its cost, a reorder point under it, and what's on hand. */
function stockLine(rng: Rng, sku: string, it: ItemDef, onHand: (s: Pick<StockLine, 'reorder' | 'target'>) => number): StockLine {
  const target = stockTarget(rng, it.cost);
  const reorder = Math.max(2, Math.round(target * rng.float(0.2, 0.35, 3)));
  const line = { sku, item: it.item, category: it.category, reorder, target, cost: it.cost };
  return withOrder({ ...line, onHand: onHand(line), qty: 0, orderCost: 0 });
}

/** Every category has something to reorder and something that doesn't need it, so every column has numbers to check. */
function balanceStock(stock: StockLine[], rng: Rng): StockLine[] {
  const out = stock.map((s) => ({ ...s }));
  for (const category of CATEGORIES) {
    const mine = out.filter((s) => s.category === category);
    if (!mine.some((s) => s.onHand <= s.reorder)) rng.pick(mine).onHand = 0;
    if (mine.every((s) => s.onHand <= s.reorder)) {
      const s = rng.pick(mine);
      s.onHand = rng.int(s.reorder + 1, s.target);
    }
  }
  return out.map(withOrder);
}

/** The Freight list in an order that isn't Z to A, so sorting it Z to A is a real change. */
function freightOrder(rng: Rng, rates: number[]): FreightRate[] {
  const rows = rng.shuffle(CATEGORIES).map((category, i) => ({ category, rate: rates[i] }));
  const descending = rows.every((r, i) => i === 0 || excelTextCompare(rows[i - 1].category, r.category) > 0);
  return descending ? [rows[1], rows[0], ...rows.slice(2)] : rows;
}

const byCategoryZtoA = (a: FreightRate, b: FreightRate) => excelTextCompare(b.category, a.category);

function reorderData(rng: Rng): RO {
  const numbers = rng
    .sample(span(1000, 1999), ITEMS.length)
    .sort((a, b) => a - b);
  const stock = rng.shuffle(ITEMS).map((it, i) => stockLine(rng, `SKU-${numbers[i]}`, it, (s) => onHandFor(rng, s)));
  const fee = rng.pick(FEES);
  const freight = freightOrder(rng, rng.sample(FREIGHT_RATES, CATEGORIES.length));
  return {
    stock: balanceStock(stock, rng),
    fee,
    freight,
    categories: rng.shuffle(CATEGORIES),
    asBuilt: { fee, rates: Object.fromEntries(freight.map((f) => [f.category, f.rate])) },
    templateEnd: ITEMS.length + 1 - rng.int(4, 7),
    bugs: [],
  };
}

/** One item per category that the Stock list doesn't carry yet. */
const NEW_ITEMS: readonly ItemDef[] = [
  { item: 'Void fill paper', category: 'Packaging', cost: 27.0 },
  { item: 'Wall anchor kit', category: 'Hardware', cost: 4.6 },
  { item: 'Power strip', category: 'Electrical', cost: 16.25 },
  { item: 'Safety glasses', category: 'Safety', cost: 3.9 },
  { item: 'Glass cleaner', category: 'Janitorial', cost: 4.2 },
];

/**
 * A new SKU in every category, each arriving with a little stock below its reorder point. Every
 * cell of every category row moves, so neither a count typed in nor a fixed range stretched to
 * today's last row keeps up.
 */
function withNewSkus(d: RO, rng: Rng): RO {
  const numbers = rng.sample(span(2000, 2999), NEW_ITEMS.length).sort((a, b) => a - b);
  const added = rng.shuffle(NEW_ITEMS).map((it, i) => stockLine(rng, `SKU-${numbers[i]}`, it, (s) => rng.int(1, s.reorder - 1)));
  return { ...d, stock: [...d.stock, ...added] };
}

const reorderVariants: Variant<RO>[] = [
  { label: 'the handling fee changes', apply: (d, rng) => ({ ...d, fee: rng.pick(FEES.filter((f) => f !== d.fee)) }) },
  { label: 'the freight rate list is sorted Z to A', apply: (d) => ({ ...d, freight: [...d.freight].sort(byCategoryZtoA) }) },
  {
    label: 'the freight rates change',
    apply: (d, rng) => ({ ...d, freight: d.freight.map((f) => ({ ...f, rate: round(f.rate + rng.pick([-0.015, -0.01, 0.01, 0.015, 0.02]), 3) })) }),
  },
  { label: 'stock levels change', apply: (d, rng) => ({ ...d, stock: balanceStock(d.stock.map((s) => ({ ...s, onHand: onHandFor(rng, s) })), rng) }) },
  { label: 'a new SKU is added in every category', apply: withNewSkus },
];

export const bughuntInventory = bugHunt<RO>({
  id: 'bughunt-inventory',
  title: 'Bug hunt: an AI-built reorder report',
  replaces: 'Sending an AI-built report on because the numbers look about right',
  minutes: 9,
  task: (d) =>
    `An AI assistant built this reorder report from last month’s template, and it has ${d.bugs.length} mistakes. Find and fix them in \`K2:S8\` so the report is right on the current data and stays right when the data changes. ${TOOLS} Change only the report’s formulas, not the Stock data, the fee in \`K10\` or the Freight list. Graded: every cell in \`K2:S8\` holds a formula and is right now, when the fee, the freight rates, their order or the stock levels change, and when new SKUs are added.`,
  concept: {
    ...HUNT_CONCEPT,
    example: 'In a column filled from one formula, every cell should read alike. =SUMIFS(…)*0.35 sitting among =SUMIFS(…)*$K$10 is a fee someone typed in.',
  },
  hints: [
    'Turn on Formulas › Show Formulas. Each column should repeat one pattern down the category rows, so a formula that reads differently from its neighbors is a suspect. Many of them carry a green error triangle. A triangle is a clue, not a verdict: the Average row may show one because it leaves out the Total row, and that’s correct.',
    kindsHint(
      'a number typed where a cell reference belongs, a fixed range sized for last month’s data, a lookup without an exact match, a subtraction the wrong way round, a total or average that takes in the wrong rows, a test on the wrong column, and a reference that slid or pointed at the wrong row',
    ),
    'Check the Total and Average rows, Freight, Handling and Share of order cost. In SKUs, On hand, Units to order and Order cost, look for fixed cell ranges in place of the Stock Table’s columns, and for a formula that reads another row’s category. In Units over reorder point and SKUs to reorder, check what each formula subtracts or tests.',
    'Use Formulas › Trace Precedents on each suspect. Every Handling cell should point at the fee in `K10`, every Freight cell at the Freight list with an exact match, every Share cell at the total in `P7`, and every row at its own category in column `J`. The answer lists each mistake with its cell.',
  ],
  report: reorderReport,
  data: reorderData,
  inputs: (d) => [tableWrite('Stock', STOCK_COLS, stockGrid(d.stock)), rangeWrite('K10', [[d.fee]]), tableWrite('Freight', FREIGHT_COLS, freightGrid(d.freight))],
  variants: reorderVariants,
});

// =====================================================================================
// Department budget vs actual (finance)
// =====================================================================================

interface LedgerLine {
  date: number;
  dept: string;
  account: string;
  amount: number;
}

interface BudgetLine {
  dept: string;
  budget: number;
}

export interface BudgetHuntData extends HuntData {
  ledger: LedgerLine[];
  /** The Budget list in sheet order, which isn't sorted. */
  budget: BudgetLine[];
  /** Payroll tax rate, in J1. */
  rate: number;
  /** Report column order, J3:P3. */
  depts: string[];
  asBuilt: { rate: number; budgets: Record<string, number> };
  /** Last row of last month’s ledger; a fixed range from the template stops here. */
  templateEnd: number;
}

type PL = BudgetHuntData;

const LEDGER_COLS: ColumnSpec[] = [{ header: 'Date', format: FMT.date }, { header: 'Dept' }, { header: 'Account' }, { header: 'Amount', format: FMT.currency }];
const BUDGET_COLS: ColumnSpec[] = [{ header: 'Dept' }, { header: 'Q3 budget', format: DOLLARS }];

const LEDGER = {
  dept: { name: 'Dept', letter: 'B', get: (l: LedgerLine) => l.dept },
  account: { name: 'Account', letter: 'C', get: (l: LedgerLine) => l.account },
  amount: { name: 'Amount', letter: 'D', get: (l: LedgerLine) => l.amount },
} satisfies Record<string, Field<LedgerLine>>;
const BUDGET = {
  dept: { name: 'Dept', letter: 'F', get: (b: BudgetLine) => b.dept },
  budget: { name: 'Q3 budget', letter: 'G', get: (b: BudgetLine) => b.budget },
} satisfies Record<string, Field<BudgetLine>>;

const ledgerTable: Source<PL, LedgerLine> = { table: 'Ledger', rows: (d) => d.ledger };
const budgetTable: Source<PL, BudgetLine> = { table: 'Budget', rows: (d) => d.budget };

const PL_RANGE = 'J4:P14';
const PL_COLS = ['J', 'K', 'L', 'M', 'N', 'O', 'P'];
const PL_LINES = [
  'Salaries and wages',
  'Payroll taxes',
  'Total personnel',
  'Rent',
  'Software subscriptions',
  'Travel',
  'Office supplies',
  'Total operating',
  'Total expenses',
  'Budget',
  'Over (under) budget',
] as const;
const PL_ROW = { salaries: 4, payroll: 5, personnel: 6, rent: 7, supplies: 10, operating: 11, total: 12, budget: 13, variance: 14 } as const;
/** Rows that sum one ledger account for the department. */
const PL_ACCOUNT_ROWS = [4, 7, 8, 9, 10];
const plLine = (row: number) => PL_LINES[row - PL_ROW.salaries];
const plDept = (d: PL, col: string) => d.depts[PL_COLS.indexOf(col)];

const deptSpend = (col: string, account: string, through?: number) =>
  ifs(ledgerTable, LEDGER.amount, [{ field: LEDGER.dept, criteria: `${col}$3` }, { field: LEDGER.account, criteria: account }], through);

function budgetColumn(col: string): Calc<PL>[] {
  return [
    deptSpend(col, '$I4'),
    op(ref(`${col}4`), '*', ref('$J$1')),
    sumOf(`${col}4:${col}5`),
    deptSpend(col, '$I7'),
    deptSpend(col, '$I8'),
    deptSpend(col, '$I9'),
    deptSpend(col, '$I10'),
    sumOf(`${col}7:${col}10`),
    op(ref(`${col}6`), '+', ref(`${col}11`)),
    xlookup(budgetTable, `${col}$3`, BUDGET.dept, BUDGET.budget),
    op(ref(`${col}12`), '-', ref(`${col}13`)),
  ];
}

const PL_BY_COL = PL_COLS.map(budgetColumn);
const PL_FORMULAS: Calc<PL>[][] = PL_LINES.map((_, r) => PL_BY_COL.map((col) => col[r]));
const plFormula = (col: string, row: number) => PL_BY_COL[PL_COLS.indexOf(col)][row - PL_ROW.salaries];

/** The department in the middle of the Budget list once it's sorted Z to A, where a binary search can land on the right row. */
const middleOfZtoA = (d: PL) => [...d.depts].sort((a, b) => excelTextCompare(b, a))[Math.floor(d.depts.length / 2)];

/** One slot per department column. */
const perDept = (make: (col: string, dept: string) => Slot<PL> | Slot<PL>[]) => (d: PL) => PL_COLS.flatMap((col) => make(col, plDept(d, col)));

const budgetKinds: BugKind<PL>[] = [
  {
    id: 'typed-input',
    hides: 'always',
    slots: (d) =>
      perDept((col, dept) => [
        {
          label: `Payroll taxes for ${dept} in ${col}5 had the rate typed in instead of pointing at J1`,
          formulas: { [`${col}5`]: op(ref(`${col}4`), '*', typed(d.asBuilt.rate)) },
          fix: writeFix(plFormula(col, PL_ROW.payroll), `${col}5`),
        },
        {
          label: `Over (under) budget for ${dept} in ${col}14 subtracted a typed budget instead of the Budget row`,
          formulas: { [`${col}14`]: op(ref(`${col}12`), '-', typed(d.asBuilt.budgets[dept])) },
          fix: writeFix(plFormula(col, PL_ROW.variance), `${col}14`),
        },
      ])(d),
  },
  {
    id: 'approx-lookup',
    hides: 'sometimes',
    slots: (d) =>
      perDept((col, dept) =>
        dept === middleOfZtoA(d)
          ? []
          : [
              {
                label: `Budget for ${dept} in ${col}13 used VLOOKUP’s approximate match (TRUE), which needs a sorted list`,
                formulas: { [`${col}13`]: vlookupApprox(budgetTable, `${col}$3`, BUDGET.dept, 2, BUDGET.budget) },
                fix: writeFix(plFormula(col, PL_ROW.budget), `${col}13`),
              },
            ],
      )(d),
  },
  {
    id: 'sign-flip',
    slots: perDept((col, dept) => ({
      label: `Over (under) budget for ${dept} in ${col}14 subtracted the wrong way round`,
      formulas: { [`${col}14`]: op(ref(`${col}13`), '-', ref(`${col}12`)) },
      fix: writeFix(plFormula(col, PL_ROW.variance), `${col}14`),
    })),
  },
  {
    id: 'double-count',
    slots: perDept((col, dept) => ({
      label: `Total expenses for ${dept} in ${col}12 added the subtotals on top of the lines they total`,
      formulas: { [`${col}12`]: sumOf(`${col}4:${col}11`) },
      fix: writeFix(plFormula(col, PL_ROW.total), `${col}12`),
    })),
  },
  {
    id: 'short-total',
    slots: perDept((col, dept) => ({
      label: `Total operating for ${dept} in ${col}11 stopped before Office supplies`,
      formulas: { [`${col}11`]: sumOf(`${col}7:${col}9`) },
      fix: writeFix(plFormula(col, PL_ROW.operating), `${col}11`),
    })),
  },
  {
    id: 'off-by-one',
    slots: perDept((col, dept) =>
      PL_ACCOUNT_ROWS.flatMap((r) =>
        [-1, 1].map((step) => ({
          label: `${plLine(r)} for ${dept} in ${col}${r} read the account from the row ${step > 0 ? 'below' : 'above'}`,
          formulas: { [`${col}${r}`]: deptSpend(col, `$I${r + step}`) },
          fix: writeFix(plFormula(col, r), `${col}${r}`),
        })),
      ),
    ),
  },
  {
    id: 'wrong-criteria',
    slots: perDept((col, dept) =>
      PL_ACCOUNT_ROWS.map((r) => ({
        label: `${plLine(r)} for ${dept} in ${col}${r} looked for the department in the Account column and the account in the Dept column`,
        formulas: { [`${col}${r}`]: ifs(ledgerTable, LEDGER.amount, [{ field: LEDGER.dept, criteria: `$I${r}` }, { field: LEDGER.account, criteria: `${col}$3` }]) },
        fix: writeFix(plFormula(col, r), `${col}${r}`),
      })),
    ),
  },
  {
    id: 'short-range',
    slots: (d) =>
      perDept((col, dept) =>
        PL_ACCOUNT_ROWS.map((r) => ({
          label: `${plLine(r)} for ${dept} in ${col}${r} used a fixed range that stops at row ${d.templateEnd}, so it missed the ledger entries below it`,
          formulas: { [`${col}${r}`]: deptSpend(col, `$I${r}`, d.templateEnd) },
          fix: writeFix(plFormula(col, r), `${col}${r}`),
        })),
      )(d),
  },
];

const budgetReport: Report<PL> = {
  range: PL_RANGE,
  formats: PL_COLS.map(() => DOLLARS),
  sheet: (d) => [
    dataBlock('Ledger', 'A1', LEDGER_COLS, ledgerGrid(d.ledger)),
    dataBlock('Budget', 'F1', BUDGET_COLS, budgetGrid(d.budget)),
    cells('I1', [['Payroll tax rate']], 'label'),
    cells('J1', [[d.rate]], 'input', '0.00%'),
    cells('I3', [['Line', ...d.depts]], 'header'),
    cells('I4', column([...PL_LINES]), 'label'),
    cells('R3', [[AI_NOTE]], 'label'),
  ],
  formulas: () => PL_FORMULAS,
  kinds: budgetKinds,
};

function ledgerGrid(rows: LedgerLine[]): Grid {
  return rows.map((l) => [l.date, l.dept, l.account, l.amount]);
}

function budgetGrid(rows: BudgetLine[]): Grid {
  return rows.map((b) => [b.dept, b.budget]);
}

const PAYROLL_RATES = [0.0765, 0.079, 0.081, 0.084];
const Q3_MONTHS = [7, 8, 9];

/** Q3 postings for one department: payroll at each month end, rent on the 1st, and the odd bill. */
function deptLedger(rng: Rng, dept: string): LedgerLine[] {
  const day = (month: number) => serial(2026, month, rng.int(2, 28));
  const lines: LedgerLine[] = [];
  const payroll = rng.int(18, 70) * 1000;
  const rent = rng.int(12, 60) * 100;
  for (const m of Q3_MONTHS) {
    lines.push({ date: serial(2026, m + 1, 0), dept, account: 'Salaries and wages', amount: round(payroll * rng.float(0.97, 1.03, 3), 2) });
    lines.push({ date: serial(2026, m, 1), dept, account: 'Rent', amount: rent });
  }
  for (let i = rng.int(1, 3); i > 0; i--) lines.push({ date: day(rng.pick(Q3_MONTHS)), dept, account: 'Software subscriptions', amount: rng.float(180, 2400, 2) });
  for (let i = rng.int(0, 3); i > 0; i--) lines.push({ date: day(rng.pick(Q3_MONTHS)), dept, account: 'Travel', amount: rng.float(120, 3800, 2) });
  for (let i = rng.int(1, 2); i > 0; i--) lines.push({ date: day(rng.pick(Q3_MONTHS)), dept, account: 'Office supplies', amount: rng.float(45, 900, 2) });
  return lines;
}

/** What the department actually spent, payroll taxes included: the Total expenses row. */
function deptTotal(d: Pick<PL, 'ledger' | 'rate'>, dept: string): number {
  const of = (account: string) => sum(d.ledger.filter((l) => l.dept === dept && l.account === account).map((l) => l.amount));
  const salaries = of('Salaries and wages');
  return salaries + salaries * d.rate + of('Rent') + of('Software subscriptions') + of('Travel') + of('Office supplies');
}

/** A budget near actual spend, in round $500s, and never exactly on it. */
function budgetFor(rng: Rng, actual: number): number {
  const b = Math.round((actual * rng.float(0.88, 1.12, 3)) / 500) * 500;
  return Math.abs(b - actual) < 1 ? b + 500 : b;
}

/** The Budget list in an order that isn't Z to A, so sorting it Z to A is a real change. */
function budgetOrder(rng: Rng, rows: BudgetLine[]): BudgetLine[] {
  const out = rng.shuffle(rows);
  const descending = out.every((r, i) => i === 0 || excelTextCompare(out[i - 1].dept, r.dept) > 0);
  return descending ? [out[1], out[0], ...out.slice(2)] : out;
}

function budgetData(rng: Rng): PL {
  const dropped = rng.pick(DEPARTMENTS);
  const depts: string[] = DEPARTMENTS.filter((x) => x !== dropped);
  const ledger = depts
    .flatMap((dept) => deptLedger(rng, dept))
    .sort((a, b) => a.date - b.date || depts.indexOf(a.dept) - depts.indexOf(b.dept));
  const rate = rng.pick(PAYROLL_RATES);
  const budget = budgetOrder(
    rng,
    depts.map((dept) => ({ dept, budget: budgetFor(rng, deptTotal({ ledger, rate }, dept)) })),
  );
  return {
    ledger,
    budget,
    rate,
    depts,
    asBuilt: { rate, budgets: Object.fromEntries(budget.map((b) => [b.dept, b.budget])) },
    templateEnd: ledger.length + 1 - rng.int(5, 9),
    bugs: [],
  };
}

/** The range a quarter-end accrual falls in, by account. */
const ACCRUALS: Record<string, [number, number]> = {
  'Salaries and wages': [400, 2500],
  Rent: [100, 600],
  'Software subscriptions': [180, 2400],
  Travel: [120, 3800],
  'Office supplies': [45, 900],
};

/**
 * A quarter-end accrual for every department and account. Every account cell moves, so neither a
 * number typed in nor a fixed range stretched to today's last row keeps up.
 */
function withAccruals(d: PL, rng: Rng): PL {
  const close = serial(2026, 9, 30);
  const added = d.depts.flatMap((dept) =>
    PL_ACCOUNT_ROWS.map((r): LedgerLine => {
      const [lo, hi] = ACCRUALS[plLine(r)];
      return { date: close, dept, account: plLine(r), amount: rng.float(lo, hi, 2) };
    }),
  );
  return { ...d, ledger: [...d.ledger, ...added] };
}

const budgetVariants: Variant<PL>[] = [
  { label: 'the payroll tax rate changes', apply: (d, rng) => ({ ...d, rate: rng.pick(PAYROLL_RATES.filter((r) => r !== d.rate)) }) },
  { label: 'the budget list is sorted Z to A', apply: (d) => ({ ...d, budget: [...d.budget].sort((a, b) => excelTextCompare(b.dept, a.dept)) }) },
  {
    label: 'the budgets change',
    apply: (d, rng) => ({
      ...d,
      budget: d.budget.map((b) => {
        const next = Math.round((b.budget * rng.float(0.9, 1.1, 3)) / 500) * 500;
        return { ...b, budget: next === b.budget ? next + 1000 : next };
      }),
    }),
  },
  { label: 'the ledger amounts change', apply: (d, rng) => ({ ...d, ledger: d.ledger.map((l) => ({ ...l, amount: round(l.amount * rng.float(0.85, 1.15, 3), 2) })) }) },
  { label: 'quarter-end accruals are posted', apply: withAccruals },
];

export const bughuntBudget = bugHunt<PL>({
  id: 'bughunt-pnl',
  title: 'Bug hunt: an AI-built budget vs actual',
  replaces: 'Forwarding a department P&L without checking how its totals were built',
  minutes: 10,
  task: (d) =>
    `An AI assistant built this Q3 budget vs actual report from last month’s template, and it has ${d.bugs.length} mistakes. Find and fix them in \`J4:P14\` so every department’s column is right on the current data and stays right when the data changes. ${TOOLS} Change only the report’s formulas, not the Ledger, the Budget list or the payroll tax rate in \`J1\`. Graded: every cell in \`J4:P14\` holds a formula and is right now, when the rate, the budgets, their order or the ledger amounts change, and when quarter-end accruals are posted.`,
  concept: {
    ...HUNT_CONCEPT,
    example: 'Each row should hold one formula filled across the departments. =M12-48500 sitting among =L12-L13 and =N12-N13 is a budget someone typed in.',
  },
  hints: [
    'Turn on Formulas › Show Formulas. Each row should repeat one formula across the departments, so a cell that reads differently from its neighbors is a suspect. Many of them carry a green error triangle. A triangle is a clue, not a verdict: some mistakes don’t get one, and a correct total can.',
    kindsHint(
      'a number typed where a cell reference belongs, a fixed range sized for last month’s ledger, a lookup without an exact match, a subtraction the wrong way round, a total that adds the wrong rows, criteria matched against the wrong columns, and a reference that pointed at the wrong row',
    ),
    'Check Payroll taxes in row `5`, Total operating and Total expenses in rows `11` and `12`, Budget in row `13` and Over (under) budget in row `14`. In the account rows (`4` and `7` to `10`), look for fixed cell ranges in place of the Ledger Table’s columns, a formula that reads another row’s account, and criteria matched against the wrong columns.',
    'Use Formulas › Trace Precedents on each suspect. Payroll taxes should point at the rate in `J1`, Total expenses should add only the two subtotals, Budget should look up its department with an exact match, and each account row should read its own label in column `I` and its department in row `3`. The answer lists each mistake with its cell.',
  ],
  report: budgetReport,
  data: budgetData,
  inputs: (d) => [tableWrite('Ledger', LEDGER_COLS, ledgerGrid(d.ledger)), tableWrite('Budget', BUDGET_COLS, budgetGrid(d.budget)), rangeWrite('J1', [[d.rate]])],
  variants: budgetVariants,
});

// =====================================================================================
// Sales commission summary (sales)
// =====================================================================================

interface Deal {
  date: number;
  rep: string;
  region: string;
  type: 'New' | 'Renewal';
  amount: number;
}

interface QuotaLine {
  rep: string;
  quota: number;
}

export interface CommissionHuntData extends HuntData {
  deals: Deal[];
  /** The Quotas list in sheet order, which isn't sorted. */
  quotas: QuotaLine[];
  /** Commission rate, in H13. */
  rate: number;
  regions: string[];
  /** Report order: the first region's three reps, then the second's. */
  reps: string[];
  asBuilt: { rate: number; quotas: Record<string, number> };
  /** Last row of last month’s Sales export; a fixed range from the template stops here. */
  templateEnd: number;
}

type CM = CommissionHuntData;

const SALES_COLS: ColumnSpec[] = [{ header: 'Date', format: FMT.date }, { header: 'Rep' }, { header: 'Region' }, { header: 'Type' }, { header: 'Amount', format: DOLLARS }];
const QUOTA_COLS: ColumnSpec[] = [{ header: 'Rep' }, { header: 'Quota', format: DOLLARS }];

const SALES = {
  rep: { name: 'Rep', letter: 'B', get: (s: Deal) => s.rep },
  region: { name: 'Region', letter: 'C', get: (s: Deal) => s.region },
  type: { name: 'Type', letter: 'D', get: (s: Deal) => s.type },
  amount: { name: 'Amount', letter: 'E', get: (s: Deal) => s.amount },
} satisfies Record<string, Field<Deal>>;
const QUOTA = {
  rep: { name: 'Rep', letter: 'G', get: (q: QuotaLine) => q.rep },
  quota: { name: 'Quota', letter: 'H', get: (q: QuotaLine) => q.quota },
} satisfies Record<string, Field<QuotaLine>>;

const salesTable: Source<CM, Deal> = { table: 'Sales', rows: (d) => d.deals };
const quotaTable: Source<CM, QuotaLine> = { table: 'Quotas', rows: (d) => d.quotas };

const CM_RANGE = 'H2:O11';
/** Each region: three rep rows and a subtotal. Then Total, and Average per rep. The rate sits in H13. */
const CM_BLOCKS = [
  { first: 2, last: 4, subtotal: 5 },
  { first: 6, last: 8, subtotal: 9 },
] as const;
const CM_ROWS = { total: 10, average: 11 } as const;
const CM_REP_ROWS = CM_BLOCKS.flatMap((b) => span(b.first, b.last));
const CM_COLUMNS = [
  { col: 'H', header: 'Deals', format: FMT.int },
  { col: 'I', header: 'Sales', format: DOLLARS },
  { col: 'J', header: 'New business', format: DOLLARS },
  { col: 'K', header: 'Renewals', format: DOLLARS },
  { col: 'L', header: 'Quota', format: DOLLARS },
  { col: 'M', header: 'Over (under) quota', format: DOLLARS },
  { col: 'N', header: 'Commission', format: FMT.currency },
  { col: 'O', header: 'Share of sales', format: FMT.pct },
] as const;
const CM_MEASURES = CM_COLUMNS.slice(0, 7).map((c) => c.col);
const cmHeader = (col: string) => CM_COLUMNS.find((c) => c.col === col)!.header;
const cmRep = (d: CM, row: number) => d.reps[CM_REP_ROWS.indexOf(row)];

const repSales = (label: string, through?: number) => ifs(salesTable, SALES.amount, [{ field: SALES.rep, criteria: label }], through);
const repQuota = (label: string) => xlookup(quotaTable, label, QUOTA.rep, QUOTA.quota);

/** Columns that tally the Sales Table per rep; off-by-one and short-range mistakes land here. */
const CM_TALLIES: Record<string, (label: string, through?: number) => Calc<CM>> = {
  H: (label, through) => ifs(salesTable, null, [{ field: SALES.rep, criteria: label }], through),
  I: (label, through) => repSales(label, through),
  J: (label, through) => ifs(salesTable, SALES.amount, [{ field: SALES.rep, criteria: label }, { field: SALES.type, criteria: '"New"' }], through),
  K: (label, through) => ifs(salesTable, SALES.amount, [{ field: SALES.rep, criteria: label }, { field: SALES.type, criteria: '"Renewal"' }], through),
  N: (label, through) => op(repSales(label, through), '*', ref('$H$13')),
};

const share = (row: number) => op(ref(`I${row}`), '/', ref('I$10'));

function commissionFormulas(): Calc<CM>[][] {
  const rows: Calc<CM>[][] = [];
  for (const block of CM_BLOCKS) {
    for (const r of span(block.first, block.last)) {
      const label = `$G${r}`;
      const t = (col: string) => CM_TALLIES[col](label);
      rows.push([t('H'), t('I'), t('J'), t('K'), repQuota(label), op(repSales(label), '-', repQuota(label)), t('N'), share(r)]);
    }
    rows.push([...CM_MEASURES.map((c) => sumOf(`${c}${block.first}:${c}${block.last}`)), share(block.subtotal)]);
  }
  rows.push([...CM_MEASURES.map((c) => op(ref(`${c}5`), '+', ref(`${c}9`))), share(CM_ROWS.total)]);
  rows.push([...CM_MEASURES.map((c) => averageOf(`${c}2:${c}4`, `${c}6:${c}8`)), share(CM_ROWS.average)]);
  return rows;
}

const CM_FORMULAS = commissionFormulas();
const cmFormula = (address: string) => {
  const cell = parseCell(address);
  return CM_FORMULAS[cell.row - 2][cell.col - parseCell('H1').col];
};

const commissionKinds: BugKind<CM>[] = [
  {
    id: 'typed-input',
    hides: 'always',
    slots: (d) => [
      ...CM_BLOCKS.flatMap((b) =>
        span(b.first, b.last).map((r) => ({
          label:
            r === b.last
              ? `Commission for ${cmRep(d, r)} in N${r} had the rate typed in instead of pointing at H13`
              : `Commission in N${r}:N${b.last} had the rate typed in instead of pointing at H13`,
          formulas: Object.fromEntries(span(r, b.last).map((k) => [`N${k}`, op(repSales(`$G${k}`), '*', typed(d.asBuilt.rate))])),
          fix: writeFix(cmFormula(`N${r}`), `N${r}`, `N${b.last}`),
        })),
      ),
      ...CM_REP_ROWS.map((r) => ({
        label: `Over (under) quota for ${cmRep(d, r)} in M${r} subtracted a typed quota instead of looking it up`,
        formulas: { [`M${r}`]: op(repSales(`$G${r}`), '-', typed(d.asBuilt.quotas[cmRep(d, r)])) },
        fix: writeFix(cmFormula(`M${r}`), `M${r}`),
      })),
    ],
  },
  {
    id: 'fill-drift',
    slots: () =>
      CM_BLOCKS.map((b) => ({
        label: `Commission in N${b.first}:N${b.last} pointed at H13 without $ signs, so the filled rows read the cells below it`,
        formulas: Object.fromEntries(span(b.first, b.last).map((k) => [`N${k}`, op(repSales(`$G${k}`), '*', ref(`H${13 + k - b.first}`))])),
        fix: writeFix(cmFormula(`N${b.first}`), `N${b.first}`, `N${b.last}`),
      })),
  },
  {
    id: 'approx-lookup',
    hides: 'sometimes',
    slots: () => [
      {
        label: 'Quota in L2:L4 and L6:L8 used VLOOKUP’s approximate match (TRUE), which needs a sorted list',
        formulas: Object.fromEntries(CM_REP_ROWS.map((r) => [`L${r}`, vlookupApprox(quotaTable, `$G${r}`, QUOTA.rep, 2, QUOTA.quota)])),
        fix: `Write =${cmFormula('L2').text} in L2, fill down to L4, and copy it to L6:L8.`,
      },
    ],
  },
  {
    id: 'sign-flip',
    slots: (d) =>
      CM_REP_ROWS.map((r) => ({
        label: `Over (under) quota for ${cmRep(d, r)} in M${r} subtracted the wrong way round`,
        formulas: { [`M${r}`]: op(repQuota(`$G${r}`), '-', repSales(`$G${r}`)) },
        fix: writeFix(cmFormula(`M${r}`), `M${r}`),
      })),
  },
  {
    id: 'double-count',
    slots: () =>
      CM_MEASURES.map((col) => ({
        label: `The ${cmHeader(col)} total in ${col}10 added the region totals on top of the reps`,
        formulas: { [`${col}10`]: sumOf(`${col}2:${col}9`) },
        fix: writeFix(cmFormula(`${col}10`), `${col}10`),
      })),
  },
  {
    id: 'average-total',
    slots: () =>
      CM_MEASURES.map((col) => ({
        label: `Average per rep for ${cmHeader(col)} in ${col}11 took in the total rows`,
        formulas: { [`${col}11`]: averageOf(`${col}2:${col}10`) },
        fix: writeFix(cmFormula(`${col}11`), `${col}11`),
      })),
  },
  {
    id: 'short-total',
    slots: (d) =>
      CM_BLOCKS.flatMap((b, i) =>
        CM_MEASURES.map((col) => ({
          label: `The ${d.regions[i]} total for ${cmHeader(col)} in ${col}${b.subtotal} left out ${cmRep(d, b.last)}`,
          formulas: { [`${col}${b.subtotal}`]: sumOf(`${col}${b.first}:${col}${b.last - 1}`) },
          fix: writeFix(cmFormula(`${col}${b.subtotal}`), `${col}${b.subtotal}`),
        })),
      ),
  },
  {
    id: 'off-by-one',
    slots: (d) =>
      Object.entries(CM_TALLIES).flatMap(([col, tally]) =>
        CM_REP_ROWS.flatMap((r) =>
          [-1, 1].map((step) => ({
            label: `${cmHeader(col)} for ${cmRep(d, r)} in ${col}${r} read the rep from the row ${step > 0 ? 'below' : 'above'}`,
            formulas: { [`${col}${r}`]: tally(`$G${r + step}`) },
            fix: writeFix(cmFormula(`${col}${r}`), `${col}${r}`),
          })),
        ),
      ),
  },
  {
    id: 'wrong-criteria',
    slots: (d) =>
      (['J', 'K'] as const).flatMap((col) =>
        CM_REP_ROWS.map((r) => {
          const type = col === 'J' ? 'New' : 'Renewal';
          return {
            label: `${cmHeader(col)} for ${cmRep(d, r)} in ${col}${r} looked for “${type}” in the Region column instead of Type`,
            formulas: { [`${col}${r}`]: ifs(salesTable, SALES.amount, [{ field: SALES.rep, criteria: `$G${r}` }, { field: SALES.region, criteria: `"${type}"` }]) },
            fix: writeFix(cmFormula(`${col}${r}`), `${col}${r}`),
          };
        }),
      ),
  },
  {
    id: 'share-drift',
    slots: () => [
      {
        label: 'Share of sales in O2:O11 divided by I10 without a $ sign, so the total slid down as the formula filled',
        formulas: Object.fromEntries(span(2, CM_ROWS.average).map((k) => [`O${k}`, op(ref(`I${k}`), '/', ref(`I${k + 8}`))])),
        fix: writeFix(cmFormula('O2'), 'O2', 'O11'),
      },
    ],
  },
  {
    id: 'short-range',
    slots: (d) =>
      Object.entries(CM_TALLIES).flatMap(([col, tally]) =>
        CM_REP_ROWS.map((r) => ({
          label: `${cmHeader(col)} for ${cmRep(d, r)} in ${col}${r} used a fixed range that stops at row ${d.templateEnd}, so it missed the deals below it`,
          formulas: { [`${col}${r}`]: tally(`$G${r}`, d.templateEnd) },
          fix: writeFix(cmFormula(`${col}${r}`), `${col}${r}`),
        })),
      ),
  },
];

const commissionReport: Report<CM> = {
  range: CM_RANGE,
  formats: CM_COLUMNS.map((c) => c.format),
  sheet: (d) => [
    dataBlock('Sales', 'A1', SALES_COLS, dealGrid(d.deals)),
    cells('G1', [['Rep', ...CM_COLUMNS.map((c) => c.header)]], 'header'),
    cells(
      'G2',
      column([...d.reps.slice(0, 3), `${d.regions[0]} total`, ...d.reps.slice(3), `${d.regions[1]} total`, 'Total', 'Average per rep']),
      'label',
    ),
    cells('G13', [['Commission rate']], 'label'),
    cells('H13', [[d.rate]], 'input', FMT.pct),
    dataBlock('Quotas', 'G15', QUOTA_COLS, quotaGrid(d.quotas)),
    cells('Q1', [[AI_NOTE]], 'label'),
  ],
  formulas: () => CM_FORMULAS,
  kinds: commissionKinds,
};

function dealGrid(rows: Deal[]): Grid {
  return rows.map((s) => [s.date, s.rep, s.region, s.type, s.amount]);
}

function quotaGrid(rows: QuotaLine[]): Grid {
  return rows.map((q) => [q.rep, q.quota]);
}

const COMMISSION_RATES = [0.04, 0.045, 0.05, 0.055, 0.06];
const dealAmount = (rng: Rng) => rng.int(40, 960) * 50;

/** Quotas near each rep's sales, in round thousands, never exactly on them and never two the same. */
function quotasFor(rng: Rng, deals: Deal[], reps: string[]): QuotaLine[] {
  const taken = new Set<number>();
  return reps.map((rep) => {
    const sales = sum(deals.filter((s) => s.rep === rep).map((s) => s.amount));
    let quota = Math.round((sales * rng.float(0.8, 1.25, 3)) / 1000) * 1000;
    while (quota === sales || taken.has(quota)) quota += 1000;
    taken.add(quota);
    return { rep, quota };
  });
}

/** The Quotas list in an order that isn't Z to A, so sorting it Z to A is a real change. */
function quotaOrder(rng: Rng, rows: QuotaLine[]): QuotaLine[] {
  const out = rng.shuffle(rows);
  const descending = out.every((r, i) => i === 0 || excelTextCompare(out[i - 1].rep, r.rep) > 0);
  return descending ? [out[1], out[0], ...out.slice(2)] : out;
}

function commissionData(rng: Rng): CM {
  const regions = rng.sample(REGIONS, 2);
  const team: string[] = rng.sample(REPS, 6);
  const reps = [...team.slice(0, 3).sort(excelTextCompare), ...team.slice(3).sort(excelTextCompare)];
  const regionOf = (rep: string) => regions[reps.indexOf(rep) < 3 ? 0 : 1];
  const start = serial(2026, 7, 1);
  const deals = reps
    .flatMap((rep) =>
      Array.from({ length: rng.int(5, 8) }, (_, i): Deal => ({
        date: start + rng.int(0, 91),
        rep,
        region: regionOf(rep),
        // Every rep has at least one of each, so both breakdown columns have numbers to check.
        type: i === 0 ? 'New' : i === 1 ? 'Renewal' : rng.chance(0.4) ? 'New' : 'Renewal',
        amount: dealAmount(rng),
      })),
    )
    .sort((a, b) => a.date - b.date || excelTextCompare(a.rep, b.rep));
  const rate = rng.pick(COMMISSION_RATES);
  const quotas = quotaOrder(rng, quotasFor(rng, deals, reps));
  return {
    deals,
    quotas,
    rate,
    regions,
    reps,
    asBuilt: { rate, quotas: Object.fromEntries(quotas.map((q) => [q.rep, q.quota])) },
    templateEnd: deals.length + 1 - rng.int(4, 7),
    bugs: [],
  };
}

/**
 * Every rep closes a new deal and a renewal on the quarter's last day. Every cell of every rep row
 * moves, so neither a count typed in nor a fixed range stretched to today's last row keeps up.
 */
function withNewDeals(d: CM, rng: Rng): CM {
  const last = serial(2026, 9, 30);
  const added = d.reps.flatMap((rep, i) =>
    (['New', 'Renewal'] as const).map((type): Deal => ({ date: last, rep, region: d.regions[i < 3 ? 0 : 1], type, amount: dealAmount(rng) })),
  );
  return { ...d, deals: [...d.deals, ...rng.shuffle(added)] };
}

const commissionVariants: Variant<CM>[] = [
  { label: 'the commission rate changes', apply: (d, rng) => ({ ...d, rate: rng.pick(COMMISSION_RATES.filter((r) => r !== d.rate)) }) },
  { label: 'the quota list is sorted Z to A', apply: (d) => ({ ...d, quotas: [...d.quotas].sort((a, b) => excelTextCompare(b.rep, a.rep)) }) },
  { label: 'the quotas change', apply: (d, rng) => ({ ...d, quotas: d.quotas.map((q) => ({ ...q, quota: q.quota + rng.pick([-6000, -4000, -3000, 3000, 4000, 6000]) })) }) },
  { label: 'the deal amounts change', apply: (d, rng) => ({ ...d, deals: d.deals.map((s) => ({ ...s, amount: dealAmount(rng) })) }) },
  { label: 'every rep closes two more deals', apply: withNewDeals },
];

export const bughuntCommission = bugHunt<CM>({
  id: 'bughunt-commission',
  title: 'Bug hunt: an AI-built commission summary',
  replaces: 'Paying commission off a summary nobody traced back to the deals',
  minutes: 9,
  task: (d) =>
    `An AI assistant built this quarter’s commission summary from last month’s template, and it has ${d.bugs.length} mistakes. Find and fix them in \`H2:O11\` so every rep’s numbers are right on the current data and stay right when the data changes. ${TOOLS} Change only the report’s formulas, not the Sales data, the Quotas list or the commission rate in \`H13\`. Graded: every cell in \`H2:O11\` holds a formula and is right now, when the rate, the quotas, their order or the deal amounts change, and when new deals close.`,
  concept: {
    ...HUNT_CONCEPT,
    example: 'A subtotal of =SUM(I2:I3) above three rep rows stops one row short; the one below the other region reads =SUM(I6:I8).',
  },
  hints: [
    'Turn on Formulas › Show Formulas. Each column should repeat one pattern down the rep rows, and the subtotal, Total and Average rows should read alike across the columns. A formula that breaks the pattern is a suspect, and many carry a green error triangle. A triangle is a clue, not a verdict: the Average row may show one because it leaves out the subtotal and Total rows, and that’s correct.',
    kindsHint(
      'a number typed where a cell reference belongs, a fixed range sized for last month’s data, a lookup without an exact match, a subtraction the wrong way round, a total or average that takes in the wrong rows, a test on the wrong column, and a reference that slid or pointed at the wrong row',
    ),
    'Check the region subtotals, the Total and Average rows, Quota, Over (under) quota, Commission and Share of sales. In Deals, Sales, New business, Renewals and Commission, look for fixed cell ranges in place of the Sales Table’s columns, and for a formula that reads another row’s rep. In New business and Renewals, check which column each formula tests.',
    'Use Formulas › Trace Precedents on each suspect. Commission should point at the rate in `H13`, Quota should look up its rep with an exact match, every Share cell should divide by the total in `I10`, and every rep row should read its own name in column `G`. The answer lists each mistake with its cell.',
  ],
  report: commissionReport,
  data: commissionData,
  inputs: (d) => [tableWrite('Sales', SALES_COLS, dealGrid(d.deals)), rangeWrite('H13', [[d.rate]]), tableWrite('Quotas', QUOTA_COLS, quotaGrid(d.quotas))],
  variants: commissionVariants,
});

// =====================================================================================

export const BUG_HUNTS: Exercise<any>[] = [bughuntInventory, bughuntBudget, bughuntCommission];

/** Each hunt's report, for tests that work the formulas out independently. */
export const HUNT_REPORTS: Record<string, Report<any>> = {
  [bughuntInventory.id]: reorderReport,
  [bughuntBudget.id]: budgetReport,
  [bughuntCommission.id]: commissionReport,
};

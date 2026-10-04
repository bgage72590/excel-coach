import { describe, expect, it } from 'vitest';
import { cellAddress, colToNumber, parseCell, parseRange, type RangeRef } from '../src/engine/address';
import { excelTextCompare } from '../src/engine/data';
import { localize } from '../src/engine/platform';
import { emptyProgress, recordDrillRun } from '../src/engine/progress';
import { Rng } from '../src/engine/rng';
import { gradeSheetCheck, hiddenTarget, type SheetFacts } from '../src/engine/sheetChecks';
import type { Cell, CheckReport, CheckStatus, Exercise, ExpectedGrid, Grid, Inspection, Layout, Platform, SheetCheck } from '../src/engine/types';
import {
  DRILLS,
  drillItemId,
  elapsedMs,
  formatDrillGap,
  formatDrillTime,
  getDrill,
  nextAction,
  pauseClock,
  recordRun,
  resumeClock,
  runTotalMs,
  runningSince,
  startClock,
  uncheckable,
  type DrillItem,
  type ItemTime,
} from '../src/drills';
import { NEW_WAREHOUSE, OLD_WAREHOUSE } from '../src/drills/clean';
import {
  CURRENCY_FORMAT,
  DATE_FORMAT,
  GENERAL_FORMAT,
  PERCENT_FORMAT,
  RED_NEGATIVE_FORMAT,
  THOUSANDS_FORMAT,
  TIME_FORMAT,
  removeDuplicateRows,
} from '../src/drills/common';
import { CUSTOMER_MIN_WIDTH, TABLE_NAME } from '../src/drills/setup';
import { SEEDS, blockRange, checkLayout, overlaps } from './helpers/suite';

const ITEMS: DrillItem[] = DRILLS.flatMap((d) => d.items);
const MANY_SEEDS = Array.from({ length: 120 }, (_, i) => i * 7919 + 11);
const PLATFORMS: Platform[] = ['mac', 'windows', 'web'];
const BANNED = /\b(please|simply|just|easy|easily|successfully|leverage|seamless)\b|!/i;
/** The widest a default column gets (points), across the fonts Excel ships with. */
const DEFAULT_WIDTH = 65;

type SheetInspection = Extract<Inspection, { kind: 'sheet' }>;

const item = (id: string): DrillItem => {
  const found = ITEMS.find((i) => drillItemId(i) === id);
  if (!found) throw new Error(`No drill item ${id}`);
  return found;
};

const sheetChecks = (ex: Exercise<any>, d: unknown): SheetInspection[] => (ex.inspections?.(d) ?? []).filter((i): i is SheetInspection => i.kind === 'sheet');

function checkOf<K extends SheetCheck['kind']>(ex: Exercise<any>, d: unknown, kind: K): Extract<SheetCheck, { kind: K }>[] {
  return sheetChecks(ex, d)
    .map((i) => i.check)
    .filter((c): c is Extract<SheetCheck, { kind: K }> => c.kind === kind);
}

// ---------- a model of the practice sheet ----------

interface ModelCell {
  value: Cell;
  formula: Cell;
  format: string;
}

type Cells = Map<string, ModelCell>;

/** The value of a drill formula. The only ones drills write are products of two cells, like =C2*D2. */
function evaluate(cells: Cells, formula: Cell): number {
  const m = /^=([A-Z]+\d+)\*([A-Z]+\d+)$/.exec(String(formula));
  if (!m) throw new Error(`Can’t evaluate ${String(formula)}`);
  return Number(cells.get(m[1])?.value) * Number(cells.get(m[2])?.value);
}

/** Every cell the coach writes, keyed by A1 address, with the value Excel would show. */
function writtenSheet(layout: Layout): Cells {
  const sheet: Cells = new Map();
  for (const b of layout.blocks) {
    const start = parseCell(b.at);
    const put = (r: number, c: number, cell: ModelCell) => sheet.set(cellAddress({ row: start.row + r, col: start.col + c }), cell);
    if (b.kind === 'data') {
      b.columns.forEach((col, c) => put(0, c, { value: col.header, formula: col.header, format: 'General' }));
      b.rows.forEach((row, r) => row.forEach((v, c) => put(r + 1, c, { value: v, formula: v, format: b.columns[c].format ?? 'General' })));
    } else {
      b.values.forEach((row, r) =>
        row.forEach((v, c) => put(r, c, { value: b.role === 'formula' ? null : v, formula: v, format: b.formats?.[c] ?? b.format ?? 'General' })),
      );
    }
  }
  for (const cell of sheet.values()) if (cell.value === null) cell.value = evaluate(sheet, cell.formula);
  return sheet;
}

function cellsOf(range: string): string[][] {
  const r = parseRange(range);
  return Array.from({ length: r.end.row - r.start.row + 1 }, (_, i) => Array.from({ length: r.end.col - r.start.col + 1 }, (_, j) => cellAddress({ row: r.start.row + i, col: r.start.col + j })));
}

const read = (sheet: Cells, range: string, field: keyof ModelCell): Grid =>
  cellsOf(range).map((row) => row.map((a) => sheet.get(a)?.[field] ?? (field === 'format' ? 'General' : '')));

const plainGrid = (expected: ExpectedGrid): Grid => expected.map((row) => row.map((c) => c as Cell));

const sortedJson = (grid: Grid) => grid.map((r) => JSON.stringify(r)).sort();

/** The practice sheet as setup leaves it, and as the learner's actions then change it. */
interface ModelSheet {
  cells: Cells;
  frozen: { rows: number; cols: number };
  filter: boolean;
  bold: Set<string>;
  hiddenCols: Set<number>;
  /** Column widths in points by column number. A column that isn't here is DEFAULT_WIDTH wide. */
  widths: Map<number, number>;
  tables: { name: string; range: RangeRef }[];
}

/** The sheet right after setup. A plain layout gets no formatting from the coach, and drills write no Tables. */
function practiceSheet(layout: Layout): ModelSheet {
  if (!layout.plain) throw new Error('The model covers plain sheets only');
  return { cells: writtenSheet(layout), frozen: { rows: 0, cols: 0 }, filter: false, bold: new Set(), hiddenCols: new Set(), widths: new Map(), tables: [] };
}

const span = (from: number, to: number) => Array.from({ length: to - from + 1 }, (_, i) => from + i);

/** What the host would read for `check` from the model. */
function factsFor(check: SheetCheck, sheet: ModelSheet): SheetFacts {
  switch (check.kind) {
    case 'freeze':
      return { kind: 'freeze', ...sheet.frozen };
    case 'filter':
      // A Table shows filter buttons too.
      return { kind: 'filter', on: sheet.filter || sheet.tables.length > 0 };
    case 'numberFormat':
      return { kind: 'numberFormat', address: check.range, formats: read(sheet.cells, check.range, 'format') };
    case 'bold':
      return { kind: 'bold', address: check.range, bold: cellsOf(check.range).map((row) => row.map((a) => sheet.bold.has(a))) };
    case 'filled':
      return { kind: 'filled', address: check.range, colors: cellsOf(check.range).map((row) => row.map(() => null)) };
    case 'hidden': {
      const t = hiddenTarget(check);
      if (!t.columns) throw new Error('The model hides columns only');
      const hidden = span(colToNumber(t.first), colToNumber(t.last)).map((c) => sheet.hiddenCols.has(c));
      return { kind: 'hidden', hidden: hidden.every(Boolean) ? true : hidden.some(Boolean) ? null : false };
    }
    case 'minWidth': {
      const r = parseRange(check.range);
      const widths = span(r.start.col, r.end.col).map((c) => (sheet.hiddenCols.has(c) ? 0 : (sheet.widths.get(c) ?? DEFAULT_WIDTH)));
      return { kind: 'minWidth', address: check.range, widths };
    }
    case 'conditionalFormat':
      return { kind: 'conditionalFormat', count: 0 };
    case 'values':
      return { kind: 'values', address: check.range, values: read(sheet.cells, check.range, 'value') };
    case 'formulas':
      return { kind: 'formulas', address: check.range, formulas: read(sheet.cells, check.range, 'formula') };
  }
}

/** Grades every inspection against the model: 'sheet' checks with the real grader, a Table by name, start and size. */
function grade(ex: Exercise<any>, d: unknown, sheet: ModelSheet): { label: string; status: CheckStatus }[] {
  return (ex.inspections?.(d) ?? []).map((insp) => {
    if (insp.kind === 'sheet') return { label: insp.label, status: gradeSheetCheck(insp, factsFor(insp.check, sheet)).status };
    if (insp.kind === 'tableExists') {
      const t = sheet.tables.find((x) => x.name.toLowerCase() === insp.table.toLowerCase());
      const ok = !!t && cellAddress(t.range.start) === insp.at && t.range.end.row - t.range.start.row + 1 === insp.rows;
      return { label: insp.label, status: ok ? 'pass' : 'fail' };
    }
    throw new Error(`The model can’t grade ${insp.kind}`);
  });
}

// ---------- the learner's actions, as Excel carries them out ----------

const isBlank = (cell: ModelCell | undefined) => cell === undefined || cell.value === '' || cell.value === null;

/** The data around A1 up to the first blank row and column: what Excel takes as the list when one cell in it is selected. */
function currentRegion(sheet: ModelSheet): RangeRef {
  let col = 1;
  while (!isBlank(sheet.cells.get(cellAddress({ row: 1, col: col + 1 })))) col++;
  let row = 1;
  while (!isBlank(sheet.cells.get(cellAddress({ row: row + 1, col: 1 })))) row++;
  return { start: { row: 1, col: 1 }, end: { row, col } };
}

/** The column whose header in row 1 reads `header`. */
function headerColumn(sheet: ModelSheet, header: string): number {
  const { end } = currentRegion(sheet);
  const col = span(1, end.col).find((c) => sheet.cells.get(cellAddress({ row: 1, col: c }))?.value === header);
  if (col === undefined) throw new Error(`No ${header} column`);
  return col;
}

/** The last row of the data under a header: where Ctrl+Shift+↓ from row 2 stops. */
function lastDataRow(sheet: ModelSheet, col: number): number {
  let row = 1;
  while (!isBlank(sheet.cells.get(cellAddress({ row: row + 1, col })))) row++;
  return row;
}

/** The region's data rows, below the header, as whole rows of cells. */
function dataRows(sheet: ModelSheet): (ModelCell | undefined)[][] {
  const { end } = currentRegion(sheet);
  return span(2, end.row).map((row) => span(1, end.col).map((col) => sheet.cells.get(cellAddress({ row, col }))));
}

/** Writes rows back under the header. Rows that are gone leave empty cells, as Excel does. */
function putDataRows(sheet: ModelSheet, rows: (ModelCell | undefined)[][], height: number, width: number) {
  for (let r = 0; r < height; r++) {
    for (let c = 0; c < width; c++) {
      const address = cellAddress({ row: r + 2, col: c + 1 });
      const cell = rows[r]?.[c];
      if (cell) sheet.cells.set(address, cell);
      else sheet.cells.delete(address);
    }
  }
}

/** Excel's ascending order for the values drills sort: numbers by size, text A to Z ignoring case. */
function compareValues(a: Cell, b: Cell): number {
  if (typeof a === 'number' && typeof b === 'number') return a - b;
  if (typeof a === 'string' && typeof b === 'string') return excelTextCompare(a, b);
  throw new Error(`The model doesn’t sort ${typeof a} against ${typeof b}`);
}

/** Data › Sort on the list, by each header in turn. Excel's sort is stable. */
function sortList(sheet: ModelSheet, keys: { header: string; descending: boolean }[]) {
  const cols = keys.map((k) => ({ index: headerColumn(sheet, k.header) - 1, sign: k.descending ? -1 : 1 }));
  const rows = dataRows(sheet);
  const sorted = [...rows].sort((a, b) => {
    for (const { index, sign } of cols) {
      const order = compareValues(a[index]?.value ?? '', b[index]?.value ?? '') * sign;
      if (order) return order;
    }
    return 0;
  });
  putDataRows(sheet, sorted, rows.length, rows[0].length);
}

/**
 * Data › Remove Duplicates with every column selected. It compares what each cell shows, ignoring
 * case (in General format a drill value shows as its plain text), keeps the first copy of each
 * row, and moves the rows below up.
 */
function removeDuplicates(sheet: ModelSheet) {
  const rows = dataRows(sheet);
  const seen = new Set<string>();
  const kept = rows.filter((row) => {
    const key = JSON.stringify(row.map((cell) => String(cell?.value ?? '').toLowerCase()));
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
  putDataRows(sheet, kept, rows.length, rows[0].length);
}

/** Moves a formula's relative row references `by` rows, as copying it down does. */
const shiftRows = (formula: string, by: number) => formula.replace(/(\$?[A-Z]{1,3})(\$?)(\d+)/g, (ref, col: string, fixed: string, row: string) => (fixed ? ref : `${col}${Number(row) + by}`));

/**
 * Double-clicking the fill handle of the cell under `header`'s first data row: Excel copies it down
 * as far as the column on its left has data.
 */
function fillDown(sheet: ModelSheet, header: string) {
  const col = headerColumn(sheet, header);
  const top = sheet.cells.get(cellAddress({ row: 2, col }))!;
  for (let row = 3; row <= lastDataRow(sheet, col - 1); row++) {
    const formula = shiftRows(String(top.formula), row - 2);
    sheet.cells.set(cellAddress({ row, col }), { value: evaluate(sheet.cells, formula), formula, format: top.format });
  }
}

/** Copy, then Paste Special › Values over the same cells. */
function pasteValues(sheet: ModelSheet, header: string) {
  const col = headerColumn(sheet, header);
  for (let row = 2; row <= lastDataRow(sheet, col); row++) {
    const address = cellAddress({ row, col });
    const cell = sheet.cells.get(address)!;
    sheet.cells.set(address, { ...cell, formula: cell.value });
  }
}

/** Selects the data under a header and applies a number format. */
function formatColumn(sheet: ModelSheet, header: string, format: string) {
  const col = headerColumn(sheet, header);
  for (let row = 2; row <= lastDataRow(sheet, col); row++) {
    const address = cellAddress({ row, col });
    sheet.cells.set(address, { ...sheet.cells.get(address)!, format });
  }
}

/** Data › Text to Columns, delimited by spaces, on the data under a header. Each part fills the next column to the right. */
function splitAtSpaces(sheet: ModelSheet, header: string) {
  const col = headerColumn(sheet, header);
  for (let row = 2; row <= lastDataRow(sheet, col); row++) {
    String(sheet.cells.get(cellAddress({ row, col }))!.value)
      .split(' ')
      .forEach((part, i) => sheet.cells.set(cellAddress({ row, col: col + i }), { value: part, formula: part, format: 'General' }));
  }
}

/** Find and Replace › Replace All with the default options: the whole sheet, any case, part of a cell counts. */
function replaceAll(sheet: ModelSheet, find: string, replacement: string) {
  const pattern = new RegExp(find.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'gi');
  for (const [address, cell] of sheet.cells) {
    if (typeof cell.value !== 'string') continue;
    const value = cell.value.replace(pattern, replacement);
    if (value !== cell.value) sheet.cells.set(address, { ...cell, value, formula: value });
  }
}

/**
 * The width AutoFit gives a column, in points: its longest entry at about 4.2 points a character,
 * the least a narrow font gives, plus the cell's padding.
 */
function autofitWidth(sheet: ModelSheet, col: number): number {
  const longest = Math.max(...span(1, lastDataRow(sheet, col)).map((row) => String(sheet.cells.get(cellAddress({ row, col }))?.value ?? '').length));
  return longest * 4.2 + 5;
}

/** The `code` the task names, matched by `pattern`: what the learner reads and acts on. */
function fromTask(task: string, pattern: RegExp): string {
  const m = pattern.exec(task);
  if (!m) throw new Error(`The task doesn’t say: ${task}`);
  return m[1];
}

/**
 * What each item's action does to the sheet, from what the task tells the learner. Formats are what
 * each shortcut writes, as Office.js reports them on Windows.
 */
const ACTIONS: Record<string, (sheet: ModelSheet, task: string) => void> = {
  'drill-freeze-top-row': (sheet) => (sheet.frozen = { rows: 1, cols: 0 }),
  'drill-filter-on': (sheet) => (sheet.filter = true),
  'drill-convert-table': (sheet, task) => sheet.tables.push({ name: fromTask(task, /Table named `([^`]+)`/), range: currentRegion(sheet) }),
  'drill-autofit-column': (sheet, task) => {
    const col = colToNumber(fromTask(task, /column `([A-Z]+)`/));
    sheet.widths.set(col, autofitWidth(sheet, col));
  },
  'drill-hide-column': (sheet, task) => sheet.hiddenCols.add(colToNumber(fromTask(task, /column, `([A-Z]+)`/))),
  'drill-currency-format': (sheet) => formatColumn(sheet, 'Amount', '$#,##0.00_);($#,##0.00)'),
  'drill-bold-header': (sheet) => {
    const { end } = currentRegion(sheet);
    for (const col of span(1, end.col)) sheet.bold.add(cellAddress({ row: 1, col }));
  },
  'drill-sort-desc': (sheet) => sortList(sheet, [{ header: 'Amount', descending: true }]),
  'drill-remove-duplicates': removeDuplicates,
  'drill-fill-down': (sheet) => fillDown(sheet, 'Total'),
  'drill-paste-values': (sheet) => pasteValues(sheet, 'Total'),
  'drill-sort-two-keys': (sheet) =>
    sortList(sheet, [
      { header: 'Region', descending: false },
      { header: 'Amount', descending: true },
    ]),
  'drill-split-names': (sheet) => splitAtSpaces(sheet, 'First name'),
  'drill-find-replace': (sheet, task) => replaceAll(sheet, fromTask(task, /Replace every “([^”]+)”/), fromTask(task, /with “([^”]+)”/)),
  'drill-percent-format': (sheet) => formatColumn(sheet, 'Discount', '0%'),
  'drill-date-format': (sheet) => formatColumn(sheet, 'Ship date', 'd-mmm-yy'),
  'drill-time-format': (sheet) => formatColumn(sheet, 'Start', 'h:mm AM/PM'),
  'drill-thousands-format': (sheet) => formatColumn(sheet, 'Units', '#,##0.00'),
  'drill-red-negatives': (sheet) => formatColumn(sheet, 'Variance', '0.00;[Red]0.00'),
  'drill-general-format': (sheet) => formatColumn(sheet, 'Order no.', 'General'),
};

// ---------- registry ----------

describe('drill sets', () => {
  it('offers two or three sets of six to eight actions, each with a reachable par', () => {
    expect(DRILLS.length).toBeGreaterThanOrEqual(2);
    expect(DRILLS.length).toBeLessThanOrEqual(3);
    for (const d of DRILLS) {
      expect(d.items.length, d.id).toBeGreaterThanOrEqual(6);
      expect(d.items.length, d.id).toBeLessThanOrEqual(8);
      // Between 6 and 20 seconds an action: brisk for an intermediate user, not a sprint.
      expect(d.parSeconds / d.items.length, `${d.id} par per action`).toBeGreaterThanOrEqual(6);
      expect(d.parSeconds / d.items.length, `${d.id} par per action`).toBeLessThanOrEqual(20);
    }
  });

  it('uses unique ids, and getDrill finds each set', () => {
    expect(new Set(DRILLS.map((d) => d.id)).size).toBe(DRILLS.length);
    const ids = ITEMS.map(drillItemId);
    expect(new Set(ids).size, 'item ids are unique across every set').toBe(ids.length);
    for (const id of ids) expect(id).toMatch(/^drill-[a-z0-9-]+$/);
    for (const d of DRILLS) expect(getDrill(d.id)).toBe(d);
    expect(getDrill('nope')).toBeUndefined();
  });

  it('covers the sheet setup and sort-and-clean actions', () => {
    expect(getDrill('sheet-setup')!.items.map(drillItemId)).toEqual([
      'drill-freeze-top-row',
      'drill-filter-on',
      'drill-convert-table',
      'drill-autofit-column',
      'drill-hide-column',
      'drill-currency-format',
      'drill-bold-header',
    ]);
    expect(getDrill('sort-clean')!.items.map(drillItemId)).toEqual(
      expect.arrayContaining(['drill-sort-desc', 'drill-remove-duplicates', 'drill-fill-down', 'drill-paste-values', 'drill-sort-two-keys']),
    );
  });

  it('keys per-item bests by item id in progress', () => {
    const drill = DRILLS[0];
    const times = Object.fromEntries(drill.items.map((it, i) => [drillItemId(it), 3000 + i * 100]));
    const first = recordDrillRun(emptyProgress(), drill.id, 30_000, times, 1);
    expect(Object.keys(first.state.drills![drill.id].bestItemMs!)).toEqual(drill.items.map(drillItemId));
    const faster = { ...times, [drillItemId(drill.items[2])]: 1000 };
    expect(recordDrillRun(first.state, drill.id, 29_000, faster, 2).newItemBests).toEqual([drillItemId(drill.items[2])]);
  });
});

// ---------- every item ----------

describe.each(ITEMS.map((entry) => [drillItemId(entry), entry] as [string, DrillItem]))('%s', (_id, entry) => {
  const ex = entry.exercise;

  it('is a one-action drill: plain sheet, graded by inspections alone', () => {
    expect(ex.module).toBe('drills');
    expect(ex.variants).toEqual([]);
    for (const seed of SEEDS) {
      const d = ex.make(new Rng(seed));
      const layout = ex.layout(d);
      expect(layout.plain).toBe(true);
      expect(layout.answer).toEqual({ kind: 'objects' });
      expect(ex.expected(d)).toEqual([]);
      expect(ex.inputs(d)).toEqual([]);
      expect((ex.inspections?.(d) ?? []).length).toBeGreaterThan(0);
      for (const b of layout.blocks) {
        // No Tables and no styled cells: the learner adds all of that.
        if (b.kind === 'data') expect(b.asTable, b.table).toBe(false);
        else expect(['input', 'formula']).toContain(b.role);
      }
    }
  });

  it.each(SEEDS)('seed %i: deterministic, with a valid layout', (seed) => {
    const d = ex.make(new Rng(seed));
    expect(JSON.stringify(d)).toBe(JSON.stringify(ex.make(new Rng(seed))));
    expect(JSON.stringify(ex.layout(d))).toBe(JSON.stringify(ex.layout(ex.make(new Rng(seed)))));
    checkLayout(ex, d, ex.layout(d));
    const ranges = ex.layout(d).blocks.map(blockRange);
    ranges.forEach((a, i) => ranges.slice(i + 1).forEach((b) => expect(overlaps(a, b), 'blocks overlap').toBe(false)));
  });

  it('varies the data from seed to seed', () => {
    expect(new Set(SEEDS.map((seed) => JSON.stringify(ex.make(new Rng(seed))))).size).toBeGreaterThan(SEEDS.length - 2);
  });

  it.each(SEEDS)('seed %i: every checked range lies inside the written data', (seed) => {
    const d = ex.make(new Rng(seed));
    const layout = ex.layout(d);
    const blocks = layout.blocks.map(blockRange);
    const box: RangeRef = {
      start: { row: Math.min(...blocks.map((b) => b.start.row)), col: Math.min(...blocks.map((b) => b.start.col)) },
      end: { row: Math.max(...blocks.map((b) => b.end.row)), col: Math.max(...blocks.map((b) => b.end.col)) },
    };
    const inside = (r: RangeRef) => r.start.row >= box.start.row && r.end.row <= box.end.row && r.start.col >= box.start.col && r.end.col <= box.end.col;
    for (const insp of ex.inspections?.(d) ?? []) {
      if (insp.kind === 'sheet') {
        const c = insp.check;
        if ('range' in c) expect(inside(parseRange(c.range)), `${insp.label}: ${c.range}`).toBe(true);
        if (c.kind === 'hidden' && c.columns) {
          for (const col of c.columns.split(':').map(colToNumber)) expect(col >= box.start.col && col <= box.end.col, c.columns).toBe(true);
        }
      } else if (insp.kind === 'tableExists') {
        const data = layout.blocks.find((b) => b.kind === 'data');
        expect(data && data.kind === 'data' && insp.at === data.at && insp.rows === data.rows.length + 1).toBe(true);
      } else {
        throw new Error(`Unexpected inspection ${insp.kind}`);
      }
    }
  });

  it('fails on the fresh sheet, and passes once the learner does what the task says', () => {
    const act = ACTIONS[drillItemId(entry)];
    expect(act, 'every item has a model of its action').toBeDefined();
    for (const seed of [...SEEDS, ...MANY_SEEDS]) {
      const d = ex.make(new Rng(seed));
      const sheet = practiceSheet(ex.layout(d));
      expect(grade(ex, d, sheet).some((g) => g.status === 'fail'), `seed ${seed}: something fails on the fresh sheet`).toBe(true);
      act(sheet, ex.task(d));
      for (const g of grade(ex, d, sheet)) expect(g.status, `seed ${seed}: ${g.label}, once done`).toBe('pass');
    }
  });

  it('has a shortcut for each platform, in that platform’s notation', () => {
    const { mac, windows } = entry.shortcut;
    expect(mac.trim().length).toBeGreaterThan(1);
    expect(windows.trim().length).toBeGreaterThan(1);
    expect(mac).not.toMatch(/\b(Ctrl|Alt|Shift|Cmd)\b/);
    expect(windows).not.toMatch(/[⌘⌃⌥⇧]/);
    for (const s of [mac, windows]) expect(s, 'menu paths use ›').not.toMatch(/ > /);
  });

  it('follows the copy rules in every string the learner reads', () => {
    for (const seed of [1, 42]) {
      const d = ex.make(new Rng(seed));
      const labels = (ex.inspections?.(d) ?? []).flatMap((i) => [i.label, ...(i.kind === 'sheet' && i.advice ? [i.advice] : [])]);
      const prose = [ex.title, ex.replaces, ex.task(d), ex.concept.summary, ...ex.hints, ...labels];
      for (const text of prose) {
        expect(text, `copy: ${text}`).not.toMatch(BANNED);
        expect(text, `straight apostrophe: ${text}`).not.toMatch(/(?<=[A-Za-z])'(?=[A-Za-z])/);
        expect(text, `menu path: ${text}`).not.toMatch(/ > /);
        expect(text, `Excel spells it AutoFit: ${text}`).not.toMatch(/\b[Aa]uto-?fit/);
      }
      for (const text of [...prose, ex.concept.syntax, ex.solution(d), entry.shortcut.mac, entry.shortcut.windows]) {
        for (const p of PLATFORMS) expect(localize(text, p), `placeholder left in: ${text}`).not.toMatch(/[{}]/);
      }
      // Cell references in the task and hints sit in backticks; the results list doesn't render them.
      for (const text of [ex.task(d), ...ex.hints]) expect(text.replace(/`[^`]*`/g, ''), `bare cell reference: ${text}`).not.toMatch(/\b[A-Z]{1,3}\d+\b/);
      for (const text of labels) expect(text).not.toMatch(/[`{}]/);
      expect(ex.task(d).length).toBeGreaterThan(20);
      expect(ex.hints.length).toBeGreaterThanOrEqual(2);
    }
  });
});

// ---------- answer keys ----------

describe('values the checks expect', () => {
  const valuesCheck = (id: string, d: unknown) => checkOf(item(id).exercise, d, 'values')[0];

  it.each(SEEDS)('seed %i: sort largest first keeps whole rows and runs strictly down by Amount', (seed) => {
    const ex = item('drill-sort-desc').exercise;
    const d = ex.make(new Rng(seed));
    const check = valuesCheck('drill-sort-desc', d);
    const before = read(writtenSheet(ex.layout(d)), check.range, 'value');
    const after = plainGrid(check.expected);
    expect(sortedJson(after)).toEqual(sortedJson(before));
    for (let i = 1; i < after.length; i++) expect(after[i][4] as number).toBeLessThan(after[i - 1][4] as number);
  });

  it.each(SEEDS)('seed %i: the two-key sort orders by Region, then Amount largest first', (seed) => {
    const ex = item('drill-sort-two-keys').exercise;
    const d = ex.make(new Rng(seed));
    const check = valuesCheck('drill-sort-two-keys', d);
    const before = read(writtenSheet(ex.layout(d)), check.range, 'value');
    const after = plainGrid(check.expected);
    expect(sortedJson(after)).toEqual(sortedJson(before));
    for (let i = 1; i < after.length; i++) {
      const order = excelTextCompare(String(after[i - 1][2]), String(after[i][2]));
      expect(order).toBeLessThanOrEqual(0);
      if (order === 0) expect(after[i][4] as number).toBeLessThan(after[i - 1][4] as number);
    }
    // Some region holds two or more rows, so the second key matters.
    expect(new Set(after.map((r) => r[2])).size).toBeLessThan(after.length);
  });

  it('never has tied amounts to sort', () => {
    for (const id of ['drill-sort-desc', 'drill-sort-two-keys']) {
      for (const seed of MANY_SEEDS) {
        const amounts = plainGrid(valuesCheck(id, item(id).exercise.make(new Rng(seed))).expected).map((r) => r[4]);
        expect(new Set(amounts).size).toBe(amounts.length);
      }
    }
  });

  it('works out what Remove Duplicates leaves behind', () => {
    const rows: Grid = [
      ['SO-1', 'Reno', 4],
      ['SO-2', 'Dallas', 7],
      ['so-1', 'RENO', 4],
      ['SO-3', 'Reno', 4],
      ['SO-2', 'Dallas', 7],
    ];
    expect(removeDuplicateRows(rows)).toEqual([
      ['SO-1', 'Reno', 4],
      ['SO-2', 'Dallas', 7],
      ['SO-3', 'Reno', 4],
      ['', '', ''],
      ['', '', ''],
    ]);
    expect(removeDuplicateRows([['a', 1], ['a', '1']])).toEqual([['a', 1], ['a', '1']]);
  });

  it.each(SEEDS)('seed %i: remove duplicates keeps first copies in order, then blanks', (seed) => {
    const ex = item('drill-remove-duplicates').exercise;
    const d = ex.make(new Rng(seed));
    const check = valuesCheck('drill-remove-duplicates', d);
    const before = read(writtenSheet(ex.layout(d)), check.range, 'value');
    const after = plainGrid(check.expected);
    const firsts = before.filter((row, i) => !before.slice(0, i).some((earlier) => JSON.stringify(earlier) === JSON.stringify(row)));
    const repeats = before.length - firsts.length;
    expect(repeats).toBeGreaterThanOrEqual(2);
    expect(repeats).toBeLessThanOrEqual(3);
    expect(after.length).toBe(before.length);
    expect(after.slice(0, firsts.length)).toEqual(firsts);
    expect(after.slice(firsts.length).every((row) => row.every((c) => c === ''))).toBe(true);
    // Order numbers are unique apart from the repeats, so nothing else counts as a duplicate.
    expect(new Set(firsts.map((r) => r[0])).size).toBe(firsts.length);
  });

  it.each(SEEDS)('seed %i: fill down expects each row’s own Qty × Price', (seed) => {
    const ex = item('drill-fill-down').exercise;
    const d = ex.make(new Rng(seed));
    const sheet = writtenSheet(ex.layout(d));
    const check = valuesCheck('drill-fill-down', d);
    const { start, end } = parseRange(check.range);
    expect(start.row).toBe(2);
    expect(sheet.get('E2')!.formula).toBe('=C2*D2');
    expect(sheet.get('E1')!.value).toBe('Total');
    for (let r = start.row; r <= end.row; r++) {
      expect(check.expected[r - 2][0]).toBe((sheet.get(`C${r}`)!.value as number) * (sheet.get(`D${r}`)!.value as number));
      if (r > 2) expect(sheet.has(`E${r}`), `E${r} starts empty`).toBe(false);
    }
    expect(sheet.has(`D${end.row}`) && !sheet.has(`D${end.row + 1}`), 'the totals stop at the last row of data').toBe(true);
  });

  it.each(SEEDS)('seed %i: paste values expects the values the formulas showed', (seed) => {
    const ex = item('drill-paste-values').exercise;
    const d = ex.make(new Rng(seed));
    const sheet = writtenSheet(ex.layout(d));
    const check = valuesCheck('drill-paste-values', d);
    const { start, end } = parseRange(check.range);
    for (let r = start.row; r <= end.row; r++) {
      expect(sheet.get(`E${r}`)!.formula).toBe(`=C${r}*D${r}`);
      expect(check.expected[r - 2][0]).toBe(sheet.get(`E${r}`)!.value);
    }
    expect(checkOf(ex, d, 'formulas')[0]).toMatchObject({ range: check.range, formulas: false });
  });

  it.each(SEEDS)('seed %i: splitting names expects first names in A and last names in B', (seed) => {
    const ex = item('drill-split-names').exercise;
    const d = ex.make(new Rng(seed));
    const sheet = writtenSheet(ex.layout(d));
    const check = valuesCheck('drill-split-names', d);
    const { start, end } = parseRange(check.range);
    expect(check.range.startsWith('A2:B')).toBe(true);
    for (let r = start.row; r <= end.row; r++) {
      const parts = String(sheet.get(`A${r}`)!.value).split(' ');
      expect(parts).toHaveLength(2);
      expect(check.expected[r - 2]).toEqual(parts);
      expect(sheet.get(`B${r}`)!.value, `B${r} starts empty`).toBe('');
    }
  });

  it.each(SEEDS)('seed %i: find and replace changes the old warehouse and nothing else', (seed) => {
    const ex = item('drill-find-replace').exercise;
    const d = ex.make(new Rng(seed));
    const check = valuesCheck('drill-find-replace', d);
    const before = read(writtenSheet(ex.layout(d)), check.range, 'value');
    const hits = before.flat().filter((c) => c === OLD_WAREHOUSE).length;
    expect(hits).toBeGreaterThanOrEqual(2);
    expect(before.flat().filter((c) => c !== OLD_WAREHOUSE && /columbus|cincinnati/i.test(String(c)))).toEqual([]);
    expect(plainGrid(check.expected)).toEqual(before.map((row) => row.map((c) => (c === OLD_WAREHOUSE ? NEW_WAREHOUSE : c))));
  });

  it.each(SEEDS)('seed %i: negative variances are Actual minus Budget, with both signs present', (seed) => {
    const ex = item('drill-red-negatives').exercise;
    const d = ex.make(new Rng(seed));
    const sheet = writtenSheet(ex.layout(d));
    const range = checkOf(ex, d, 'numberFormat')[0].range;
    const values = read(sheet, range, 'value').map((r) => r[0] as number);
    expect(values.filter((v) => v < 0).length).toBeGreaterThanOrEqual(2);
    expect(values.filter((v) => v > 0).length).toBeGreaterThanOrEqual(2);
  });
});

describe('autofit', () => {
  it('always has a customer name long enough to need a wider column', () => {
    const ex = item('drill-autofit-column').exercise;
    for (const seed of MANY_SEEDS) {
      const d = ex.make(new Rng(seed));
      const check = checkOf(ex, d, 'minWidth')[0];
      const names = read(writtenSheet(ex.layout(d)), check.range, 'value').map((r) => String(r[0]));
      const longest = Math.max(...names.map((n) => n.length));
      // About 4.2 points a character in a narrow font, the least AutoFit could give.
      expect(longest * 4.2, `seed ${seed}`).toBeGreaterThan(CUSTOMER_MIN_WIDTH + 25);
    }
    expect(CUSTOMER_MIN_WIDTH).toBeGreaterThan(DEFAULT_WIDTH + 25);
  });

  it('selects the whole column before the Windows keys, which fit only the selected cells', () => {
    // Home › Format › AutoFit Column Width (Alt, H, O, I) with one cell selected fits that cell alone.
    const { windows } = item('drill-autofit-column').shortcut;
    expect(windows).toMatch(/^Ctrl\+Space in column B, then Alt, H, O, I$/);
    expect(item('drill-autofit-column').exercise.task(item('drill-autofit-column').exercise.make(new Rng(1)))).toContain('column `B`');
  });
});

describe('the Table drill', () => {
  it('uses a Table name a learner’s own workbook is unlikely to hold already', () => {
    // Table names are workbook-wide. A clash makes Excel refuse the rename, and the check finds the other Table.
    expect(TABLE_NAME).not.toMatch(/^(Orders?|Sales|Data|Table\d*)$/i);
    const ex = item('drill-convert-table').exercise;
    const d = ex.make(new Rng(1));
    expect(ex.inspections!(d)[0]).toMatchObject({ kind: 'tableExists', table: TABLE_NAME });
    for (const text of [ex.task(d), ex.hints[1], ex.solution(d)]) expect(text).toContain(TABLE_NAME);
  });
});

// ---------- number formats ----------

describe('number format patterns', () => {
  const cases: [string, RegExp, string[], string[]][] = [
    [
      'currency',
      CURRENCY_FORMAT,
      ['$#,##0.00', '$#,##0.00_);($#,##0.00)', '$#,##0.00_);[Red]($#,##0.00)', '[$$-409]#,##0.00', '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)'],
      ['General', '0.00', '#,##0.00', '$#,##0', '0.00%', '[$-409]m/d/yyyy', '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)'],
    ],
    ['percent', PERCENT_FORMAT, ['0%', '0.0%', '0.00%'], ['General', '0.00', '"%"0', '$#,##0.00']],
    ['date', DATE_FORMAT, ['d-mmm-yy', 'm/d/yyyy', 'yyyy-mm-dd', '[$-409]mmmm d, yyyy;@', 'm/d/yy'], ['General', '0.00', 'h:mm AM/PM', '[Red]0.00', '"Day "0', '$#,##0.00']],
    [
      'time',
      TIME_FORMAT,
      ['h:mm AM/PM', '[$-F400]h:mm:ss AM/PM', 'h:mm', 'hh:mm:ss', '[h]:mm', 'mm:ss'],
      ['General', '0.00', 'm/d/yyyy h:mm', 'd-mmm-yy', '0%', '$#,##0.00', '#,##0'],
    ],
    [
      'thousands',
      THOUSANDS_FORMAT,
      ['#,##0.00', '#,##0', '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)', '#,##0;[Red]-#,##0'],
      ['General', '0', '0.00', '$#,##0.00', '0.00%', '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)'],
    ],
    [
      'red negatives',
      RED_NEGATIVE_FORMAT,
      ['0.00;[Red]0.00', '#,##0.00;[Red]#,##0.00', '$#,##0.00_);[Red]($#,##0.00)', '0.00_);[Red](0.00)', '#,##0;[Red]-#,##0'],
      ['General', '[Red]0.00', '0.00;-0.00', '0.00_);(0.00)'],
    ],
    ['general', GENERAL_FORMAT, ['General'], ['m/d/yy', '0', '@', '#,##0']],
  ];

  it.each(cases)('%s accepts what Excel writes and rejects the rest', (_name, pattern, good, bad) => {
    for (const f of good) expect(pattern.test(f), `should accept ${f}`).toBe(true);
    for (const f of bad) expect(pattern.test(f), `should reject ${f}`).toBe(false);
    expect(pattern.global || pattern.sticky, 'no state between tests').toBe(false);
  });

  it('every number-format item starts on a format its check rejects', () => {
    for (const entry of ITEMS) {
      const d = entry.exercise.make(new Rng(7));
      const sheet = writtenSheet(entry.exercise.layout(d));
      for (const c of checkOf(entry.exercise, d, 'numberFormat')) {
        for (const f of read(sheet, c.range, 'format').flat()) expect(c.matches.test(String(f)), `${drillItemId(entry)} starts as ${String(f)}`).toBe(false);
      }
    }
  });
});

// ---------- times ----------

describe('drill times', () => {
  it('reads like a stopwatch, to the tenth', () => {
    expect(formatDrillTime(0)).toBe('0:00.0');
    expect(formatDrillTime(4280)).toBe('0:04.2');
    expect(formatDrillTime(59_999)).toBe('0:59.9');
    expect(formatDrillTime(60_000)).toBe('1:00.0');
    expect(formatDrillTime(754_321)).toBe('12:34.3');
    expect(formatDrillTime(-50)).toBe('0:00.0');
  });

  it('shows gaps in seconds under a minute', () => {
    expect(formatDrillGap(8400)).toBe('8.4 s');
    expect(formatDrillGap(-2650)).toBe('2.6 s');
    expect(formatDrillGap(400)).toBe('0.4 s');
    expect(formatDrillGap(62_300)).toBe('1:02.3');
  });
});

// ---------- the clock and the run ----------

describe('the drill clock', () => {
  it('runs from setup, stops at Done while the check runs, and leaves the check’s time out', () => {
    let clock = startClock(10_000);
    expect(elapsedMs(clock, 14_000)).toBe(4000);
    expect(runningSince(clock)).toBe(10_000);

    clock = pauseClock(clock, 15_000);
    expect(runningSince(clock), 'paused').toBeUndefined();
    expect(elapsedMs(clock, 18_000), 'frozen at Done').toBe(5000);
    expect(pauseClock(clock, 16_000), 'pausing again changes nothing').toBe(clock);

    // The check fails 2.5 s later and the clock picks up where it stopped.
    clock = resumeClock(clock, 17_500);
    expect(elapsedMs(clock, 17_500)).toBe(5000);
    expect(elapsedMs(clock, 20_000)).toBe(7500);
    expect(runningSince(clock)).toBe(12_500);
    expect(resumeClock(clock, 21_000), 'resuming a running clock changes nothing').toBe(clock);
  });

  it('adds up the time of every check that didn’t pass', () => {
    let clock = startClock(0);
    for (const [done, back] of [
      [3000, 3800],
      [6000, 7100],
      [9000, 9300],
    ]) {
      clock = resumeClock(pauseClock(clock, done), back);
    }
    expect(clock.pausedMs).toBe(800 + 1100 + 300);
    expect(elapsedMs(clock, 12_000)).toBe(12_000 - 2200);
  });

  it('records an action’s time at the Done that passes, however long the check takes', () => {
    const clock = resumeClock(pauseClock(startClock(1000), 4000), 4600);
    const passed = pauseClock(clock, 8200);
    expect(elapsedMs(passed, 8200)).toBe(8200 - 1000 - 600);
    expect(elapsedMs(passed, 60_000)).toBe(6600);
  });

  it('reads as now minus runningSince while it runs, which a live display relies on', () => {
    const clock = resumeClock(pauseClock(startClock(500), 2500), 3100);
    for (const now of [3100, 4000, 10_000]) expect(elapsedMs(clock, now)).toBe(now - runningSince(clock)!);
    // runningSince changes whenever the clock starts again, so a display knows to read the time afresh.
    expect(runningSince(clock)).not.toBe(runningSince(startClock(500)));
  });

  it('never reads below zero', () => {
    expect(elapsedMs(startClock(1000), 900)).toBe(0);
    expect(resumeClock({ startedAt: 0, pausedMs: 0, checkAt: 500 }, 400).pausedMs).toBe(0);
  });
});

describe('a drill run', () => {
  const drill = getDrill('sheet-setup')!;
  const n = drill.items.length;
  const report = (...statuses: CheckStatus[]): CheckReport => ({
    passed: statuses.length > 0 && !statuses.includes('fail'),
    items: statuses.map((status, i) => ({ id: `sheet-${i}`, label: `Check ${i + 1}`, status })),
    marks: [],
  });

  it('moves on to the next action after each one, then ends', () => {
    expect(nextAction(drill, [])).toBe(0);
    expect(nextAction(drill, [4200])).toBe(1);
    expect(nextAction(drill, [4200, null])).toBe(2);
    expect(nextAction(drill, Array.from({ length: n - 1 }, () => 3000))).toBe(n - 1);
    expect(nextAction(drill, Array.from({ length: n }, () => 3000))).toBeNull();
  });

  it('totals the checked actions and leaves skipped ones out', () => {
    expect(runTotalMs([])).toBe(0);
    expect(runTotalMs([4200, null, 3100])).toBe(7300);
    expect(runTotalMs([null, null])).toBe(0);
  });

  it('doesn’t count a check that checked nothing as a pass', () => {
    // Without ExcelApi 1.9, the bold check skips both its items, and the host still reports a pass.
    expect(report('skip', 'skip').passed).toBe(true);
    expect(uncheckable(report('skip', 'skip'))).toBe(true);
    expect(uncheckable(report('skip'))).toBe(true);
    expect(uncheckable(report('pass', 'skip'))).toBe(false);
    expect(uncheckable(report('fail', 'skip'))).toBe(false);
    expect(uncheckable(report('pass'))).toBe(false);
    expect(uncheckable(report())).toBe(false);
  });

  it('times a run the way the view does: setup, a failed check, a pass, the next action', () => {
    const times: ItemTime[] = [];
    // First action: ready at 0, a failed check from 4 s to 5 s, then a pass at Done 7 s.
    let clock = startClock(0);
    clock = resumeClock(pauseClock(clock, 4000), 5000);
    times.push(elapsedMs(pauseClock(clock, 7000), 7000));
    expect(nextAction(drill, times)).toBe(1);
    // Second action: its sheet takes 1.5 s to set up after the 600 ms pause, which doesn't count.
    clock = startClock(7000 + 600 + 1500);
    times.push(elapsedMs(pauseClock(clock, 12_100), 12_100));
    expect(times).toEqual([6000, 3000]);
    expect(runTotalMs(times)).toBe(9000);
  });

  it('records a full run’s total and each action’s best', () => {
    const times = drill.items.map((_, i) => 4000 + i * 500);
    const total = runTotalMs(times);
    const first = recordRun(emptyProgress(), drill, times, 1);
    expect(first).toMatchObject({ totalMs: total, skipped: 0, personalBest: true, newItemBests: [], previousBestMs: undefined });
    expect(first.state.drills![drill.id]).toMatchObject({ runs: 1, bestMs: total });
    expect(first.state.drills![drill.id].bestItemMs).toEqual(Object.fromEntries(drill.items.map((it, i) => [drillItemId(it), times[i]])));

    const slower = recordRun(first.state, drill, times.map((t, i) => (i === 3 ? t - 1000 : t + 2000)), 2);
    expect(slower).toMatchObject({ personalBest: false, newItemBests: [drillItemId(drill.items[3])], previousBestMs: total });
    expect(slower.state.drills![drill.id]).toMatchObject({ runs: 2, bestMs: total });
  });

  it('keeps a run with skipped actions from setting the best total, but counts its actions’ bests', () => {
    const full = drill.items.map(() => 9000);
    const before = recordRun(emptyProgress(), drill, full, 1).state;
    // Without ExcelApi 1.9, filter, AutoFit and bold can't be checked.
    const skippedIds = ['drill-filter-on', 'drill-autofit-column', 'drill-bold-header'];
    const partial = drill.items.map((it) => (skippedIds.includes(drillItemId(it)) ? null : 5000));
    const run = recordRun(before, drill, partial, 2);
    expect(run).toMatchObject({ totalMs: 5000 * (n - 3), skipped: 3, personalBest: false, previousBestMs: 9000 * n });
    expect(run.totalMs).toBeLessThan(9000 * n);
    expect(run.state.drills![drill.id]).toMatchObject({ runs: 2, bestMs: 9000 * n, lastRunAt: 2 });
    expect(run.newItemBests.sort()).toEqual(drill.items.map(drillItemId).filter((id) => !skippedIds.includes(id)).sort());
    for (const id of skippedIds) expect(run.state.drills![drill.id].bestItemMs![id], `${id} keeps its best`).toBe(9000);

    // A first run with a skip sets no best total at all.
    const fresh = recordRun(emptyProgress(), drill, partial, 3);
    expect(fresh.personalBest).toBe(false);
    expect(fresh.state.drills![drill.id].bestMs).toBeUndefined();
    expect(fresh.state.drills![drill.id].runs).toBe(1);
  });
});

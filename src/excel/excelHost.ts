import { cellAddress, numberToCol, parseCell, parseRange, rangeAddress, rangeSize, stripSheet } from '../engine/address';
import { gradeBugHunt, type BugHuntRun } from '../engine/bughunt';
import type { RangeRead } from '../engine/fix';
import { isFormula } from '../engine/formula';
import { gradeCellCheck, gradeQueryRows, gradeRules, gradeStructure, gradeValues, gradeVariant, type AnswerRead } from '../engine/grade';
import { Rng } from '../engine/rng';
import type { AnswerArea, Block, Cell, CellMark, CheckItem, CheckReport, Exercise, ExpectedGrid, Grid, InputWrite, Inspection, Layout } from '../engine/types';
import type { SheetFormulas } from '../engine/scan';
import { CoachError } from './errors';
import { FIX_PREFIX, SHEET_PREFIX, fixSheetNameFor, type CoachHost } from './host';
import { inspectSheet, inspectValidation } from './sheetInspect';

const COLOR = {
  answer: '#FFF4CE',
  ok: '#DFF6DD',
  bad: '#FDE7E9',
  header: '#F3F2F1',
  headerRule: '#8A8886',
  tab: '#107C41',
  fixTab: '#C19C00',
};

/** How many rows below a spill anchor get formatted and cleared between checks. */
const SPILL_ROWS = 80;
const MIN_ANSWER_WIDTH = 92;

type Ctx = Excel.RequestContext;

const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms));
const reducedMotion = () => typeof matchMedia !== 'undefined' && matchMedia('(prefers-reduced-motion: reduce)').matches;
const fill2d = <T,>(rows: number, cols: number, value: T): T[][] => Array.from({ length: rows }, () => Array.from({ length: cols }, () => value));

function blockSize(b: Block): { rows: number; cols: number } {
  return b.kind === 'data' ? { rows: b.rows.length + 1, cols: b.columns.length } : { rows: b.values.length, cols: Math.max(...b.values.map((r) => r.length)) };
}

function answerColumns(area: AnswerArea, expected: ExpectedGrid): number[] {
  if (area.kind === 'cells' || area.kind === 'bugHunt') {
    const r = parseRange(area.range);
    return Array.from({ length: r.end.col - r.start.col + 1 }, (_, i) => r.start.col + i);
  }
  if (area.kind === 'spill') {
    const a = parseCell(area.anchor);
    const width = Math.max(1, expected[0]?.length ?? 1, area.formats?.length ?? 0);
    return Array.from({ length: width }, (_, i) => a.col + i);
  }
  if (area.kind === 'cellChecks') return [...new Set(area.checks.map((c) => parseCell(c.cell).col))];
  return [];
}

/** Counts LAMBDA parameters: LAMBDA(a, b, calc) has two. */
export function lambdaParamCount(formula: string): number | null {
  const m = /^=\s*LAMBDA\s*\(/i.exec(formula.replace(/_xlfn\.|_xlpm\./gi, ''));
  if (!m) return null;
  const body = formula.replace(/_xlfn\.|_xlpm\./gi, '').slice(m[0].length);
  let depth = 0;
  let inString = false;
  let args = 1;
  for (const ch of body) {
    if (ch === '"') inString = !inString;
    if (inString) continue;
    if (ch === '(' || ch === '{') depth++;
    else if (ch === ')' || ch === '}') {
      if (depth === 0) break;
      depth--;
    } else if (ch === ',' && depth === 0) args++;
  }
  return args - 1;
}

export class ExcelHost implements CoachHost {
  readonly live = true;

  constructor(private readonly localize: (text: string) => string) {}

  // ---------- setup ----------

  async setup(ex: Exercise<any>, data: unknown, sheetName: string): Promise<void> {
    const layout = ex.layout(data);
    const expected = layout.answer.kind === 'pivot' ? [] : ex.expected(data);

    await Excel.run(async (ctx) => {
      const wb = ctx.workbook;
      const sheets = wb.worksheets;
      sheets.load('items/name');
      const owned = (ex.ownsNames ?? []).map((n) => wb.names.getItemOrNullObject(n));
      owned.forEach((n) => n.load('name'));
      await ctx.sync();

      // Add the new sheet first so the workbook is never left without one.
      const tempName = `${SHEET_PREFIX}new-${Math.floor(Math.random() * 1e6)}`;
      const ws = sheets.add(tempName);
      for (const s of sheets.items) if (s.name.startsWith(SHEET_PREFIX)) s.delete();
      for (const n of owned) if (!n.isNullObject) n.delete();
      await ctx.sync();

      // Table names are workbook-wide; a learner's own table could already use one.
      const tableNames = layout.blocks.filter((b): b is Extract<Block, { kind: 'data' }> => b.kind === 'data' && b.asTable).map((b) => b.table);
      const clashes = tableNames.map((n) => wb.tables.getItemOrNullObject(n));
      clashes.forEach((t) => t.load('name'));
      await ctx.sync();
      const clash = clashes.find((t) => !t.isNullObject);
      if (clash) {
        ws.delete();
        await ctx.sync();
        throw new CoachError(`This workbook already has a Table named ${clash.name}. Open a new blank workbook for practice.`);
      }

      ws.name = sheetName;
      ws.tabColor = COLOR.tab;
      ws.showGridlines = true;
      this.writeLayout(ws, layout, expected);
      await ctx.sync();

      if (!layout.plain) {
        // Size columns to their contents, keeping answer columns comfortably wide.
        ws.getUsedRange().format.autofitColumns();
        // Later mission steps' answer areas too, so a spilled list isn't cut off when it arrives.
        const areas = [layout.answer, ...(layout.alsoStyle ?? [])];
        const colNumbers = new Set(areas.flatMap((a, i) => answerColumns(a, i === 0 ? expected : [])));
        const cols = [...colNumbers].map((c) => ws.getRange(`${numberToCol(c)}:${numberToCol(c)}`));
        cols.forEach((c) => c.format.load('columnWidth'));
        await ctx.sync();
        cols.forEach((c) => {
          if (c.format.columnWidth < MIN_ANSWER_WIDTH) c.format.columnWidth = MIN_ANSWER_WIDTH;
        });
        if (layout.blocks.some((b) => b.kind === 'data' && parseCell(b.at).row === 1)) ws.freezePanes.freezeRows(1);
      }
      await this.refreshStalePivots(ctx, tableNames);
      ws.activate();
      const first = this.firstAnswerCell(layout);
      if (first) ws.getRange(first).select();
      await ctx.sync();
    });
  }

  /**
   * A PivotTable left over from an earlier rep keeps a cache tied to its source Table's name. A new
   * Table with that name inherits the cache, so a PivotTable built from it shows the old fields
   * and rows until refreshed. Seen in Excel for Mac, where Office.js reports those orphaned
   * PivotTables' source as an empty string (type LocalRange). So refresh local-source PivotTables
   * whose source is empty or one of this sheet's Tables; skip Data Model and external sources.
   * Best effort: a refresh that fails (a source that's truly gone) is skipped.
   */
  private async refreshStalePivots(ctx: Ctx, tableNames: string[]) {
    if (!tableNames.length || !Office.context.requirements.isSetSupported('ExcelApi', '1.15')) return;
    const want = new Set(tableNames.map((n) => n.toLowerCase()));
    const pivots = ctx.workbook.pivotTables;
    pivots.load('items/name');
    await ctx.sync();
    for (const pt of pivots.items) {
      try {
        const source = pt.getDataSourceString();
        const type = pt.getDataSourceType();
        await ctx.sync();
        const name = String(source.value).trim().toLowerCase();
        const local = type.value === Excel.DataSourceType.localRange || type.value === Excel.DataSourceType.localTable;
        if (local && (name === '' || want.has(name))) {
          pt.refresh();
          await ctx.sync();
        }
      } catch {
        // Orphaned beyond repair, or an unsupported source; leave it alone.
      }
    }
  }

  private writeLayout(ws: Excel.Worksheet, layout: Layout, expected: ExpectedGrid) {
    for (const b of layout.blocks) {
      const start = parseCell(b.at);
      const { rows, cols } = blockSize(b);
      if (b.kind === 'data') {
        const header = ws.getRange(rangeAddress(start, 1, cols));
        header.values = [b.columns.map((c) => c.header)];
        if (b.rows.length) ws.getRange(rangeAddress({ row: start.row + 1, col: start.col }, b.rows.length, cols)).values = b.rows;
        b.columns.forEach((c, i) => {
          if (c.format && b.rows.length) ws.getRange(rangeAddress({ row: start.row + 1, col: start.col + i }, b.rows.length, 1)).numberFormat = fill2d(b.rows.length, 1, c.format);
        });
        if (b.asTable) {
          const table = ws.tables.add(rangeAddress(start, rows, cols), true);
          table.name = b.table;
          table.style = 'TableStyleMedium2';
        } else if (!layout.plain) {
          this.styleHeader(header);
        }
      } else {
        const range = ws.getRange(rangeAddress(start, rows, cols));
        const values = b.values.map((r) => Array.from({ length: cols }, (_, i) => r[i] ?? ''));
        if (b.role === 'formula') range.formulas = values;
        else range.values = values;
        if (b.format) range.numberFormat = fill2d(rows, cols, b.format);
        b.formats?.forEach((f, i) => {
          if (f) ws.getRange(rangeAddress({ row: start.row, col: start.col + i }, rows, 1)).numberFormat = fill2d(rows, 1, f);
        });
        if (b.role === 'header') this.styleHeader(range);
        if (b.role === 'label') range.format.font.bold = true;
      }
    }

    for (const a of [layout.answer, ...(layout.alsoStyle ?? [])]) this.styleAnswer(ws, a, expected);
  }

  private styleAnswer(ws: Excel.Worksheet, a: AnswerArea, expected: ExpectedGrid) {
    if (a.kind === 'cells') {
      const r = ws.getRange(a.range);
      r.format.fill.color = COLOR.answer;
      const { rows, cols } = rangeSize(a.range);
      if (a.format) r.numberFormat = fill2d(rows, cols, a.format);
    } else if (a.kind === 'spill') {
      const anchor = parseCell(a.anchor);
      const width = Math.max(1, expected[0]?.length ?? 1, a.formats?.length ?? 0);
      for (let i = 0; i < width; i++) {
        const format = a.formats?.[i] ?? a.format;
        if (format) ws.getRange(rangeAddress({ row: anchor.row, col: anchor.col + i }, SPILL_ROWS, 1)).numberFormat = fill2d(SPILL_ROWS, 1, format);
      }
      ws.getRange(a.anchor).format.fill.color = COLOR.answer;
    } else if (a.kind === 'cellChecks') {
      for (const c of a.checks) if (c.answer) ws.getRange(c.cell).format.fill.color = COLOR.answer;
    }
  }

  private styleHeader(range: Excel.Range) {
    range.format.font.bold = true;
    range.format.fill.color = COLOR.header;
    const edge = range.format.borders.getItem('EdgeBottom');
    edge.style = 'Continuous';
    edge.color = COLOR.headerRule;
  }

  private firstAnswerCell(layout: Layout): string | undefined {
    const a = layout.answer;
    if (a.kind === 'cells') return cellAddress(parseRange(a.range).start);
    if (a.kind === 'spill') return a.anchor;
    if (a.kind === 'cellChecks') return a.checks.find((c) => c.answer)?.cell;
    if (a.kind === 'tableColumn') {
      const block = layout.blocks.find((b) => b.kind === 'data' && b.table === a.table);
      if (block && block.kind === 'data') {
        const s = parseCell(block.at);
        return cellAddress({ row: s.row, col: s.col + block.columns.length });
      }
    }
    return undefined;
  }

  // ---------- check ----------

  async check(ex: Exercise<any>, data: unknown, sheetName: string, seed: number): Promise<CheckReport> {
    const layout = ex.layout(data);
    const area = layout.answer;

    return Excel.run(async (ctx) => {
      const ws = ctx.workbook.worksheets.getItemOrNullObject(sheetName);
      ws.load('name');
      await ctx.sync();
      if (ws.isNullObject) throw new CoachError('The practice sheet is gone. Set it up again to continue.', 'missing-sheet');

      const items: CheckItem[] = [];

      for (const [index, inspection] of (ex.inspections?.(data) ?? []).entries()) {
        const result = await this.inspect(ctx, ws, inspection, index);
        items.push(...result.items);
        if (result.blocking) {
          for (const [i, v] of ex.variants.entries()) items.push({ id: `variant-${i}`, label: `Still correct when ${v.label}`, status: 'skip' });
          return this.finish(items, []);
        }
      }
      if (area.kind === 'pivot' || area.kind === 'objects') return this.finish(items, []);
      if (area.kind === 'cellChecks') return this.checkCells(ctx, ws, area, items);
      if (area.kind === 'bugHunt') return this.checkBugHunt(ctx, ws, ex, data, seed, area, items);
      if (area.kind === 'query') {
        items.push(...(await this.checkQuery(ctx, layout, area, ex.expected(data))));
        return this.finish(items, []);
      }

      const expected = ex.expected(data);
      const read = await this.readAnswer(ctx, ws, area, expected);
      if ('missing' in read) {
        items.push(read.missing);
        return this.finish(items, []);
      }

      const rows = area.kind === 'tableColumn' ? read.values.length : expected.length;
      const structure = gradeStructure(area, read, rows, expected[0]?.length ?? 1);
      items.push(...structure.items);
      // With no consistency to keep (each cell its own formula), a "range slides as it fills" hint misleads.
      const values = gradeValues(read, expected, area.kind === 'cells' && area.consistency === 'none' ? [] : structure.formulas);
      items.push(values.item);
      items.push(...gradeRules(ex.rules, structure.formulas));

      const marks: CellMark[] = [...values.marks];
      const formulaOk = structure.items.find((i) => i.id === 'formula')?.status === 'pass';

      if (values.item.status === 'pass' && formulaOk && ex.variants.length) {
        const app = ctx.workbook.application;
        app.load('calculationMode');
        await ctx.sync();
        const manual = app.calculationMode !== 'Automatic';
        const bad = new Set<string>();
        try {
          for (const [i, v] of ex.variants.entries()) {
            const variantData = v.apply(data, new Rng(seed + 7919 * (i + 1)));
            await this.applyInputs(ctx, ws, ex.inputs(variantData));
            if (manual) app.calculate(Excel.CalculationType.full);
            await ctx.sync();
            const variantExpected = ex.expected(variantData);
            const vr = await this.readAnswer(ctx, ws, area, variantExpected);
            if ('missing' in vr) {
              items.push({ id: `variant-${i}`, label: `Still correct when ${v.label}`, status: 'fail', detail: vr.missing.detail });
              continue;
            }
            const g = gradeVariant(i, v.label, v.explain, vr, variantExpected);
            items.push(g.item);
            g.badCells.forEach((c) => bad.add(c));
          }
        } finally {
          await this.applyInputs(ctx, ws, ex.inputs(data));
          if (manual) app.calculate(Excel.CalculationType.full);
          await ctx.sync();
        }
        if (bad.size) {
          const known = new Set(marks.map((m) => m.address));
          for (const m of marks) if (bad.has(m.address)) m.ok = false;
          for (const c of bad) if (!known.has(c)) marks.push({ address: c, ok: false });
        }
      } else {
        for (const [i, v] of ex.variants.entries()) items.push({ id: `variant-${i}`, label: `Still correct when ${v.label}`, status: 'skip' });
      }

      const report = { ...this.finish(items, marks), formulas: structure.formulas };
      await this.paint(ctx, ws, area, expected, read, report);
      return report;
    });
  }

  /**
   * Bug hunts read the range on the current data and under every variant, whether or not the
   * current data passes: a typed-in number only shows up once the data changes.
   */
  private async checkBugHunt(
    ctx: Ctx,
    ws: Excel.Worksheet,
    ex: Exercise<any>,
    data: unknown,
    seed: number,
    area: Extract<AnswerArea, { kind: 'bugHunt' }>,
    items: CheckItem[],
  ): Promise<CheckReport> {
    const readRun = async (expected: ExpectedGrid): Promise<AnswerRead> => {
      const r = await this.readAnswer(ctx, ws, area, expected);
      if ('missing' in r) throw new CoachError('The coach couldn’t read the report range. Set the sheet up again.');
      return r;
    };
    // Grade against the coach's own inputs: a learner may have nudged one while hunting.
    const app = ctx.workbook.application;
    app.load('calculationMode');
    await ctx.sync();
    const manual = app.calculationMode !== 'Automatic';
    await this.applyInputs(ctx, ws, ex.inputs(data));
    if (manual) app.calculate(Excel.CalculationType.full);
    await ctx.sync();

    const expected = ex.expected(data);
    const base = await readRun(expected);
    const runs: BugHuntRun[] = [{ label: '', read: base, expected }];

    if (ex.variants.length) {
      try {
        for (const [i, v] of ex.variants.entries()) {
          const variantData = v.apply(data, new Rng(seed + 7919 * (i + 1)));
          await this.applyInputs(ctx, ws, ex.inputs(variantData));
          if (manual) app.calculate(Excel.CalculationType.full);
          await ctx.sync();
          const variantExpected = ex.expected(variantData);
          runs.push({ label: v.label, explain: v.explain, read: await readRun(variantExpected), expected: variantExpected });
        }
      } finally {
        await this.applyInputs(ctx, ws, ex.inputs(data));
        if (manual) app.calculate(Excel.CalculationType.full);
        await ctx.sync();
      }
    }

    const formulas = [...new Set(base.formulas.flat().filter(isFormula))];
    const graded = gradeBugHunt(area, runs);
    items.push(...graded.items, ...gradeRules(ex.rules, formulas));
    const report = { ...this.finish(items, graded.marks), formulas };
    await this.paint(ctx, ws, area, expected, base, report);
    return report;
  }

  private async checkCells(ctx: Ctx, ws: Excel.Worksheet, area: Extract<AnswerArea, { kind: 'cellChecks' }>, items: CheckItem[]): Promise<CheckReport> {
    const ranges = area.checks.map((c) => {
      const r = ws.getRange(c.cell);
      r.load(['values', 'formulas']);
      return r;
    });
    await ctx.sync();
    const marks: CellMark[] = [];
    area.checks.forEach((c, i) => {
      const item = gradeCellCheck(c, ranges[i].values[0][0] as Cell, ranges[i].formulas[0][0] as Cell);
      items.push(item);
      if (c.answer) marks.push({ address: c.cell, ok: item.status === 'pass' });
    });
    const report = this.finish(items, marks);
    for (const m of marks) ws.getRange(m.address).format.fill.color = m.ok ? COLOR.ok : COLOR.bad;
    if (report.focus) {
      ws.activate();
      ws.getRange(report.focus).select();
    }
    await ctx.sync();
    return report;
  }

  /** Finds the table Power Query loaded, confirms it came from a query, and compares its rows. */
  private async checkQuery(ctx: Ctx, layout: Layout, area: Extract<AnswerArea, { kind: 'query' }>, expected: ExpectedGrid): Promise<CheckItem[]> {
    const items: CheckItem[] = [];
    const sources = new Set(layout.blocks.filter((b) => b.kind === 'data').map((b) => (b as Extract<Block, { kind: 'data' }>).table.toLowerCase()));
    const want = area.columns.map((c) => c.trim().toLowerCase());
    const sourceName = layout.blocks.find((b): b is Extract<Block, { kind: 'data' }> => b.kind === 'data' && b.asTable)?.table ?? 'the source Table';
    const tableLabel = `A table with ${area.columns.join(', ')}`;

    // Query metadata needs ExcelApi 1.14; older builds still get the value checks.
    let queryNames: string[] | null = null;
    try {
      const queries = ctx.workbook.queries;
      queries.load('items/name');
      await ctx.sync();
      queryNames = queries.items.map((q) => q.name);
    } catch {
      queryNames = null;
    }

    const tables = ctx.workbook.tables;
    tables.load('items/name');
    await ctx.sync();
    const candidates = tables.items
      .filter((t) => !sources.has(t.name.toLowerCase()))
      .map((t) => {
        const header = t.getHeaderRowRange();
        header.load('values');
        return { table: t, header };
      });
    await ctx.sync();
    const headerOf = (c: (typeof candidates)[number]) => (c.header.values[0] ?? []).map((v) => String(v).trim().toLowerCase());
    const matches = candidates.filter((c) => {
      const h = headerOf(c);
      return h.length === want.length && want.every((w) => h.includes(w));
    });

    // Excel names a query's output Table after the query, swapping characters a Table name can't
    // hold (spaces, parentheses) and adding _2 when the name is taken: "Export (2)" -> "Export__2".
    const norm = (x: string) => x.toLowerCase().replace(/[^a-z0-9]/g, '');
    const loadedByQuery = (tableName: string) =>
      !queryNames || queryNames.some((q) => norm(q) === norm(tableName) || norm(q) === norm(tableName.replace(/_+\d+$/, '')));

    const rowsOf = async (c: (typeof candidates)[number]): Promise<Grid> => {
      const body = c.table.getDataBodyRange();
      body.load('values');
      await ctx.sync();
      const header = headerOf(c);
      const order = want.map((w) => header.indexOf(w));
      const rows = (body.values as Grid).map((r) => order.map((i) => r[i]));
      // A Table always keeps one row; an all-blank row means the query returned nothing.
      return rows.length === 1 && rows[0].every((cell) => cell === '' || cell === null) ? [] : rows;
    };

    // Earlier attempts can leave older output Tables behind. Grade the one that's right if any is,
    // otherwise the newest one loaded by a query.
    let match: (typeof candidates)[number] | undefined;
    let rowsItem: CheckItem | undefined;
    for (const c of [...matches].reverse()) {
      const item = gradeQueryRows(await rowsOf(c), expected, area.order);
      if (item.status === 'pass' && loadedByQuery(c.table.name)) {
        match = c;
        rowsItem = item;
        break;
      }
      if (!match && loadedByQuery(c.table.name)) {
        match = c;
        rowsItem = item;
      }
    }
    if (!match && matches.length) {
      match = matches[matches.length - 1];
      rowsItem = gradeQueryRows(await rowsOf(match), expected, area.order);
    }

    if (!match || !rowsItem) {
      const near = candidates.find((c) => want.every((w) => headerOf(c).includes(w)));
      const detail = near
        ? `${near.table.name} has extra columns: ${headerOf(near).filter((h) => !want.includes(h)).join(', ')}. Remove them in the Power Query editor.`
        : candidates.length
          ? `No table has exactly these columns yet. Your table ${candidates[0].table.name} has ${headerOf(candidates[0]).join(', ')}.`
          : this.localize(`No output table yet. Choose {fromTable:${sourceName}}, shape the data, then Close & Load.`);
      items.push({ id: 'query-table', label: tableLabel, status: 'fail', detail });
      return items;
    }
    items.push({ id: 'query-table', label: tableLabel, status: 'pass' });

    if (queryNames) {
      items.push(
        loadedByQuery(match.table.name)
          ? { id: 'query-loaded', label: 'Loaded by Power Query', status: 'pass' }
          : {
              id: 'query-loaded',
              label: 'Loaded by Power Query',
              status: 'fail',
              detail: this.localize(`${match.table.name} wasn’t loaded by a query. Build it with {fromTable:${sourceName}}, then Close & Load, so it refreshes next month.`),
            },
      );
    }

    items.push(rowsItem);
    return items;
  }

  private finish(items: CheckItem[], marks: CellMark[]): CheckReport {
    const passed = items.length > 0 && items.every((i) => i.status !== 'fail');
    const focus = items.find((i) => i.status === 'fail' && i.focus)?.focus;
    return { passed, items, marks, focus };
  }

  private async readAnswer(ctx: Ctx, ws: Excel.Worksheet, area: AnswerArea, expected: ExpectedGrid): Promise<AnswerRead | { missing: CheckItem }> {
    let range: Excel.Range;
    if (area.kind === 'cells' || area.kind === 'bugHunt') {
      range = ws.getRange(area.range);
    } else if (area.kind === 'spill') {
      const rows = Math.max(1, expected.length) + 1;
      const cols = Math.max(1, expected[0]?.length ?? 1) + 1;
      range = ws.getRange(rangeAddress(parseCell(area.anchor), rows, cols));
    } else if (area.kind === 'tableColumn') {
      const table = ctx.workbook.tables.getItemOrNullObject(area.table);
      await ctx.sync();
      if (table.isNullObject) return { missing: { id: 'formula', label: `The ${area.table} Table has a ${area.column} column`, status: 'fail', detail: `The ${area.table} Table is missing. Set up the sheet again.` } };
      const column = table.columns.getItemOrNullObject(area.column);
      await ctx.sync();
      if (column.isNullObject) {
        return {
          missing: {
            id: 'formula',
            label: `The ${area.table} Table has a ${area.column} column`,
            status: 'fail',
            detail: `Add a column named “${area.column}” by typing that name in the first empty header cell to the right of the Table.`,
          },
        };
      }
      range = column.getDataBodyRange();
    } else {
      throw new Error(`Answers of kind ${area.kind} are not read as a range`);
    }
    range.load(['address', 'values', 'formulas', 'formulasR1C1']);
    await ctx.sync();
    return {
      address: stripSheet(range.address),
      values: range.values as Grid,
      formulas: range.formulas as Grid,
      r1c1: range.formulasR1C1 as Grid,
    };
  }

  /** Writes input regions. Table writes go column by column, so learner-added columns stay intact. */
  private async applyInputs(ctx: Ctx, ws: Excel.Worksheet, writes: InputWrite[]): Promise<void> {
    for (const w of writes) {
      if (w.kind === 'range') {
        ws.getRange(w.address).values = w.values;
        continue;
      }
      const table = ctx.workbook.tables.getItem(w.table);
      const whole = table.getRange();
      const body = table.getDataBodyRange();
      whole.load(['address', 'columnCount']);
      body.load('rowCount');
      await ctx.sync();

      const current = body.rowCount;
      const wanted = w.rows.length;
      if (wanted > current) {
        const start = parseRange(whole.address).start;
        table.resize(rangeAddress(start, wanted + 1, whole.columnCount));
        await ctx.sync();
      } else if (wanted < current) {
        for (let i = current - 1; i >= wanted; i--) table.rows.getItemAt(i).delete();
        await ctx.sync();
      }
      w.columns.forEach((name, c) => {
        table.columns.getItem(name).getDataBodyRange().values = w.rows.map((r) => [r[c]]);
      });
    }
    await ctx.sync();
  }

  private async paint(ctx: Ctx, ws: Excel.Worksheet, area: AnswerArea, expected: ExpectedGrid, read: AnswerRead, report: CheckReport) {
    // Reset the answer area to its resting color.
    if (area.kind === 'cells') ws.getRange(area.range).format.fill.color = COLOR.answer;
    if (area.kind === 'tableColumn') {
      // The column didn't exist at setup, so its number format is applied once it does.
      const column = ws.getRange(read.address);
      column.format.fill.clear();
      if (area.format) column.numberFormat = fill2d(read.values.length, 1, area.format);
    }
    if (area.kind === 'bugHunt') ws.getRange(area.range).format.fill.clear();
    if (area.kind === 'spill') {
      const a = parseCell(area.anchor);
      const width = Math.max(1, expected[0]?.length ?? 1) + 1;
      ws.getRange(rangeAddress(a, SPILL_ROWS, width)).format.fill.clear();
      ws.getRange(area.anchor).format.fill.color = COLOR.answer;
    }
    await ctx.sync();

    const ok = report.marks.filter((m) => m.ok).map((m) => m.address);
    const bad = report.marks.filter((m) => !m.ok).map((m) => m.address);
    const paintCells = (cells: string[], color: string) => {
      for (let i = 0; i < cells.length; i += 30) ws.getRanges(cells.slice(i, i + 30).join(',')).format.fill.color = color;
    };

    if (report.passed && ok.length > 1 && !reducedMotion()) {
      // A short sweep: the green arrives in a few beats instead of all at once.
      const steps = Math.min(5, ok.length);
      const per = Math.ceil(ok.length / steps);
      for (let s = 0; s < steps; s++) {
        paintCells(ok.slice(s * per, (s + 1) * per), COLOR.ok);
        await ctx.sync();
        await sleep(55);
      }
    } else {
      paintCells(ok, COLOR.ok);
      paintCells(bad, COLOR.bad);
      await ctx.sync();
    }

    if (report.focus) {
      ws.activate();
      ws.getRange(report.focus).select();
      await ctx.sync();
    }
  }

  // ---------- structural inspections ----------

  private async inspect(ctx: Ctx, ws: Excel.Worksheet, insp: Inspection, index: number): Promise<{ items: CheckItem[]; blocking: boolean }> {
    const wb = ctx.workbook;
    switch (insp.kind) {
      case 'tableExists': {
        const table = wb.tables.getItemOrNullObject(insp.table);
        const onSheet = ws.tables;
        onSheet.load('items/name');
        await ctx.sync();
        if (table.isNullObject) {
          const other = onSheet.items[0];
          const detail = other
            ? `Your Table is named ${other.name}. Rename it to ${insp.table} in the Table tab (Table Design on Windows).`
            : this.localize(`There’s no Table yet. Click in the data, press {tableKey}, then name it ${insp.table}.`);
          return { items: [{ id: 'table', label: insp.label, status: 'fail', detail }], blocking: true };
        }
        const range = table.getRange();
        range.load(['address', 'rowCount']);
        table.worksheet.load('name');
        await ctx.sync();
        const start = cellAddress(parseRange(range.address).start);
        if (table.worksheet.name !== ws.name || start !== insp.at) {
          return { items: [{ id: 'table', label: insp.label, status: 'fail', detail: `The ${insp.table} Table should start at ${insp.at} on the practice sheet.` }], blocking: true };
        }
        if (range.rowCount !== insp.rows) {
          return {
            items: [{ id: 'table', label: insp.label, status: 'fail', detail: `The Table covers ${stripSheet(range.address)}, but the data runs to row ${insp.rows}. Resize it to include every row.` }],
            blocking: true,
          };
        }
        return { items: [{ id: 'table', label: insp.label, status: 'pass' }], blocking: false };
      }

      case 'tableColumn': {
        const table = wb.tables.getItemOrNullObject(insp.table);
        await ctx.sync();
        if (table.isNullObject) return { items: [{ id: 'column', label: insp.label, status: 'fail', detail: `The ${insp.table} Table is missing. Set the sheet up again.` }], blocking: true };
        const column = table.columns.getItemOrNullObject(insp.column);
        await ctx.sync();
        return column.isNullObject
          ? {
              items: [{ id: 'column', label: insp.label, status: 'fail', detail: `Type “${insp.column}” in the first empty header cell right of the Table; it grows to include the new column.` }],
              blocking: true,
            }
          : { items: [{ id: 'column', label: insp.label, status: 'pass' }], blocking: false };
      }

      case 'lambdaName': {
        let named = wb.names.getItemOrNullObject(insp.name);
        named.load(['formula']);
        await ctx.sync();
        if (named.isNullObject) {
          named = ws.names.getItemOrNullObject(insp.name);
          named.load(['formula']);
          await ctx.sync();
        }
        if (named.isNullObject) {
          return { items: [{ id: 'lambda', label: insp.label, status: 'fail', detail: this.localize(`There’s no name called ${insp.name} yet. Create it in {nameManager}.`) }], blocking: true };
        }
        const params = lambdaParamCount(String(named.formula));
        if (params === null) {
          return { items: [{ id: 'lambda', label: insp.label, status: 'fail', detail: `${insp.name} exists, but it refers to ${named.formula}. It should start with =LAMBDA(.` }], blocking: true };
        }
        if (params !== insp.params) {
          return { items: [{ id: 'lambda', label: insp.label, status: 'fail', detail: `${insp.name} takes ${params} input${params === 1 ? '' : 's'}; it should take ${insp.params}.` }], blocking: true };
        }
        return { items: [{ id: 'lambda', label: insp.label, status: 'pass' }], blocking: false };
      }

      case 'pivot':
        return this.inspectPivot(ctx, insp);

      case 'chart':
        return this.inspectChart(ctx, ws, insp);

      case 'slicer': {
        let slicers: Excel.SlicerCollection;
        try {
          slicers = ctx.workbook.slicers;
          slicers.load('items/name,items/caption');
          await ctx.sync();
        } catch {
          return { items: [{ id: 'slicer', label: insp.label, status: 'skip', detail: 'This version of Excel can’t report slicers to add-ins.' }], blocking: false };
        }
        const want = insp.field.toLowerCase();
        const hit = slicers.items.find((s) => s.caption.toLowerCase().includes(want) || s.name.toLowerCase().includes(want));
        return {
          items: [
            hit
              ? { id: 'slicer', label: insp.label, status: 'pass' }
              : {
                  id: 'slicer',
                  label: insp.label,
                  status: 'fail',
                  detail: slicers.items.length
                    ? `There’s a slicer for ${slicers.items[0].caption}, but none for ${insp.field}.`
                    : `There’s no slicer yet. Click in the Table or PivotTable, then Insert Slicer and tick ${insp.field}.`,
                },
          ],
          blocking: false,
        };
      }

      case 'sheet':
        return { items: [await inspectSheet(ctx, ws, insp, index)], blocking: false };

      case 'validationList':
        return { items: [await inspectValidation(ctx, ws, insp, index)], blocking: false };

      default:
        return { items: [], blocking: false };
    }
  }

  private async inspectChart(ctx: Ctx, ws: Excel.Worksheet, insp: Extract<Inspection, { kind: 'chart' }>): Promise<{ items: CheckItem[]; blocking: boolean }> {
    const charts = ws.charts;
    charts.load('items/name,items/chartType');
    await ctx.sync();
    if (charts.items.length === 0) {
      return { items: [{ id: 'chart', label: insp.label, status: 'fail', detail: 'There’s no chart on the practice sheet yet.' }], blocking: true };
    }
    const loaded = charts.items.map((c) => {
      c.series.load('items/name,items/chartType,items/axisGroup');
      c.title.load('text,visible');
      return c;
    });
    await ctx.sync();

    const assess = (c: Excel.Chart) => {
      const series = c.series.items;
      const types = [String(c.chartType), ...series.map((s) => String(s.chartType))];
      return {
        c,
        enough: insp.minSeries === undefined || series.length >= insp.minSeries,
        typeOk: !insp.types || types.some((t) => insp.types!.test(t)),
        comboOk: !insp.secondaryLine || series.some((s) => /^Line/i.test(String(s.chartType)) && String(s.axisGroup) === 'Secondary'),
        titleOk: !insp.title || (c.title.visible && !!c.title.text && c.title.text.trim() !== 'Chart Title'),
      };
    };
    const scored = loaded.map(assess);
    const best = scored.reduce((a, b) => (+b.enough + +b.typeOk + +b.comboOk + +b.titleOk > +a.enough + +a.typeOk + +a.comboOk + +a.titleOk ? b : a));
    const items: CheckItem[] = [{ id: 'chart', label: insp.label, status: 'pass' }];
    if (insp.minSeries !== undefined)
      items.push(
        best.enough
          ? { id: 'chart-series', label: `Shows at least ${insp.minSeries} series`, status: 'pass' }
          : { id: 'chart-series', label: `Shows at least ${insp.minSeries} series`, status: 'fail', detail: `The chart has ${best.c.series.items.length}. Select all the columns you need before inserting it, or use Select Data.` },
      );
    if (insp.types)
      items.push(
        best.typeOk
          ? { id: 'chart-type', label: 'Right chart type', status: 'pass' }
          : { id: 'chart-type', label: 'Right chart type', status: 'fail', detail: `This is a ${String(best.c.chartType)} chart. Change Chart Type to match the task.` },
      );
    if (insp.secondaryLine)
      items.push(
        best.comboOk
          ? { id: 'chart-combo', label: 'One series is a line on the secondary axis', status: 'pass' }
          : {
              id: 'chart-combo',
              label: 'One series is a line on the secondary axis',
              status: 'fail',
              detail: 'Change Chart Type › Combo, set the percentage series to Line, and tick Secondary Axis for it.',
            },
      );
    if (insp.title)
      items.push(
        best.titleOk
          ? { id: 'chart-title', label: 'Has a real title', status: 'pass' }
          : { id: 'chart-title', label: 'Has a real title', status: 'fail', detail: 'Give the chart a title that says what it shows, not “Chart Title”.' },
      );
    return { items, blocking: false };
  }

  private async inspectPivot(ctx: Ctx, insp: Extract<Inspection, { kind: 'pivot' }>): Promise<{ items: CheckItem[]; blocking: boolean }> {
    const pivots = ctx.workbook.pivotTables;
    pivots.load('items/name');
    await ctx.sync();
    if (pivots.items.length === 0) {
      return { items: [{ id: 'pivot', label: 'A PivotTable exists', status: 'fail', detail: 'There’s no PivotTable yet. Click inside the data Table on the practice sheet, then Insert › PivotTable.' }], blocking: true };
    }

    const loaded = pivots.items.map((pt) => {
      pt.rowHierarchies.load('items/name');
      pt.columnHierarchies.load('items/name');
      pt.dataHierarchies.load('items/name,items/showAs,items/summarizeBy');
      pt.filterHierarchies.load('items/name');
      return pt;
    });
    await ctx.sync();

    // Excel lists its built-in "Values" placeholder (Σ Values) among row or column fields.
    const realField = (name: string) => !/^(Σ\s*)?Values$/i.test(name.trim());
    const describe = (pt: Excel.PivotTable) => {
      const rows = pt.rowHierarchies.items.map((h) => h.name).filter(realField);
      const cols = pt.columnHierarchies.items.map((h) => h.name).filter(realField);
      const filters = pt.filterHierarchies.items.map((h) => h.name);
      const data = pt.dataHierarchies.items.find((h) => new RegExp(insp.valuesField, 'i').test(h.name)) ?? pt.dataHierarchies.items[0];
      return { rows, cols, filters, data };
    };
    const wantCols = insp.columns ? [insp.columns] : [];
    const score = (pt: Excel.PivotTable) => {
      const d = describe(pt);
      return (d.rows.includes(insp.rows) ? 1 : 0) + (wantCols.every((c) => d.cols.includes(c)) ? 1 : 0) + (d.data ? 1 : 0);
    };
    const best = loaded.reduce((a, b) => (score(b) > score(a) ? b : a));
    const { rows, cols, filters, data } = describe(best);
    const list = (xs: string[]) => (xs.length ? xs.join(', ') : 'nothing');
    const items: CheckItem[] = [];

    items.push(
      rows.length === 1 && rows[0] === insp.rows
        ? { id: 'pivot-rows', label: `Rows: ${insp.rows}`, status: 'pass' }
        : { id: 'pivot-rows', label: `Rows: ${insp.rows}`, status: 'fail', detail: `Rows has ${list(rows)}. Drag ${insp.rows} into Rows, and remove anything else.` },
    );
    if (insp.columns) {
      items.push(
        cols.length === 1 && cols[0] === insp.columns
          ? { id: 'pivot-cols', label: `Columns: ${insp.columns}`, status: 'pass' }
          : { id: 'pivot-cols', label: `Columns: ${insp.columns}`, status: 'fail', detail: `Columns has ${list(cols)}. Drag ${insp.columns} into Columns.` },
      );
    } else if (cols.length) {
      items.push({ id: 'pivot-cols', label: 'Nothing in Columns', status: 'fail', detail: `Columns has ${list(cols)}. Drag it out of Columns.` });
    }
    if (insp.filter) {
      items.push(
        filters.includes(insp.filter)
          ? { id: 'pivot-filter', label: `Filters: ${insp.filter}`, status: 'pass' }
          : { id: 'pivot-filter', label: `Filters: ${insp.filter}`, status: 'fail', detail: `Filters has ${list(filters)}. Drag ${insp.filter} into Filters.` },
      );
    }

    const valueOk = !!data && new RegExp(insp.valuesField, 'i').test(data.name) && String(data.summarizeBy) === insp.summarizeBy;
    items.push(
      valueOk
        ? { id: 'pivot-values', label: `Values: ${insp.summarizeBy} of ${insp.valuesField}`, status: 'pass' }
        : {
            id: 'pivot-values',
            label: `Values: ${insp.summarizeBy} of ${insp.valuesField}`,
            status: 'fail',
            detail: data ? `Values shows ${data.name}. Use ${insp.summarizeBy} of ${insp.valuesField}.` : `Drag ${insp.valuesField} into Values.`,
          },
    );

    const calc = String(data?.showAs?.calculation ?? 'None');
    const showAsLabel: Record<string, string> = {
      PercentOfRowTotal: '% of Row Total',
      PercentOfColumnTotal: '% of Column Total',
      PercentOfGrandTotal: '% of Grand Total',
      None: 'No Calculation',
    };
    if (insp.showAs !== 'None') {
      items.push(
        calc === insp.showAs
          ? { id: 'pivot-showas', label: `Show Values As: ${showAsLabel[insp.showAs]}`, status: 'pass' }
          : {
              id: 'pivot-showas',
              label: `Show Values As: ${showAsLabel[insp.showAs]}`,
              status: 'fail',
              detail: `The values show ${showAsLabel[calc] ?? calc}. Right-click a value › Show Values As › ${showAsLabel[insp.showAs]}.`,
            },
      );
    }
    return { items, blocking: false };
  }

  // ---------- small helpers ----------

  async sheetExists(sheet: string): Promise<boolean> {
    return Excel.run(async (ctx) => {
      const ws = ctx.workbook.worksheets.getItemOrNullObject(sheet);
      await ctx.sync();
      return !ws.isNullObject;
    });
  }

  async readWorkbook(): Promise<{ name: string; sheets: SheetFormulas[]; truncated: boolean }> {
    const MAX_CELLS = 60_000;
    return Excel.run(async (ctx) => {
      ctx.workbook.load('name');
      const sheets = ctx.workbook.worksheets;
      sheets.load('items/name');
      await ctx.sync();
      const used = sheets.items
        .filter((ws) => !ws.name.startsWith(SHEET_PREFIX) && !ws.name.startsWith(FIX_PREFIX))
        .map((ws) => {
          const range = ws.getUsedRangeOrNullObject();
          range.load(['address', 'rowCount', 'columnCount']);
          return { ws, range };
        });
      await ctx.sync();

      let truncated = false;
      const reads = used
        .filter((u) => !u.range.isNullObject)
        .map((u) => {
          // Large sheets: read the top rows only, so the scan stays fast.
          const maxRows = Math.max(1, Math.floor(MAX_CELLS / Math.max(1, u.range.columnCount)));
          const target = u.range.rowCount > maxRows ? u.range.getResizedRange(maxRows - u.range.rowCount, 0) : u.range;
          if (u.range.rowCount > maxRows) truncated = true;
          target.load(['address', 'formulas', 'formulasR1C1']);
          return { sheet: u.ws.name, target };
        });
      await ctx.sync();
      return {
        name: ctx.workbook.name,
        truncated,
        sheets: reads.map((r) => ({
          sheet: r.sheet,
          address: stripSheet(r.target.address),
          formulas: r.target.formulas as Grid,
          r1c1: r.target.formulasR1C1 as Grid,
        })),
      };
    });
  }

  async writeInputs(sheet: string, writes: InputWrite[]): Promise<void> {
    if (!writes.length) return;
    await Excel.run(async (ctx) => {
      const ws = ctx.workbook.worksheets.getItemOrNullObject(sheet);
      await ctx.sync();
      if (ws.isNullObject) throw new CoachError('The practice sheet is gone. Set it up again to continue.', 'missing-sheet');
      await this.applyInputs(ctx, ws, writes);
    });
  }

  async copySheetForFix(sheet: string): Promise<string> {
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.10')) {
      throw new CoachError('This version of Excel can’t copy sheets for add-ins. Update Excel to use guided fixes.');
    }
    const name = fixSheetNameFor(sheet);
    return Excel.run(async (ctx) => {
      const source = ctx.workbook.worksheets.getItemOrNullObject(sheet);
      const earlier = ctx.workbook.worksheets.getItemOrNullObject(name);
      await ctx.sync();
      if (source.isNullObject) throw new CoachError(`The sheet ${sheet} isn’t in this workbook anymore. Scan again.`, 'missing-sheet');
      if (!earlier.isNullObject) {
        earlier.delete();
        await ctx.sync();
      }
      const copy = source.copy(Excel.WorksheetPositionType.after, source);
      copy.name = name;
      copy.tabColor = COLOR.fixTab;
      copy.activate();
      await ctx.sync();
      return name;
    });
  }

  async readRange(sheet: string, address: string): Promise<RangeRead> {
    return Excel.run(async (ctx) => {
      const ws = ctx.workbook.worksheets.getItemOrNullObject(sheet);
      await ctx.sync();
      if (ws.isNullObject) throw new CoachError(`The sheet ${sheet} isn’t in this workbook anymore.`, 'missing-sheet');
      const range = ws.getRange(address);
      range.load(['address', 'values', 'formulas']);
      await ctx.sync();
      return { address: stripSheet(range.address), values: range.values as Grid, formulas: range.formulas as Grid };
    });
  }

  async select(sheet: string, address: string): Promise<void> {
    await Excel.run(async (ctx) => {
      const ws = ctx.workbook.worksheets.getItem(sheet);
      ws.activate();
      ws.getRange(address).select();
      await ctx.sync();
    });
  }
}

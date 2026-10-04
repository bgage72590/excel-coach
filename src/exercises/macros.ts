import { ITEMS, VENDORS, WAREHOUSES, serial, skuCode, sum } from '../engine/data';
import { round, type Rng } from '../engine/rng';
import type { CellMatcher, Exercise, Grid, Inspection, Layout, SheetCheck } from '../engine/types';
import { cells, defineExercise } from './common';

/**
 * The add-in can't read VBA, so these exercises grade what the macro leaves behind with 'sheet'
 * inspections. Every rep has a different number of rows, so a macro recorded with fixed ranges
 * passes once and breaks on the next rep. Sheets are plain (no autofit, freeze or header styling)
 * and the data has no number formats, so all formatting on the sheet is the learner's.
 */

export const MIN_ROWS = 12;
export const MAX_ROWS = 40;

/**
 * Currency and accounting formats as Office.js reports them, with or without a symbol:
 * "$#,##0.00", "$#,##0.00_);[Red]($#,##0.00)", "[$$-409]#,##0.00", "#,##0.00" (Currency with
 * no symbol), _("$"* #,##0.00_);… (Accounting) and _(* #,##0.00_);… (Comma Style). A "$" right
 * after "[" is a locale tag on a date format ([$-409]m/d/yyyy), not a symbol. Plain "0.00" and
 * "#,##0", percentages and dates don't count.
 */
export const MONEY_FORMAT = /^(?!.*%)(?=.*[0#])(?=.*(?:(?<!\[)\$|[€£¥]|\*|#,##0\.00))/;

/**
 * Date formats: a d or y code outside [brackets], "quotes" and \escapes. Accepts m/d/yyyy,
 * yyyy-mm-dd, d-mmm-yy and [$-409]mmmm d, yyyy;@. Rejects General, numbers, currency and
 * time-only formats like h:mm.
 */
export const DATE_FORMAT = /^(?:\[[^\]]*\]|"[^"]*"|\\.|[^"[\\dy])*[dy]/i;

/**
 * Description must be at least this wide (points) after autofit. A default column is roughly 48
 * to 65 points wide, depending on the workbook's font; every description is 45+ characters, which
 * autofits to about 190 points or more even in a narrow font.
 */
export const MIN_DESC_WIDTH = 140;

/** A 'sheet' inspection. */
const sheetCheck = (check: SheetCheck, label: string, advice?: string): Inspection => ({ kind: 'sheet', check, label, advice });

// ---------- the weekly order report ----------

export const WEEKLY_HEADERS = ['Date', 'Order', 'Customer', 'Description', 'Units', 'Sales', 'Cost'] as const;

const CUSTOMERS = [
  'Northgate Hardware',
  'Lakeside Builders',
  'Pinecrest Facilities',
  'Summit Property Group',
  'Riverbend Schools',
  'Harbor Logistics',
  'Maple Street Market',
  'Cedar Ridge Clinic',
] as const;

/** Delivery notes, long enough that autofit has to widen Description well past its default. */
const NOTES = [
  'deliver to the north dock before noon Friday',
  'split shipment, second half follows next week',
  'customer picks up at the will-call counter',
  'rush order, ship by overnight freight today',
  'replaces units damaged in the previous delivery',
  'standing order, same quantity every week',
  'hold for credit approval before it ships',
  'call the receiving desk an hour before arrival',
] as const;

export interface WeeklyRow {
  date: number;
  order: string;
  customer: string;
  description: string;
  units: number;
  sales: number;
  cost: number;
}

export interface WeeklyData {
  rows: WeeklyRow[];
}

/** One week of orders, Monday to Friday, in date order. */
function weeklyReport(rng: Rng): WeeklyData {
  const count = rng.int(MIN_ROWS, MAX_ROWS);
  const monday = serial(2026, 1, 5) + 7 * rng.int(0, 47);
  const firstOrder = rng.int(40, 90) * 100;
  const dates = Array.from({ length: count }, () => monday + rng.int(0, 4)).sort((a, b) => a - b);
  const rows = dates.map((date, i) => {
    const it = rng.pick(ITEMS);
    // Cheap items sell by the case, expensive ones a few at a time.
    const units = it.cost < 1 ? rng.int(4, 40) * 25 : rng.int(2, 48);
    const price = round(it.cost * rng.float(1.25, 1.6, 3), 2);
    return {
      date,
      order: `SO-${firstOrder + i}`,
      customer: rng.pick(CUSTOMERS),
      description: `${it.item}, ${rng.pick(NOTES)}`,
      units,
      sales: round(units * price, 2),
      cost: round(units * it.cost, 2),
    };
  });
  return { rows };
}

const weeklyGrid = (rows: WeeklyRow[]): Grid => rows.map((r) => [r.date, r.order, r.customer, r.description, r.units, r.sales, r.cost]);

/** Header and data as plain input cells: no number formats, no header styling. */
const weeklyLayout = (d: WeeklyData): Layout => ({
  blocks: [cells('A1', [[...WEEKLY_HEADERS]], 'input'), cells('A2', weeklyGrid(d.rows), 'input')],
  answer: { kind: 'objects' },
  plain: true,
});

/** The sheet row of the last order (the header is row 1). */
export const lastDataRow = (d: WeeklyData) => d.rows.length + 1;

/** The row the Total row belongs in: right under the last order. */
export const totalRow = (d: WeeklyData) => d.rows.length + 2;

/** What =SUM(F2:F…) and =SUM(G2:G…) show on the Total row. */
export const weeklyTotals = (d: WeeklyData): [number, number] => [round(sum(d.rows.map((r) => r.sales)), 2), round(sum(d.rows.map((r) => r.cost)), 2)];

export const TOTAL_LABEL: CellMatcher = { match: /^\s*total\s*$/i, describe: '“Total”' };

// ---------- Record a macro that formats a report ----------

const FORMAT_VBA = [
  "' Recorded, then tidied. Stored in the Personal Macro Workbook,",
  "' it formats whichever sheet is active.",
  'Sub FormatWeekly()',
  "    ' Header row: bold with a light blue fill",
  '    With Range("A1").CurrentRegion.Rows(1)',
  '        .Font.Bold = True',
  '        .Interior.Color = RGB(221, 235, 247)',
  '    End With',
  '',
  "    ' Whole columns, so the formats reach every row",
  '    Columns("A").NumberFormat = "yyyy-mm-dd"',
  '    Columns("F:G").NumberFormat = "$#,##0.00"',
  '',
  "    ' Widths last, so they fit the formatted values",
  '    Range("A1").CurrentRegion.Columns.AutoFit',
  '',
  "    ' Freeze row 1 only, counted from the top of the sheet",
  '    ActiveWindow.FreezePanes = False',
  '    ActiveWindow.ScrollRow = 1',
  '    ActiveWindow.ScrollColumn = 1',
  '    ActiveWindow.SplitColumn = 0',
  '    ActiveWindow.SplitRow = 1',
  '    ActiveWindow.FreezePanes = True',
  '',
  "    ' AutoFilter switches the buttons on and off, so only call it when they are off",
  '    If Not ActiveSheet.AutoFilterMode Then',
  '        Range("A1").CurrentRegion.AutoFilter',
  '    End If',
  'End Sub',
].join('\n');

export const macroRecordFormat = defineExercise<WeeklyData>({
  id: 'macro-record-format',
  module: 'macros',
  title: 'Record a macro that formats a report',
  replaces: 'Formatting the same weekly export by hand every Monday',
  minutes: 8,
  task: (d) => {
    const last = lastDataRow(d);
    return `This weekly order report gets the same formatting every week. Record a macro named \`FormatWeekly\` that makes the header row bold with a fill color, shows Date as dates and Sales and Cost as currency, autofits the columns, freezes the top row and turns on filter buttons. Store it in your Personal Macro Workbook. This report has ${d.rows.length} rows and the next one won’t, so on the next rep, run FormatWeekly instead of formatting by hand. Graded: \`A1:G1\` bold and filled, dates in \`A2:A${last}\`, currency in \`F2:G${last}\`, column \`D\` wide enough to read, the top row frozen and filters on.`;
  },
  concept: {
    summary:
      'The macro recorder turns each step you take into VBA code, so a routine you repeat every week becomes one command. It replays exactly what you selected. Click whole column letters and row numbers rather than dragging over cells: Columns("F:G") covers every row, while Range("F2:G28") stops at this week’s last row.',
    syntax: '{recordMacro} (or Developer › Record Macro)\nMacro name: FormatWeekly\nStore macro in: Personal Macro Workbook\nDo each step once, then Stop Recording.',
    example:
      'Click the F and G column letters and choose Currency, and the recorder writes Columns("F:G").Select, then Selection.NumberFormat = "$#,##0.00". That formats every row, however long next week’s report is.',
    tip: 'The Personal Macro Workbook is a hidden workbook that opens with Excel, so a macro stored there runs on any workbook, including this practice one. Choose Save when Excel asks about it as you quit. Stored in This Workbook instead, a macro is kept only if you save the file as an Excel Macro-Enabled Workbook (.xlsm).',
  },
  hints: [
    'Start the recorder with {recordMacro}. Name the macro `FormatWeekly`, set Store macro in to Personal Macro Workbook, then select OK.',
    'Select whole rows and columns by clicking their headings, not by dragging over cells. Click row `1` for the header, column `A` for the dates and columns `F:G` for the money, so each step covers any number of rows.',
    'Format first, then autofit: select columns `A:G` and double-click the border between two column letters, so the widths fit the formatted values. Then click `A1`, freeze the top row from the View tab, choose Data › Filter and stop recording.',
    'On the next rep, run it with {runMacro}, select FormatWeekly, then Run. For quicker access, turn on the Developer tab in {developerTab}.',
  ],
  solution: () => FORMAT_VBA,
  make: weeklyReport,
  layout: weeklyLayout,
  expected: () => [],
  inputs: () => [],
  variants: [],
  inspections: (d) => {
    const last = lastDataRow(d);
    return [
      sheetCheck({ kind: 'bold', range: 'A1:G1', bold: true }, 'Header row is bold', 'Bold the header while recording; clicking the row 1 heading selects all of it.'),
      sheetCheck({ kind: 'filled', range: 'A1:G1', filled: true }, 'Header row has a fill color', 'Pick any fill color except white.'),
      sheetCheck(
        { kind: 'numberFormat', range: `A2:A${last}`, matches: DATE_FORMAT, describe: 'formatted as dates' },
        `Dates in A2:A${last} use a date format`,
        `Unformatted, dates show as serial numbers, like ${d.rows[0].date}. Format all of column A so every row is covered.`,
      ),
      sheetCheck(
        { kind: 'numberFormat', range: `F2:G${last}`, matches: MONEY_FORMAT, describe: 'formatted as currency or accounting' },
        `Sales and Cost in F2:G${last} use a currency format`,
        'Select the whole F and G columns before you format them. A format recorded on a fixed range misses the extra rows in a longer report.',
      ),
      sheetCheck(
        { kind: 'minWidth', range: `D1:D${last}`, points: MIN_DESC_WIDTH },
        'Description column is wide enough to read',
        'In the macro, autofit columns A to G after the number formats.',
      ),
      sheetCheck({ kind: 'freeze', rows: 1, cols: 0 }, 'Top row is frozen', 'Freeze Top Row on the View tab freezes row 1 alone in one step.'),
      sheetCheck({ kind: 'filter', on: true }, 'Filter buttons are on', 'AutoFilter switches the buttons on and off, so running a macro twice can turn them off again.'),
    ];
  },
});

// ---------- Make a macro work on any number of rows ----------

const ANY_ROWS_VBA = [
  'Sub FormatWeekly()',
  '    Dim lastRow As Long, totalRow As Long',
  '',
  "    ' The last row with data in column A, however long the report is",
  '    lastRow = Cells(Rows.Count, 1).End(xlUp).Row',
  "    ' Stop if this sheet already has its Total row",
  '    If Cells(lastRow, 1).Value = "Total" Then Exit Sub',
  '    totalRow = lastRow + 1',
  '',
  '    With Range("A1").CurrentRegion.Rows(1)',
  '        .Font.Bold = True',
  '        .Interior.Color = RGB(221, 235, 247)',
  '    End With',
  '    Columns("A").NumberFormat = "yyyy-mm-dd"',
  '    Columns("F:G").NumberFormat = "$#,##0.00"',
  '',
  '    ActiveWindow.FreezePanes = False',
  '    ActiveWindow.ScrollRow = 1',
  '    ActiveWindow.ScrollColumn = 1',
  '    ActiveWindow.SplitColumn = 0',
  '    ActiveWindow.SplitRow = 1',
  '    ActiveWindow.FreezePanes = True',
  '    If Not ActiveSheet.AutoFilterMode Then',
  '        Range("A1").CurrentRegion.AutoFilter',
  '    End If',
  '',
  "    ' Total row: R2C is row 2 of the same column, R[-1]C the row above",
  '    Cells(totalRow, 1).Value = "Total"',
  '    Range(Cells(totalRow, 6), Cells(totalRow, 7)).FormulaR1C1 = "=SUM(R2C:R[-1]C)"',
  '    Range(Cells(totalRow, 1), Cells(totalRow, 7)).Font.Bold = True',
  '',
  "    ' Widths last, so they fit the formatted values and the totals",
  '    Range("A1").CurrentRegion.Columns.AutoFit',
  'End Sub',
].join('\n');

export const macroAnyRows = defineExercise<WeeklyData>({
  id: 'macro-any-rows',
  module: 'macros',
  title: 'Make a macro work on any number of rows',
  replaces: 'Fixing the Total row by hand whenever the report has more or fewer rows',
  minutes: 8,
  task: (d) => {
    const t = totalRow(d);
    return `Edit \`FormatWeekly\` so it also adds a Total row right under the last data row: “Total” in column \`A\`, SUM formulas for Sales and Cost in columns \`F\` and \`G\`, and the row bold from \`A\` to \`G\`. This report has ${d.rows.length} rows and the next one won’t, so find the last row in code, with Cells(Rows.Count, 1).End(xlUp).Row or CurrentRegion, never a fixed address. Graded: “Total” in \`A${t}\`, SUM formulas in \`F${t}:G${t}\` that add up every data row, and \`A${t}:G${t}\` bold.`;
  },
  concept: {
    summary:
      'The recorder saves exact addresses, so a macro recorded on a 30-row report writes its Total row to row 32 every time. Edit the code to find the edge of the data each time it runs. Cells(Rows.Count, 1).End(xlUp).Row starts at the bottom of column A and jumps up to the last filled cell. Range("A1").CurrentRegion is the whole block of data around A1.',
    syntax: 'lastRow = Cells(Rows.Count, 1).End(xlUp).Row\nCells(lastRow + 1, 1).Value = "Total"\nCells(lastRow + 1, 6).Formula = "=SUM(F2:F" & lastRow & ")"',
    example: 'With 30 data rows, lastRow is 31, so “Total” goes in A32 and the formula becomes =SUM(F2:F31). Next week the same lines fit 18 rows or 40.',
    tip: 'Use Relative References, on the Developer tab, makes the recorder write moves as offsets from the active cell, such as ActiveCell.Offset(1, 0), instead of fixed cells. A recorded AutoSum still adds a fixed number of rows above it, like =SUM(R[-30]C:R[-1]C), so replace it with a range that starts at row 2.',
  },
  hints: [
    'Open {vbaEditor}. FormatWeekly is in {personalProject}, under Modules. Double-click the module to see the code.',
    'At the top of the macro, find the last data row with lastRow = Cells(Rows.Count, 1).End(xlUp).Row. It works like pressing {jumpUp} from the bottom of column `A`.',
    'The Total row is lastRow + 1. Write “Total” there in column `A`, then build each SUM from lastRow: Cells(lastRow + 1, 6).Formula = "=SUM(F2:F" & lastRow & ")", and the same for column `G`, which is column 7.',
    'Bold it with Range(Cells(lastRow + 1, 1), Cells(lastRow + 1, 7)).Font.Bold = True. Run the macro with {runMacro} here, then again on the next rep: the Total row should land right under the data both times.',
  ],
  solution: () => ANY_ROWS_VBA,
  make: weeklyReport,
  layout: weeklyLayout,
  expected: () => [],
  inputs: () => [],
  variants: [],
  inspections: (d) => {
    const t = totalRow(d);
    return [
      sheetCheck(
        { kind: 'values', range: `A${t}`, expected: [[TOTAL_LABEL]], describe: 'the Total label' },
        `“Total” in A${t}, right under the last data row`,
        'The Total row goes right under the last data row. Find it with Cells(Rows.Count, 1).End(xlUp).Row + 1, not a fixed address.',
      ),
      sheetCheck(
        { kind: 'values', range: `F${t}:G${t}`, expected: [weeklyTotals(d)], describe: 'the totals' },
        `Sales and Cost totals in F${t}:G${t}`,
        'Each SUM should run from row 2 to the last data row. A recorded AutoSum adds a fixed number of rows, so build the range from the last row instead.',
      ),
      sheetCheck(
        { kind: 'formulas', range: `F${t}:G${t}`, formulas: true, pattern: /SUM\(/i },
        `F${t}:G${t} hold SUM formulas`,
        'Write SUM formulas with .Formula or .FormulaR1C1, so the totals update when a row changes.',
      ),
      sheetCheck({ kind: 'bold', range: `A${t}:G${t}`, bold: true }, `Total row is bold, A${t}:G${t}`, 'Bold the whole row from A to G, not only the cells with values.'),
    ];
  },
});

// ---------- Highlight reorder rows with a loop ----------

export const STOCK_HEADERS = ['SKU', 'Item', 'Warehouse', 'On hand', 'Reorder point', 'Supplier'] as const;

export interface StockLine {
  sku: string;
  item: string;
  warehouse: string;
  onHand: number;
  reorder: number;
  supplier: string;
}

export interface ReorderData {
  rows: StockLine[];
}

/**
 * A stock list with 3–6 rows below their reorder point and exactly one row sitting on its reorder
 * point, which isn't below it. The first and last data rows are always low: a skipped row that
 * isn't low looks the same as one the loop checked, so only a low row at each end makes a loop
 * that starts at row 3 or stops one row short fail on every rep.
 */
function stockList(rng: Rng): ReorderData {
  const count = rng.int(MIN_ROWS, MAX_ROWS);
  const low = new Set([0, count - 1]);
  const lowCount = rng.int(3, 6);
  while (low.size < lowCount) low.add(rng.int(1, count - 2));
  const boundary = rng.pick(Array.from({ length: count }, (_, i) => i).filter((i) => !low.has(i)));
  const numbers = rng.sample(Array.from({ length: 900 }, (_, i) => 100 + i), count).sort((a, b) => a - b);
  const rows = numbers.map((n, i) => {
    const reorder = rng.int(4, 40) * 5;
    const onHand = low.has(i) ? rng.int(0, reorder - 1) : i === boundary ? reorder : reorder + rng.int(1, 2 * reorder);
    return { sku: skuCode(n), item: rng.pick(ITEMS).item, warehouse: rng.pick(WAREHOUSES), onHand, reorder, supplier: rng.pick(VENDORS) };
  });
  return { rows };
}

const stockGrid = (rows: StockLine[]): Grid => rows.map((r) => [r.sku, r.item, r.warehouse, r.onHand, r.reorder, r.supplier]);

/** Sheet rows (the header is row 1) where On hand is below Reorder point. */
export const lowRows = (d: ReorderData): number[] => d.rows.flatMap((r, i) => (r.onHand < r.reorder ? [i + 2] : []));

/**
 * The rows checked for no fill: the one exactly at its reorder point (the < versus <= slip),
 * plus the first and last of the other rows that aren't low.
 */
export function okSample(d: ReorderData): number[] {
  const ok = d.rows.flatMap((r, i) => (r.onHand >= r.reorder ? [i + 2] : []));
  const atPoint = ok.filter((row) => d.rows[row - 2].onHand === d.rows[row - 2].reorder);
  const others = ok.filter((row) => !atPoint.includes(row));
  return [...new Set([...atPoint.slice(0, 1), others[0], others[others.length - 1]])].sort((a, b) => a - b);
}

const HIGHLIGHT_VBA = [
  'Sub HighlightReorders()',
  '    Dim lastRow As Long, lastCol As Long, r As Long',
  '',
  '    lastRow = Cells(Rows.Count, 1).End(xlUp).Row',
  '    lastCol = Cells(1, Columns.Count).End(xlToLeft).Column',
  '',
  "    ' Clear old fills, so restocked rows lose their color",
  '    Range(Cells(2, 1), Cells(lastRow, lastCol)).Interior.ColorIndex = xlNone',
  '',
  "    ' Column D is On hand, column E is Reorder point",
  '    For r = 2 To lastRow',
  '        If Cells(r, 4).Value < Cells(r, 5).Value Then',
  '            Range(Cells(r, 1), Cells(r, lastCol)).Interior.Color = RGB(255, 199, 206)',
  '        End If',
  '    Next r',
  'End Sub',
].join('\n');

export const macroHighlightRows = defineExercise<ReorderData>({
  id: 'macro-highlight-rows',
  module: 'macros',
  title: 'Highlight reorder rows with a loop',
  replaces: 'Scanning the stock list and coloring reorder rows one at a time',
  minutes: 10,
  task: (d) =>
    `Write a macro named \`HighlightReorders\` that loops over the data rows and fills each row where On hand (column \`D\`) is below Reorder point (column \`E\`), across columns \`A\` to \`F\`. Leave every other row unfilled; a row exactly at its reorder point isn’t below it. Set the fill with Interior.Color, not conditional formatting. This list has ${d.rows.length} rows and the next one won’t, so find the last row in code. Graded: each low row filled across \`A\` to \`F\`; the header row and a sample of the other rows, including the one at its reorder point, left unfilled.`,
  concept: {
    summary:
      'A For loop runs the same lines once for each row, so a macro can decide row by row: fill this one, skip that one. Find the last row first, then loop from row 2 down to it. Conditional formatting isn’t the point here: it changes what you see, not the cell’s own fill, and the coach checks the fill. Set it with Interior.Color.',
    syntax:
      'For r = 2 To lastRow\n    If Cells(r, 4).Value < Cells(r, 5).Value Then\n        Range(Cells(r, 1), Cells(r, 6)).Interior.Color = RGB(255, 199, 206)\n    End If\nNext r',
    example: 'In row 7, Cells(7, 4) is D7 (On hand) and Cells(7, 5) is E7 (Reorder point). With 4 on hand and a reorder point of 12, the macro fills A7:F7.',
    tip: 'Clear old fills before the loop with Range("A2:F" & lastRow).Interior.ColorIndex = xlNone, so a row that’s been restocked loses its color the next time the macro runs.',
  },
  hints: [
    'Open {vbaEditor}, select {personalProject}, then Insert › Module. If it’s not listed, that workbook doesn’t exist yet: record any short macro into it first, or use this workbook’s project. Type Sub HighlightReorders() and press {enter}; the editor adds End Sub.',
    'Find the last row with lastRow = Cells(Rows.Count, 1).End(xlUp).Row, then loop with For r = 2 To lastRow and close the loop with Next r.',
    'Inside the loop, compare Cells(r, 4).Value with Cells(r, 5).Value. When On hand is smaller, set Range(Cells(r, 1), Cells(r, 6)).Interior.Color. Use <, not <=, so a row at its reorder point stays unfilled.',
    'Run it with {runMacro}. If an earlier version filled the wrong rows, clear those fills first with Home › Fill Color › No Fill, or clear every data row in the macro before the loop, as in the tip.',
  ],
  solution: () => HIGHLIGHT_VBA,
  make: stockList,
  layout: (d) => ({
    blocks: [cells('A1', [[...STOCK_HEADERS]], 'input'), cells('A2', stockGrid(d.rows), 'input')],
    answer: { kind: 'objects' },
    plain: true,
  }),
  expected: () => [],
  inputs: () => [],
  variants: [],
  inspections: (d) => {
    // Compared as text, "On hand" < "Reorder point" is True, so a loop from row 1 fills the header.
    const header = sheetCheck(
      { kind: 'filled', range: 'A1:F1', filled: false },
      'Header row has no fill',
      'Start the loop at row 2. Compared as text, “On hand” is less than “Reorder point”, so a loop from row 1 fills the headers too. Clear row 1’s fill with Home › Fill Color › No Fill.',
    );
    const low = lowRows(d).map((row) => {
      const r = d.rows[row - 2];
      return {
        row,
        inspection: sheetCheck(
          { kind: 'filled', range: `A${row}:F${row}`, filled: true },
          `Row ${row} is filled (${r.onHand} on hand, reorder point ${r.reorder})`,
          `On hand is below the reorder point here. Fill A${row}:F${row} with Interior.Color; conditional formatting doesn’t change a cell’s own fill.`,
        ),
      };
    });
    const ok = okSample(d).map((row) => {
      const r = d.rows[row - 2];
      const atPoint = r.onHand === r.reorder;
      return {
        row,
        inspection: sheetCheck(
          { kind: 'filled', range: `A${row}:F${row}`, filled: false },
          atPoint ? `Row ${row} has no fill (exactly at its reorder point of ${r.reorder})` : `Row ${row} has no fill (${r.onHand} on hand, reorder point ${r.reorder})`,
          atPoint
            ? 'On hand equals the reorder point, which isn’t below it. Compare with <, not <=.'
            : 'This row isn’t below its reorder point. Check that the loop compares column D with column E, and clear old fills before the loop.',
        ),
      };
    });
    return [header, ...[...low, ...ok].sort((a, b) => a.row - b.row).map((x) => x.inspection)];
  },
});

export const MACROS: Exercise<any>[] = [macroRecordFormat, macroAnyRows, macroHighlightRows];

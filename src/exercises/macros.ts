import { ITEMS, VENDORS, WAREHOUSES, serial, skuCode, sum } from '../engine/data';
import { round, type Rng } from '../engine/rng';
import type { CellMatcher, Exercise, GuideStep, Grid, Inspection, Layout, SheetCheck, StepDone } from '../engine/types';
import { cells, defineExercise } from './common';
import { checkStep, isoDate, money, part, raw } from './guides';

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

/**
 * A walkthrough step that ticks off when one of the exercise's own sheet inspections passes: the one
 * of `kind`, on `range` when two share a kind.
 */
function whenPasses(list: Inspection[], kind: SheetCheck['kind'], range?: string): StepDone {
  const hit = list.find((i) => i.kind === 'sheet' && i.check.kind === kind && (range === undefined || ('range' in i.check && i.check.range === range)));
  if (!hit) throw new Error(`No ${kind} inspection${range ? ` on ${range}` : ''}`);
  return { kind: 'inspect', inspection: hit };
}

/** Whole data rows `A` to `F` for the given sheet rows, runs merged: [2, 3, 9] gives "A2:F3,A9:F9". */
function rowBands(rows: number[]): string {
  const sorted = [...new Set(rows)].sort((a, b) => a - b);
  const out: string[] = [];
  for (let i = 0; i < sorted.length; ) {
    let j = i;
    while (j + 1 < sorted.length && sorted[j + 1] === sorted[j] + 1) j++;
    out.push(`A${sorted[i]}:F${sorted[j]}`);
    i = j + 1;
  }
  return out.join(',');
}

/** Where a recorded or saved macro lives in the editor, for the walkthroughs. */
const OPEN_MODULE = 'Open {vbaEditor}. In the project list on the left, open {personalProject}, then **Modules**, and double-click the module inside it.';

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
  "    ' AutoFilter switches the buttons on and off, so only call it when they are off.",
  "    ' Field:=1 with no criteria shows every row; Excel for Mac needs a field here.",
  '    If Not ActiveSheet.AutoFilterMode Then',
  '        Range("A1").CurrentRegion.AutoFilter Field:=1',
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
    'On the next rep, run it with {runMacro}, select FormatWeekly, then Run. For quicker access, turn on the Developer tab in {developerTab}. If Excel stops at an AutoFilter line with error 1004, open {vbaEditor} and add Field:=1 to the end of that line.',
  ],
  solution: () => FORMAT_VBA,
  guide: (d): GuideStep[] => {
    const last = lastDataRow(d);
    const checks: Inspection[] = macroRecordFormat.inspections!(d);
    const firstDate = d.rows[0].date;
    return [
      {
        do: `Meet the report: ${d.rows.length} orders from one week, with no formatting yet. You’ll format it once with the macro recorder running, so next week it’s one command.`,
        why: 'Next week’s report won’t have the same number of rows. So every step below selects whole rows and columns by clicking their headings (the letters along the top and the numbers down the left edge). A step recorded on a whole column covers any number of rows.',
        show: [
          { label: 'Header row', at: 'A1:G1', note: 'Row `1`. It’ll be bold with a fill color.' },
          { label: 'Dates', at: `A2:A${last}`, note: `Dates show as plain numbers, like ${firstDate}, until they get a date format.` },
          { label: 'Sales and Cost', at: `F2:G${last}`, note: 'Plain numbers for now. They’ll show as currency.' },
          { label: 'Description', at: `D1:D${last}`, note: 'Column `D` is too narrow to read. AutoFit will widen it.' },
        ],
      },
      {
        do: 'Start the recorder: choose **{recordMacro}**.',
        why: 'The **Record Macro** box opens. From the moment you click OK until you stop it, Excel writes each thing you do as a line of VBA code, so take the steps one at a time. If you’ve turned on the Developer tab ({developerTab}), **Developer › Record Macro** opens the same box.',
      },
      {
        do: 'In **Macro name**, type `FormatWeekly`.',
        why: 'A macro name can’t contain spaces. Leave **Shortcut key** empty.',
      },
      {
        do: 'Set **Store macro in** to **Personal Macro Workbook**, then click **OK**.',
        why: 'The Personal Macro Workbook is a hidden workbook that opens with Excel, so a macro stored there runs on any workbook, including next week’s report. When you quit Excel and it asks about saving it, choose **Save**. The recorder is now running.',
      },
      {
        do: 'Click the row `1` heading (the 1 at the left edge of the sheet), then click **Bold** on the **Home** tab.',
        why: 'Clicking the heading selects the whole row, so the step covers the header however many columns it has.',
        done: whenPasses(checks, 'bold'),
      },
      {
        do: 'With row `1` still selected, click the arrow beside **Fill Color** (the paint bucket) on the **Home** tab and pick a light color.',
        why: 'Any color but white counts. A light one keeps the header text readable.',
        done: whenPasses(checks, 'filled'),
      },
      {
        do: 'Click the column `A` heading. In the number format box on the **Home** tab (it shows **General**), choose **Short Date**.',
        why: `Excel stores a date as a serial number: ${firstDate} is ${isoDate(firstDate)}. A date format changes only how it looks. Formatting all of column \`A\` covers every row, however long next week’s report is.`,
        done: whenPasses(checks, 'numberFormat', `A2:A${last}`),
      },
      {
        do: 'Click the column `F` heading, then Shift-click the `G` heading so both columns are selected. In the same number format box, choose **Currency**.',
        why: 'The recorder writes `Columns("F:G").Select`, then sets the format on the selection. That reaches every row, however long next week’s report is.',
        done: whenPasses(checks, 'numberFormat', `F2:G${last}`),
      },
      {
        do: 'Click the column `A` heading, then Shift-click the `G` heading. Double-click the border between any two of the selected column letters, for example between `D` and `E`.',
        why: 'That’s AutoFit: every selected column widens or narrows to fit its longest entry. Doing it after the number formats means the widths fit the dates and dollar amounts as they now look.',
        done: whenPasses(checks, 'minWidth'),
      },
      {
        do: 'Click `A1`. Then on the **View** tab, choose **Freeze Top Row** (it may sit inside the **Freeze Panes** menu).',
        why: 'Freeze Top Row freezes row `1` alone in one step, so the headings stay in view as you scroll down the orders. Clicking `A1` first drops the column selection, so the filter in the next step covers the whole report.',
        done: whenPasses(checks, 'freeze'),
      },
      {
        do: 'Choose **Data › Filter**.',
        why: 'Filter buttons appear on each heading. **Filter** switches them on and off, so choose it once.',
        done: whenPasses(checks, 'filter'),
      },
      {
        do: 'Stop the recorder. The menu you started it from now offers **Stop Recording**; the small square **Stop** button in the status bar, at the bottom of the Excel window, does the same.',
        why: 'FormatWeekly is saved. The formatting you did while recording stays on this sheet, and that’s what the coach checks.',
      },
      {
        do: OPEN_MODULE,
        why: 'That’s where the recorder put FormatWeekly. Yours is longer than the tidied version above: the recorder writes a `.Select` line before most steps and lists settings you didn’t change. Each tidied line matches something you did.',
        formula: [
          part('Sub FormatWeekly()\n', 'The macro starts here, under the name you gave it.'),
          part('    With Range("A1").CurrentRegion.Rows(1)\n', 'The header row: the first row of the block of data around `A1`. The lines up to `End With` apply to it.', 'A1:G1'),
          part('        .Font.Bold = True\n', 'Bold, as you clicked.'),
          part('        .Interior.Color = RGB(221, 235, 247)\n', 'The fill color, as amounts of red, green and blue. Yours holds the color you picked.'),
          raw('    End With\n'),
          part('    Columns("A").NumberFormat = "yyyy-mm-dd"\n', 'A date format on all of column `A`, however many rows there are.', `A2:A${last}`),
          part('    Columns("F:G").NumberFormat = "$#,##0.00"\n', 'Currency on all of columns `F` and `G`.', `F2:G${last}`),
          part('    Range("A1").CurrentRegion.Columns.AutoFit\n', 'AutoFit for every column of the report, after the formats.', `A1:G${last}`),
          part('    ActiveWindow.FreezePanes = False\n', 'This line and the five under it freeze row `1` alone: clear any old freeze, scroll to the top, split below row 1, then freeze. Yours may look different and do the same job.'),
          raw('    ActiveWindow.ScrollRow = 1\n    ActiveWindow.ScrollColumn = 1\n    ActiveWindow.SplitColumn = 0\n    ActiveWindow.SplitRow = 1\n    ActiveWindow.FreezePanes = True\n'),
          part('    If Not ActiveSheet.AutoFilterMode Then\n', 'Only when the filter buttons are off. AutoFilter switches them on and off, so a second run would turn them off again.'),
          part('        Range("A1").CurrentRegion.AutoFilter Field:=1\n', 'Turns the filter buttons on. `Field:=1` with no criteria (nothing to filter by) shows every row. Excel for Mac needs a field here.', 'A1:G1'),
          raw('    End If\n'),
          part('End Sub', 'The macro ends here.'),
        ],
      },
      {
        do: 'In your code, find the line that ends with the word `AutoFilter`. Click at its end and type a space, then `Field:=1`.',
        why: 'Run on Excel for Mac, a bare `AutoFilter` line stops the macro with error 1004. With `Field:=1` and no criteria it shows every row, on Mac and Windows alike. If the line already has something after `AutoFilter`, leave it. Then switch back to the workbook.',
      },
      checkStep(
        'The coach reads the formatting on this sheet. On the next rep the report has a different number of rows: run FormatWeekly with **{runMacro}**, select **FormatWeekly**, then click **Run**, instead of formatting by hand.',
      ),
    ];
  },
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
  '        Range("A1").CurrentRegion.AutoFilter Field:=1',
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
  guide: (d): GuideStep[] => {
    const last = lastDataRow(d);
    const t = totalRow(d);
    const [sales, cost] = weeklyTotals(d);
    const checks: Inspection[] = macroAnyRows.inspections!(d);
    return [
      {
        do: `Meet the report: ${d.rows.length} orders, so the last one is on row ${last} and the Total row belongs in row ${t}.`,
        why: `A recorded macro would write its Total row to row ${t} every time. Next week’s report has a different number of rows, so the macro has to find the last row each time it runs.`,
        show: [
          { label: 'Last order', at: `A${last}:G${last}`, note: `Row ${last}: the last row with data.` },
          { label: 'Where the Total row goes', at: `A${t}:G${t}`, note: `Row ${t} is empty now. The macro will write the Total row here.` },
        ],
      },
      {
        do: OPEN_MODULE,
        why: 'You’re adding to FormatWeekly, the macro from the Record a macro skill, so one command formats the report and adds its Total row. No FormatWeekly yet? Select this workbook’s project, choose **Insert › Module**, type `Sub FormatWeekly()` and press {enter}; the editor adds `End Sub` under it.',
      },
      {
        do: 'Click at the end of the `Sub FormatWeekly()` line, press {enter}, and type these four lines.',
        formula: [
          part('    Dim lastRow As Long, totalRow As Long\n', 'Makes two variables: named boxes that hold a number while the macro runs. `Long` is a whole number big enough for any row.'),
          part(
            '    lastRow = Cells(Rows.Count, 1).End(xlUp).Row\n',
            `Finds the last order. It starts at the very bottom of column \`A\` (\`Rows.Count\` is the sheet’s last row) and jumps up to the last filled cell, as {jumpUp} does. On this report that’s row ${last}.`,
            `A${last}`,
          ),
          part('    If Cells(lastRow, 1).Value = "Total" Then Exit Sub\n', 'Stops if the report already has its Total row, so running the macro twice doesn’t add a second one or switch the filter buttons back off.'),
          part('    totalRow = lastRow + 1', `The row right under the last order: row ${t} here, and wherever it lands next week.`, `A${t}:G${t}`),
        ],
        why: 'Nothing on the sheet changes yet. These lines only work out where the Total row goes.',
      },
      {
        do: 'Click at the end of the last line above `End Sub`, press {enter}, and type these three lines.',
        formula: [
          part('    Cells(totalRow, 1).Value = "Total"\n', '`Cells(row, column)` counts columns as numbers, so column 1 is `A`. This writes “Total” there.', `A${t}`),
          part(
            '    Cells(totalRow, 6).Formula = "=SUM(F2:F" & lastRow & ")"\n',
            `Writes a SUM formula in column 6, \`F\`. \`&\` joins the text to the number in lastRow, so on this report it writes \`=SUM(F2:F${last})\`.`,
            `F2:F${last}`,
          ),
          part('    Cells(totalRow, 7).Formula = "=SUM(G2:G" & lastRow & ")"', `The same for column 7, \`G\`: \`=SUM(G2:G${last})\` here.`, `G2:G${last}`),
        ],
        why: 'Building each range from lastRow is what makes it fit any report. A recorded AutoSum adds a fixed number of rows above it instead.',
      },
      {
        do: 'Under those, type these two lines.',
        formula: [
          part('    Range(Cells(totalRow, 1), Cells(totalRow, 7)).Font.Bold = True\n', `Bolds the Total row from column 1 to column 7: \`A${t}:G${t}\` on this report.`, `A${t}:G${t}`),
          part('    Range("A1").CurrentRegion.Columns.AutoFit', 'Fits the column widths again now that the totals are there, so none shows as ####.', `A1:G${t}`),
        ],
        why: 'If your recorded code has an AutoFit line higher up, leave it. This one runs last.',
      },
      {
        do: 'Switch back to the workbook and click any cell on the practice sheet. Choose **{runMacro}**, select **FormatWeekly** and click **Run**.',
        why: `The report gets its formatting and a Total row in row ${t}: ${money(sales)} in \`F${t}\` and ${money(cost)} in \`G${t}\`. If Excel stops with an error, the editor highlights the line it stopped on; compare it with the cards. If it stops at an AutoFilter line with error 1004, add \`Field:=1\` to the end of that line.`,
        done: whenPasses(checks, 'values', `F${t}:G${t}`),
      },
      checkStep(
        `The coach checks for “Total” in \`A${t}\`, SUM formulas in \`F${t}:G${t}\` that add up every order, and a bold row. On the next rep, run FormatWeekly again: the report will have a different number of rows, and the Total row should still land right under the last one.`,
      ),
    ];
  },
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
  guide: (d): GuideStep[] => {
    const last = d.rows.length + 1;
    const low = lowRows(d);
    const atIndex = d.rows.findIndex((r) => r.onHand === r.reorder);
    const at = d.rows[atIndex];
    const atRow = atIndex + 2;
    const checks: Inspection[] = macroHighlightRows.inspections!(d);
    return [
      {
        do: `Meet the stock list: ${d.rows.length} SKUs, with On hand in column \`D\` and Reorder point in column \`E\`.`,
        why: `A row needs reordering when On hand is below its reorder point. On this list that’s ${low.length} rows. The next list has a different number of rows, so the macro finds the last row by itself.`,
        show: [
          { label: 'Rows to fill', at: rowBands(low), note: `These ${low.length} rows are below their reorder point. The macro should fill each one across \`A\` to \`F\`.` },
          { label: `Row ${atRow}`, at: `A${atRow}:F${atRow}`, note: `${at.onHand} on hand with a reorder point of ${at.reorder}: exactly at it, not below, so it stays unfilled.` },
          { label: 'On hand', at: `D2:D${last}` },
          { label: 'Reorder point', at: `E2:E${last}` },
        ],
      },
      {
        do: 'Open {vbaEditor}. In the project list on the left, select {personalProject}, then choose **Insert › Module**.',
        why: 'A new, empty code window opens. If {personalProject} isn’t listed, that workbook doesn’t exist yet: record any short macro into it first, or select this workbook’s project instead.',
      },
      {
        do: 'In the new window, type `Sub HighlightReorders()` and press {enter}.',
        why: 'The editor adds `End Sub` under it and leaves the cursor on the empty line between them. The rest of the macro goes there.',
      },
      {
        do: 'Type these two lines.',
        formula: [
          part('    Dim lastRow As Long, r As Long\n', 'Two whole-number variables: lastRow for the last data row, and r for the row the loop is on.'),
          part(
            '    lastRow = Cells(Rows.Count, 1).End(xlUp).Row',
            `Starts at the bottom of column \`A\` and jumps up to the last filled cell, as {jumpUp} does: row ${last} on this list.`,
            `A${last}`,
          ),
        ],
      },
      {
        do: 'Under them, type this line.',
        formula: [
          part(
            '    Range("A2:F" & lastRow).Interior.ColorIndex = xlNone',
            `Clears any fill from the data rows first (\`A2:F${last}\` here), so a row that’s been restocked loses its color the next time the macro runs.`,
            `A2:F${last}`,
          ),
        ],
      },
      {
        do: 'Under that, type the loop.',
        formula: [
          part('    For r = 2 To lastRow\n', `Repeats the lines down to \`Next r\` once per data row: r is 2, then 3, and so on to ${last}. Starting at 2 skips the header row.`, `A2:A${last}`),
          raw('        If '),
          part('Cells(r, 4).Value', 'On hand on row r. Column 4 is `D`.', `D2:D${last}`),
          part(' < ', 'Is less than. Use `<`, not `<=`: a row exactly at its reorder point isn’t below it.'),
          part('Cells(r, 5).Value', 'Reorder point on row r. Column 5 is `E`.', `E2:E${last}`),
          raw(' Then\n'),
          part('            Range(Cells(r, 1), Cells(r, 6)).Interior.Color = RGB(255, 199, 206)\n', 'Fills columns 1 to 6 (`A` to `F`) of row r. `RGB(255, 199, 206)` is a light red; any color works.'),
          raw('        End If\n'),
          part('    Next r', 'Goes back to `For` with the next row, until r passes lastRow.'),
        ],
        why: 'The indents are optional; they show which lines sit inside the loop. Conditional formatting would color the rows without changing their own fill, and the coach checks the fill, so the macro sets it with `Interior.Color`.',
      },
      {
        do: 'Switch back to the workbook and click any cell on the practice sheet. Choose **{runMacro}**, select **HighlightReorders** and click **Run**.',
        why: `The ${low.length} low rows turn light red, down to row ${last} at the bottom. If row ${atRow} fills too, the comparison is \`<=\`: change it to \`<\` and run the macro again. If Excel stops with an error, the editor highlights the line it stopped on; compare it with the cards.`,
        done: whenPasses(checks, 'filled', `A${last}:F${last}`),
      },
      checkStep(
        `The coach checks that each low row is filled across \`A\` to \`F\`, and that the header row and a sample of the other rows, row ${atRow} among them, aren’t. On the next rep, run HighlightReorders on the new list.`,
      ),
    ];
  },
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

import { describe, expect, it } from 'vitest';
import { parseRange, rangeSize } from '../src/engine/address';
import { cellMatches } from '../src/engine/compare';
import { gradeSheetCheck, type SheetFacts } from '../src/engine/sheetChecks';
import { serial } from '../src/engine/data';
import { localize } from '../src/engine/platform';
import { Rng } from '../src/engine/rng';
import type { Cell, CellsBlock, Exercise, Inspection, Platform, SheetCheck } from '../src/engine/types';
import {
  DATE_FORMAT,
  MACROS,
  MAX_ROWS,
  MIN_DESC_WIDTH,
  MIN_ROWS,
  MONEY_FORMAT,
  STOCK_HEADERS,
  TOTAL_LABEL,
  WEEKLY_HEADERS,
  lastDataRow,
  lowRows,
  macroAnyRows,
  macroHighlightRows,
  macroRecordFormat,
  okSample,
  totalRow,
  weeklyTotals,
  type ReorderData,
  type WeeklyData,
} from '../src/exercises/macros';
import { SEEDS, exerciseSuite } from './helpers/suite';

exerciseSuite(MACROS);

const MANY_SEEDS = Array.from({ length: 200 }, (_, i) => i * 7919 + 3);
const PLATFORMS: Platform[] = ['mac', 'windows', 'web'];
const BANNED = /\b(please|simply|just|easy|easily|successfully|leverage|seamless)\b|!/i;

type SheetInspection = Extract<Inspection, { kind: 'sheet' }>;

function sheetChecks(ex: Exercise<any>, d: unknown): SheetInspection[] {
  const all = ex.inspections?.(d) ?? [];
  expect(all.every((i) => i.kind === 'sheet'), `${ex.id} uses only sheet inspections`).toBe(true);
  return all as SheetInspection[];
}

function checksOf<K extends SheetCheck['kind']>(list: SheetInspection[], kind: K): Extract<SheetCheck, { kind: K }>[] {
  return list.map((i) => i.check).filter((c): c is Extract<SheetCheck, { kind: K }> => c.kind === kind);
}

/** The column letter of a header, from the order setup writes them in. */
const colOf = (headers: readonly string[], header: string) => String.fromCharCode(65 + headers.indexOf(header));

/** Prose the learner reads: everything except VBA in the syntax box and the answer. */
function prose(ex: Exercise<any>, d: unknown): string[] {
  const labels = (ex.inspections?.(d) ?? []).flatMap((i) => [i.label, ...(i.kind === 'sheet' && i.advice ? [i.advice] : [])]);
  const c = ex.concept;
  return [ex.title, ex.replaces, ex.task(d), c.summary, c.example ?? '', c.tip ?? '', ...ex.hints, ...labels].filter(Boolean);
}

describe('area basics', () => {
  it('uses unique ids in the macros module', () => {
    expect(new Set(MACROS.map((e) => e.id)).size).toBe(MACROS.length);
    expect(MACROS.every((e) => e.module === 'macros')).toBe(true);
    expect(MACROS.map((e) => e.id)).toEqual(['macro-record-format', 'macro-any-rows', 'macro-highlight-rows']);
  });

  it.each(SEEDS)('seed %i: plain sheets of unformatted input cells, graded by inspections alone', (seed) => {
    for (const ex of MACROS) {
      const d = ex.make(new Rng(seed));
      const layout = ex.layout(d);
      expect(layout.plain).toBe(true);
      expect(layout.answer).toEqual({ kind: 'objects' });
      for (const b of layout.blocks) {
        expect(b.kind).toBe('cells');
        const block = b as CellsBlock;
        expect(block.role, `${ex.id} ${block.at}`).toBe('input');
        expect(block.format).toBeUndefined();
        expect(block.formats).toBeUndefined();
      }
      expect(ex.expected(d)).toEqual([]);
      expect(ex.inputs(d)).toEqual([]);
      expect(ex.variants).toEqual([]);
      expect(sheetChecks(ex, d).length).toBeGreaterThan(0);
    }
  });

  it('gives every rep a different-sized sheet, so a recorded fixed range breaks on the next one', () => {
    for (const ex of MACROS) {
      const counts = MANY_SEEDS.map((seed) => (ex.make(new Rng(seed)) as { rows: unknown[] }).rows.length);
      expect(Math.min(...counts)).toBeGreaterThanOrEqual(MIN_ROWS);
      expect(Math.max(...counts)).toBeLessThanOrEqual(MAX_ROWS);
      expect(new Set(counts).size, `${ex.id} row counts`).toBeGreaterThan(20);
    }
  });

  it('every inspection range parses and stays inside the sheet the coach writes, plus the Total row', () => {
    for (const ex of MACROS) {
      for (const seed of SEEDS) {
        const d = ex.make(new Rng(seed)) as { rows: unknown[] };
        const extra = ex === macroAnyRows ? 1 : 0;
        for (const i of sheetChecks(ex, d)) {
          if (!('range' in i.check)) continue;
          const r = parseRange(i.check.range);
          expect(r.start.row).toBeGreaterThanOrEqual(1);
          expect(r.end.row, `${ex.id} ${i.check.range}`).toBeLessThanOrEqual(d.rows.length + 1 + extra);
          expect(r.end.col).toBeLessThanOrEqual(7);
        }
      }
    }
  });

  it('labels are unique within each exercise', () => {
    for (const ex of MACROS) {
      const labels = sheetChecks(ex, ex.make(new Rng(42))).map((i) => i.label);
      expect(new Set(labels).size).toBe(labels.length);
    }
  });
});

describe('weekly report data', () => {
  it('puts the columns where the tasks, hints and VBA say they are', () => {
    expect(colOf(WEEKLY_HEADERS, 'Date')).toBe('A');
    expect(colOf(WEEKLY_HEADERS, 'Description')).toBe('D');
    expect(colOf(WEEKLY_HEADERS, 'Sales')).toBe('F');
    expect(colOf(WEEKLY_HEADERS, 'Cost')).toBe('G');
    expect(WEEKLY_HEADERS.length).toBe(7);
  });

  it('is one Monday-to-Friday week of orders in date order, with distinct order numbers', () => {
    for (const seed of MANY_SEEDS) {
      const d = macroRecordFormat.make(new Rng(seed));
      const dates = d.rows.map((r) => r.date);
      expect(dates).toEqual([...dates].sort((a, b) => a - b));
      expect(Math.max(...dates) - Math.min(...dates)).toBeLessThanOrEqual(4);
      // serial(2026, 1, 5) is a Monday; every date is Monday to Friday.
      for (const s of dates) expect((s - serial(2026, 1, 5)) % 7).toBeLessThanOrEqual(4);
      expect(new Set(d.rows.map((r) => r.order)).size).toBe(d.rows.length);
    }
  });

  it('writes dates as serial numbers and money as plain numbers, so the macro has formatting to do', () => {
    const d = macroRecordFormat.make(new Rng(7));
    const data = macroRecordFormat.layout(d).blocks[1] as CellsBlock;
    expect(data.at).toBe('A2');
    expect(data.values.length).toBe(d.rows.length);
    for (const row of data.values) {
      expect(row.length).toBe(WEEKLY_HEADERS.length);
      expect(Number.isInteger(row[0])).toBe(true);
      expect(typeof row[5]).toBe('number');
      expect(typeof row[6]).toBe('number');
    }
  });

  it('keeps money to whole cents, with sales above cost', () => {
    for (const seed of MANY_SEEDS) {
      for (const r of macroRecordFormat.make(new Rng(seed)).rows) {
        expect(Math.round(r.sales * 100) / 100).toBe(r.sales);
        expect(Math.round(r.cost * 100) / 100).toBe(r.cost);
        expect(r.cost).toBeGreaterThan(0);
        expect(r.sales).toBeGreaterThan(r.cost);
        expect(Number.isInteger(r.units)).toBe(true);
      }
    }
  });

  it('every description is long enough that autofit must widen column D well past the default', () => {
    let shortest = Infinity;
    for (const seed of MANY_SEEDS) for (const r of macroRecordFormat.make(new Rng(seed)).rows) shortest = Math.min(shortest, r.description.length);
    expect(shortest).toBeGreaterThanOrEqual(45);
    // A narrow font averages over 4.2 points a character; the default width is at most about 65 points.
    expect(shortest * 4.2).toBeGreaterThan(MIN_DESC_WIDTH * 1.3);
    expect(MIN_DESC_WIDTH).toBeGreaterThan(65 * 2);
  });
});

describe('macro-record-format', () => {
  it.each(SEEDS)('seed %i: inspection ranges follow this rep’s row count', (seed) => {
    const d = macroRecordFormat.make(new Rng(seed));
    const last = lastDataRow(d);
    expect(last).toBe(d.rows.length + 1);
    const list = sheetChecks(macroRecordFormat, d);

    expect(checksOf(list, 'bold')).toEqual([{ kind: 'bold', range: 'A1:G1', bold: true }]);
    expect(checksOf(list, 'filled')).toEqual([{ kind: 'filled', range: 'A1:G1', filled: true }]);
    const formats = checksOf(list, 'numberFormat');
    expect(formats.map((f) => f.range)).toEqual([`A2:A${last}`, `F2:G${last}`]);
    expect(formats[0].matches).toBe(DATE_FORMAT);
    expect(formats[1].matches).toBe(MONEY_FORMAT);
    expect(checksOf(list, 'minWidth')).toEqual([{ kind: 'minWidth', range: `D1:D${last}`, points: MIN_DESC_WIDTH }]);
    expect(checksOf(list, 'freeze')).toEqual([{ kind: 'freeze', rows: 1, cols: 0 }]);
    expect(checksOf(list, 'filter')).toEqual([{ kind: 'filter', on: true }]);
    expect(macroRecordFormat.task(d)).toContain(`\`F2:G${last}\``);
    expect(macroRecordFormat.task(d)).toContain(`\`A2:A${last}\``);
  });

  it('a range recorded on a shorter rep doesn’t cover a longer one', () => {
    const sizes = MANY_SEEDS.map((seed) => macroRecordFormat.make(new Rng(seed)).rows.length);
    const short = MANY_SEEDS[sizes.indexOf(Math.min(...sizes))];
    const long = MANY_SEEDS[sizes.indexOf(Math.max(...sizes))];
    const range = (seed: number) => checksOf(sheetChecks(macroRecordFormat, macroRecordFormat.make(new Rng(seed))), 'numberFormat')[1].range;
    expect(rangeSize(range(long)).rows).toBeGreaterThan(rangeSize(range(short)).rows);
  });
});

describe('number format patterns', () => {
  const money = [
    '$#,##0.00',
    '$#,##0',
    '$0.00',
    '\\$#,##0.00',
    '"$"#,##0.00',
    '$#,##0.00_);($#,##0.00)',
    '$#,##0.00_);[Red]($#,##0.00)',
    '"$"#,##0.00_);[Red]\\("$"#,##0.00\\)',
    '$#,##0_);($#,##0)',
    '[$$-409]#,##0.00',
    '[$$-en-US]#,##0.00',
    '[$€-x-euro2] #,##0.00',
    '#,##0.00',
    '#,##0.00_);(#,##0.00)',
    '#,##0.00_);[Red](#,##0.00)',
    '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)',
    '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)',
    '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)',
    '_(* #,##0_);_(* \\(#,##0\\);_(* "-"_);_(@_)',
  ];
  const notMoney = ['General', '0', '0.00', '#,##0', '0%', '0.00%', '$0.00%', '@', 'm/d/yyyy', 'yyyy-mm-dd', '[$-409]m/d/yyyy', '[$-409]mmmm d, yyyy;@', 'h:mm AM/PM', ''];

  it.each(money)('counts %s as currency', (f) => expect(MONEY_FORMAT.test(f)).toBe(true));
  it.each(notMoney)('does not count %s as currency', (f) => expect(MONEY_FORMAT.test(f)).toBe(false));

  const dates = ['m/d/yyyy', 'm/d/yy', 'mm/dd/yyyy', 'yyyy-mm-dd', 'd-mmm-yy', 'd-mmm', 'mmm d, yyyy', 'dddd, mmmm d, yyyy', '[$-409]mmmm d, yyyy;@', '[$-F800]dddd, mmmm dd, yyyy', 'm/d/yyyy h:mm', 'DD/MM/YYYY'];
  const notDates = ['General', '0', '0.00', '$#,##0.00', '#,##0', '0%', 'h:mm', 'h:mm:ss AM/PM', '[h]:mm:ss', '[Red]0.00', '0 "days"', '@', ''];

  it.each(dates)('counts %s as a date', (f) => expect(DATE_FORMAT.test(f)).toBe(true));
  it.each(notDates)('does not count %s as a date', (f) => expect(DATE_FORMAT.test(f)).toBe(false));

  it('accepts the formats the answer’s VBA writes', () => {
    for (const ex of [macroRecordFormat, macroAnyRows]) {
      const vba = ex.solution(ex.make(new Rng(1)));
      expect(DATE_FORMAT.test(/Columns\("A"\)\.NumberFormat = "([^"]+)"/.exec(vba)![1])).toBe(true);
      expect(MONEY_FORMAT.test(/Columns\("F:G"\)\.NumberFormat = "([^"]+)"/.exec(vba)![1])).toBe(true);
    }
  });

  it('carries no global flag, so repeated tests can’t skip cells', () => {
    expect(MONEY_FORMAT.global || MONEY_FORMAT.sticky || DATE_FORMAT.global || DATE_FORMAT.sticky).toBe(false);
  });
});

describe('macro-any-rows', () => {
  /** Totals added up in whole cents, independently of the exercise’s own sum. */
  const centsTotal = (values: number[]) => values.reduce((a, v) => a + Math.round(v * 100), 0) / 100;

  it.each(SEEDS)('seed %i: checks the row right under this rep’s data', (seed) => {
    const d = macroAnyRows.make(new Rng(seed));
    const t = totalRow(d);
    expect(t).toBe(d.rows.length + 2);
    const list = sheetChecks(macroAnyRows, d);
    const values = checksOf(list, 'values');
    expect(values.map((v) => v.range)).toEqual([`A${t}`, `F${t}:G${t}`]);
    expect(checksOf(list, 'formulas')).toEqual([{ kind: 'formulas', range: `F${t}:G${t}`, formulas: true, pattern: /SUM\(/i }]);
    expect(checksOf(list, 'bold')).toEqual([{ kind: 'bold', range: `A${t}:G${t}`, bold: true }]);
    expect(macroAnyRows.task(d)).toContain(`\`A${t}\``);
    expect(macroAnyRows.task(d)).toContain(`\`F${t}:G${t}\``);
  });

  it('expects totals that add up every data row', () => {
    for (const seed of MANY_SEEDS) {
      const d = macroAnyRows.make(new Rng(seed));
      const [sales, cost] = weeklyTotals(d);
      expect(sales).toBe(centsTotal(d.rows.map((r) => r.sales)));
      expect(cost).toBe(centsTotal(d.rows.map((r) => r.cost)));
      const values = checksOf(sheetChecks(macroAnyRows, d), 'values');
      expect(values[1].expected).toEqual([[sales, cost]]);
    }
  });

  it('a total that stops one row short, as a recorded AutoSum would on a longer rep, fails', () => {
    const d = macroAnyRows.make(new Rng(42));
    const [sales] = weeklyTotals(d);
    const short = centsTotal(d.rows.slice(0, -1).map((r) => r.sales));
    expect(cellMatches(short, sales)).toBe(false);
    expect(cellMatches(sales, sales)).toBe(true);
  });

  it('accepts “Total” in any case and nothing else', () => {
    for (const ok of ['Total', 'TOTAL', 'total', ' Total ']) expect(cellMatches(ok, TOTAL_LABEL), ok).toBe(true);
    for (const bad of ['', 'Totals', 'Grand total', 0]) expect(cellMatches(bad, TOTAL_LABEL), String(bad)).toBe(false);
    const d = macroAnyRows.make(new Rng(3));
    expect(checksOf(sheetChecks(macroAnyRows, d), 'values')[0].expected).toEqual([[TOTAL_LABEL]]);
  });

  it('the SUM pattern takes the formulas a macro writes and rejects typed totals', () => {
    const pattern = checksOf(sheetChecks(macroAnyRows, macroAnyRows.make(new Rng(1))), 'formulas')[0].pattern!;
    for (const f of ['=SUM(F2:F31)', '=SUM(F$2:F31)', '=SUM($F$2:$F$31)']) expect(pattern.test(f), f).toBe(true);
    for (const f of ['12345.67', '=F2+F3', '=SUMIF(A2:A31,"x",F2:F31)']) expect(pattern.test(f), f).toBe(false);
  });
});

describe('macro-highlight-rows', () => {
  it('puts On hand in D, Reorder point in E and the last column in F, as the VBA assumes', () => {
    expect(colOf(STOCK_HEADERS, 'On hand')).toBe('D');
    expect(colOf(STOCK_HEADERS, 'Reorder point')).toBe('E');
    expect(STOCK_HEADERS.length).toBe(6);
  });

  it('every rep has at least two low rows, two that are fine and one exactly at its reorder point', () => {
    for (const seed of MANY_SEEDS) {
      const d = macroHighlightRows.make(new Rng(seed));
      const low = d.rows.filter((r) => r.onHand < r.reorder).length;
      const ok = d.rows.length - low;
      expect(low, `seed ${seed}`).toBeGreaterThanOrEqual(2);
      expect(ok, `seed ${seed}`).toBeGreaterThanOrEqual(2);
      expect(d.rows.filter((r) => r.onHand === r.reorder).length).toBe(1);
      expect(d.rows.every((r) => r.onHand >= 0 && r.reorder > 0)).toBe(true);
      expect(new Set(d.rows.map((r) => r.sku)).size).toBe(d.rows.length);
    }
  });

  it('makes the first and last data rows low, so a loop that starts late or stops early misses one', () => {
    for (const seed of MANY_SEEDS) {
      const d = macroHighlightRows.make(new Rng(seed));
      const low = lowRows(d);
      expect(low[0], `seed ${seed}`).toBe(2);
      expect(low[low.length - 1], `seed ${seed}`).toBe(d.rows.length + 1);
      expect(low.length).toBeGreaterThanOrEqual(3);
      expect(low.length).toBeLessThanOrEqual(6);
    }
  });

  it('low rows land in different places from rep to rep', () => {
    const middles = new Set(MANY_SEEDS.flatMap((seed) => lowRows(macroHighlightRows.make(new Rng(seed))).slice(1, -1)));
    expect(middles.size).toBeGreaterThan(20);
  });

  it.each(SEEDS)('seed %i: checks the header and every low row, and samples the others', (seed) => {
    const d = macroHighlightRows.make(new Rng(seed));
    const list = sheetChecks(macroHighlightRows, d);
    const filled = checksOf(list, 'filled');
    expect(filled.length).toBe(list.length);
    const want = d.rows.flatMap((r, i) => (r.onHand < r.reorder ? [i + 2] : []));

    // The header first: a loop from row 1 compares the header text and fills it.
    expect(filled[0]).toEqual({ kind: 'filled', range: 'A1:F1', filled: false });
    const body = filled.slice(1);

    const on = body.filter((c) => c.filled).map((c) => c.range);
    expect(on).toEqual(want.map((row) => `A${row}:F${row}`));

    const off = body.filter((c) => !c.filled).map((c) => parseRange(c.range));
    expect(off.length).toBe(3);
    for (const r of off) {
      expect(r.start.row).toBe(r.end.row);
      expect([r.start.col, r.end.col]).toEqual([1, 6]);
      const line = d.rows[r.start.row - 2];
      expect(line.onHand).toBeGreaterThanOrEqual(line.reorder);
    }
    expect(off.some((r) => d.rows[r.start.row - 2].onHand === d.rows[r.start.row - 2].reorder), 'samples the row at its reorder point').toBe(true);

    // In sheet order, so the results read top to bottom.
    const rows = filled.map((c) => parseRange(c.range).start.row);
    expect(rows).toEqual([...rows].sort((a, b) => a - b));
  });

  it('samples rows by their data, not by position', () => {
    const d: ReorderData = {
      rows: [
        { sku: 'SKU-1101', item: 'Broom', warehouse: 'Reno', onHand: 3, reorder: 10, supplier: 'Apex Supply' },
        { sku: 'SKU-1102', item: 'Mop head', warehouse: 'Reno', onHand: 40, reorder: 10, supplier: 'Apex Supply' },
        { sku: 'SKU-1103', item: 'Hard hat', warehouse: 'Dallas', onHand: 15, reorder: 15, supplier: 'Cobalt Parts' },
        { sku: 'SKU-1104', item: 'Spill kit', warehouse: 'Atlanta', onHand: 22, reorder: 20, supplier: 'Delta Packaging' },
        { sku: 'SKU-1105', item: 'Ear plugs', warehouse: 'Columbus', onHand: 0, reorder: 50, supplier: 'Evergreen Mfg' },
        { sku: 'SKU-1106', item: 'Broom', warehouse: 'Dallas', onHand: 90, reorder: 30, supplier: 'Apex Supply' },
      ],
    };
    expect(lowRows(d)).toEqual([2, 6]);
    expect(okSample(d)).toEqual([3, 4, 7]);
  });
});

describe('teaching content', () => {
  it('follows the copy rules in every string the learner reads, on every platform', () => {
    for (const ex of MACROS) {
      for (const seed of [1, 42]) {
        const d = ex.make(new Rng(seed));
        for (const text of prose(ex, d)) {
          expect(text, `copy: ${text}`).not.toMatch(BANNED);
          expect(text, `straight apostrophe: ${text}`).not.toMatch(/(?<=[A-Za-z])'(?=[A-Za-z])/);
          expect(text, `menu path: ${text}`).not.toMatch(/ > /);
          for (const p of PLATFORMS) expect(localize(text, p), `placeholder left in: ${text}`).not.toMatch(/[{}]/);
        }
        for (const text of [ex.concept.syntax, ex.solution(d)]) for (const p of PLATFORMS) expect(localize(text, p)).not.toMatch(/[{}]/);
      }
    }
  });

  it('labels and advice skip backticks and placeholders, which the results list doesn’t render', () => {
    for (const ex of MACROS) {
      for (const i of sheetChecks(ex, ex.make(new Rng(7)))) {
        for (const text of [i.label, i.advice ?? '']) expect(text).not.toMatch(/[`{}]/);
        expect(i.advice, `${ex.id}: ${i.label} has advice`).toBeTruthy();
      }
    }
  });

  it('hints walk through the macro menus with the platform placeholders', () => {
    const all = MACROS.flatMap((e) => e.hints).join(' ');
    for (const key of ['{recordMacro}', '{runMacro}', '{vbaEditor}']) expect(all).toContain(key);
    for (const ex of MACROS) expect(ex.hints.some((h) => /\{(recordMacro|runMacro|vbaEditor)\}/.test(h)), ex.id).toBe(true);
    expect(macroRecordFormat.hints.join(' ')).toMatch(/Developer tab/);
    expect(macroAnyRows.concept.tip).toMatch(/Use Relative References/);
    expect(macroRecordFormat.concept.tip).toMatch(/Personal Macro Workbook[\s\S]*\.xlsm/);
    expect(macroHighlightRows.concept.summary).toMatch(/Conditional formatting/);
  });

  it('localizes the recorder path for each platform', () => {
    const first = macroRecordFormat.hints[0];
    expect(localize(first, 'mac')).toContain('Tools › Macro › Record New Macro');
    expect(localize(first, 'windows')).toContain('View › Macros › Record Macro');
  });
});

describe('VBA answers', () => {
  const answers = MACROS.map((ex) => [ex.id, ex.solution(ex.make(new Rng(1)))] as const);

  it.each(answers)('%s: plain ASCII, so it pastes cleanly into the Visual Basic Editor', (_id, vba) => {
    expect(/^[\x09\x0a\x20-\x7e]*$/.test(vba)).toBe(true);
  });

  it.each(answers)('%s: one complete Sub with balanced blocks', (_id, vba) => {
    const lines = vba.split('\n').map((l) => l.replace(/'.*$/, '').trim());
    const count = (re: RegExp) => lines.filter((l) => re.test(l)).length;
    expect(count(/^Sub \w+\(\)$/)).toBe(1);
    expect(count(/^End Sub$/)).toBe(1);
    expect(count(/^With\b/)).toBe(count(/^End With$/));
    expect(count(/^For\b/)).toBe(count(/^Next\b/));
    expect(count(/^If\b.*\bThen$/)).toBe(count(/^End If$/));
  });

  it('FormatWeekly freezes, filters and autofits the way that works on Mac and Windows', () => {
    const vba = macroRecordFormat.solution(macroRecordFormat.make(new Rng(1)));
    expect(vba).toContain('Sub FormatWeekly()');
    const order = ['ActiveWindow.FreezePanes = False', 'ActiveWindow.SplitColumn = 0', 'ActiveWindow.SplitRow = 1', 'ActiveWindow.FreezePanes = True'].map((s) => vba.indexOf(s));
    expect(order.every((i) => i >= 0)).toBe(true);
    expect(order).toEqual([...order].sort((a, b) => a - b));
    // Excel for Mac raises error 1004 on AutoFilter with no arguments (seen in Excel 16.113).
    expect(vba).toMatch(/If Not ActiveSheet\.AutoFilterMode Then\s+Range\("A1"\)\.CurrentRegion\.AutoFilter Field:=1/);
    expect(vba).toContain('.Columns.AutoFit');
    expect(vba.indexOf('.Columns.AutoFit')).toBeGreaterThan(vba.indexOf('NumberFormat'));
  });

  it('the any-rows and highlight answers find the last row instead of using fixed addresses', () => {
    for (const ex of [macroAnyRows, macroHighlightRows] as Exercise<any>[]) {
      const vba = ex.solution(ex.make(new Rng(1)));
      expect(vba).toContain('Cells(Rows.Count, 1).End(xlUp).Row');
      // No cell address past row 2: that would be this rep’s data, recorded.
      expect(vba, ex.id).not.toMatch(/\b\$?[A-Z]{1,3}\$?(?:[3-9]|[1-9]\d+)\b/);
    }
    const anyRows = macroAnyRows.solution(macroAnyRows.make(new Rng(1)));
    expect(anyRows).toContain('Sub FormatWeekly()');
    expect(anyRows).toContain('=SUM(R2C:R[-1]C)');
    expect(anyRows).toMatch(/\.Font\.Bold = True/);
    const highlight = macroHighlightRows.solution(macroHighlightRows.make(new Rng(1)));
    expect(highlight).toContain('If Cells(r, 4).Value < Cells(r, 5).Value Then');
    expect(highlight).toContain('.Interior.Color = RGB(');
    expect(highlight).toContain('.Interior.ColorIndex = xlNone');
  });

  it('the any-rows answer still formats the report, so it can replace FormatWeekly whole', () => {
    const anyRows = macroAnyRows.solution(macroAnyRows.make(new Rng(1)));
    const format = macroRecordFormat.solution(macroRecordFormat.make(new Rng(1)));
    for (const line of ['Columns("A").NumberFormat', 'Columns("F:G").NumberFormat', 'ActiveWindow.SplitRow = 1', 'Range("A1").CurrentRegion.AutoFilter Field:=1', '.Columns.AutoFit']) {
      expect(format).toContain(line);
      expect(anyRows).toContain(line);
    }
  });
});

// ---------- graded end to end: what each macro leaves on the sheet ----------

/** A practice sheet after a macro ran, as cell-by-cell facts. Rows and columns are 1-based. */
interface SheetModel {
  value?(row: number, col: number): Cell;
  formula?(row: number, col: number): Cell;
  format?(row: number, col: number): string;
  bold?(row: number, col: number): boolean;
  fill?(row: number, col: number): string | null;
  width?(col: number): number;
  freeze?: { rows: number; cols: number };
  filter?: boolean;
}

/** What the host would read from `model` for one check. */
function factsFor(check: SheetCheck, model: SheetModel): SheetFacts {
  const grid = <T,>(range: string, at: (row: number, col: number) => T): T[][] => {
    const { start, end } = parseRange(range);
    return Array.from({ length: end.row - start.row + 1 }, (_, r) => Array.from({ length: end.col - start.col + 1 }, (_, c) => at(start.row + r, start.col + c)));
  };
  switch (check.kind) {
    case 'freeze':
      return { kind: 'freeze', ...(model.freeze ?? { rows: 0, cols: 0 }) };
    case 'filter':
      return { kind: 'filter', on: model.filter ?? false };
    case 'numberFormat':
      return { kind: 'numberFormat', address: check.range, formats: grid(check.range, model.format ?? (() => 'General')) };
    case 'bold':
      return { kind: 'bold', address: check.range, bold: grid(check.range, model.bold ?? (() => false)) };
    case 'filled':
      return { kind: 'filled', address: check.range, colors: grid(check.range, model.fill ?? (() => '#FFFFFF')) };
    case 'minWidth': {
      const { start, end } = parseRange(check.range);
      return { kind: 'minWidth', address: check.range, widths: Array.from({ length: end.col - start.col + 1 }, (_, i) => (model.width ?? (() => 48))(start.col + i)) };
    }
    case 'values':
      return { kind: 'values', address: check.range, values: grid(check.range, model.value ?? (() => '')) };
    case 'formulas':
      return { kind: 'formulas', address: check.range, formulas: grid(check.range, model.formula ?? model.value ?? (() => '')) };
    default:
      throw new Error(`No model for ${check.kind}`);
  }
}

/** Labels of the checks that fail on `model`, using the real grader. */
function failures(ex: Exercise<any>, d: unknown, model: SheetModel): string[] {
  return sheetChecks(ex, d)
    .map((i) => gradeSheetCheck(i, factsFor(i.check, model)))
    .filter((item) => item.status !== 'pass')
    .map((item) => item.label);
}

/** The sheet after FormatWeekly. `formattedTo` is the last row its formats reach (a recorded fixed range). */
function formatted(d: WeeklyData, formattedTo = Infinity): SheetModel {
  const longest = Math.max(...d.rows.map((r) => r.description.length));
  return {
    format: (row, col) => (row > formattedTo ? 'General' : col === 1 ? 'yyyy-mm-dd' : col === 6 || col === 7 ? '$#,##0.00' : 'General'),
    bold: (row) => row === 1,
    fill: (row) => (row === 1 ? '#DDEBF7' : '#FFFFFF'),
    // Autofit in a narrow font: about 4.2 points a character.
    width: (col) => (col === 4 ? longest * 4.2 : 60),
    freeze: { rows: 1, cols: 0 },
    filter: true,
  };
}

/** The sheet after the Total row is written at `row`, summing `from` through the row above. */
function withTotals(d: WeeklyData, row: number, from = 2): SheetModel {
  const data = d.rows.slice(from - 2, row - 2);
  // Excel's SUM adds left to right in doubles; so does reduce.
  const sales = data.reduce((a, r) => a + r.sales, 0);
  const cost = data.reduce((a, r) => a + r.cost, 0);
  return {
    value: (r, c) => (r !== row ? '' : c === 1 ? 'Total' : c === 6 ? sales : c === 7 ? cost : ''),
    formula: (r, c) => (r !== row ? '' : c === 1 ? 'Total' : c === 6 ? `=SUM(F$${from}:F${row - 1})` : c === 7 ? `=SUM(G$${from}:G${row - 1})` : ''),
    bold: (r) => r === row,
  };
}

interface Loop {
  /** How the macro compares On hand with Reorder point. */
  compare?: (onHand: string | number, reorder: string | number) => boolean;
  /** The loop's first and last sheet rows. */
  start?: number;
  end?: number;
}

/**
 * The sheet after HighlightReorders, written with this comparison and loop. Row 1 compares the
 * header text, as VBA does: "On hand" < "Reorder point" is True.
 */
function highlighted(d: ReorderData, { compare = (a, b) => a < b, start = 2, end = d.rows.length + 1 }: Loop = {}): SheetModel {
  const cell = (row: number, col: 4 | 5): string | number => (row === 1 ? STOCK_HEADERS[col - 1] : col === 4 ? d.rows[row - 2].onHand : d.rows[row - 2].reorder);
  return {
    fill: (row) => (row >= start && row <= Math.min(end, d.rows.length + 1) && compare(cell(row, 4), cell(row, 5)) ? '#FFC7CE' : '#FFFFFF'),
  };
}

describe('graded by the sheet checks', () => {
  it.each(SEEDS)('seed %i: FormatWeekly passes every check; an untouched sheet fails every one', (seed) => {
    const d = macroRecordFormat.make(new Rng(seed));
    expect(failures(macroRecordFormat, d, formatted(d))).toEqual([]);
    expect(failures(macroRecordFormat, d, {}).length).toBe(sheetChecks(macroRecordFormat, d).length);
  });

  it('formats recorded on a shorter rep fail on a longer one', () => {
    const d = macroRecordFormat.make(new Rng(SEEDS.find((s) => macroRecordFormat.make(new Rng(s)).rows.length > MIN_ROWS + 2)!));
    const recordedOn = MIN_ROWS + 1; // the last sheet row of the 12-row rep it was recorded on
    const failed = failures(macroRecordFormat, d, formatted(d, recordedOn));
    expect(failed).toHaveLength(2);
    expect(failed.join(' ')).toMatch(/Dates in A2:A\d+[\s\S]*Sales and Cost/);
  });

  it.each(SEEDS)('seed %i: a Total row right under the data passes; one stuck where it was recorded fails', (seed) => {
    const d = macroAnyRows.make(new Rng(seed));
    const t = totalRow(d);
    expect(failures(macroAnyRows, d, withTotals(d, t))).toEqual([]);
    expect(failures(macroAnyRows, d, withTotals(d, t === 30 ? 25 : 30)).length).toBe(4);
  });

  it('a recorded AutoSum that adds a fixed number of rows gets the totals wrong on a longer rep', () => {
    const seed = SEEDS.find((s) => macroAnyRows.make(new Rng(s)).rows.length > 20)!;
    const d = macroAnyRows.make(new Rng(seed));
    const t = totalRow(d);
    // Recorded on a 20-row rep: =SUM(R[-20]C:R[-1]C) starts 20 rows up, not at row 2.
    const failed = failures(macroAnyRows, d, withTotals(d, t, t - 20));
    expect(failed).toEqual([`Sales and Cost totals in F${t}:G${t}`]);
  });

  it.each(SEEDS)('seed %i: HighlightReorders passes; <= fails on the row at its reorder point', (seed) => {
    const d = macroHighlightRows.make(new Rng(seed));
    expect(failures(macroHighlightRows, d, highlighted(d))).toEqual([]);

    const atPoint = d.rows.findIndex((r) => r.onHand === r.reorder) + 2;
    expect(failures(macroHighlightRows, d, highlighted(d, { compare: (a, b) => a <= b }))).toEqual([`Row ${atPoint} has no fill (exactly at its reorder point of ${d.rows[atPoint - 2].reorder})`]);
  });

  it('a loop that stops one row short, starts at row 3 or starts on the header fails on every rep', () => {
    for (const seed of MANY_SEEDS) {
      const d = macroHighlightRows.make(new Rng(seed));
      const lastRow = d.rows.length + 1;
      const rowLabel = (row: number) => expect.stringMatching(new RegExp(`^Row ${row} is filled `));
      // For r = 2 To lastRow - 1, or a loop end fixed on a rep one row shorter.
      expect(failures(macroHighlightRows, d, highlighted(d, { end: lastRow - 1 })), `seed ${seed}`).toEqual([rowLabel(lastRow)]);
      expect(failures(macroHighlightRows, d, highlighted(d, { start: 3 })), `seed ${seed}`).toEqual([rowLabel(2)]);
      expect(failures(macroHighlightRows, d, highlighted(d, { start: 1 })), `seed ${seed}`).toEqual(['Header row has no fill']);
    }
  });

  it.each(SEEDS)('seed %i: an untouched sheet fails every low-row check', (seed) => {
    const d = macroHighlightRows.make(new Rng(seed));
    const failed = failures(macroHighlightRows, d, {});
    expect(failed).toHaveLength(lowRows(d).length);
    expect(failed.every((label) => / is filled /.test(label))).toBe(true);
  });
});

import { afterEach, describe, expect, it, vi } from 'vitest';
import { gradeSheetCheck, gradeValidationList, hiddenTarget, isNoFill, parseListSource, type SheetFacts, type ValidationFacts } from '../src/engine/sheetChecks';
import type { CheckItem, Grid, Inspection, SheetCheck } from '../src/engine/types';
import { frozenPanes, inspectSheet, inspectValidation } from '../src/excel/sheetInspect';

type SheetInspection = Extract<Inspection, { kind: 'sheet' }>;
type ValidationInspection = Extract<Inspection, { kind: 'validationList' }>;

const sheet = (check: SheetCheck, advice?: string): SheetInspection => ({ kind: 'sheet', check, label: 'A sheet fact', advice });
const grade = (check: SheetCheck, facts: SheetFacts, advice?: string) => gradeSheetCheck(sheet(check, advice), facts);

/** Every failing detail produced in this file, for the copy check at the end. */
const details: string[] = [];
const failed = (item: CheckItem): CheckItem => {
  expect(item.status, item.detail).toBe('fail');
  details.push(item.detail ?? '');
  return item;
};

const CURRENCY = /^(\[\$\$-409\]|\\?\$|"\$")#,##0\.00/;

describe('freeze', () => {
  const check: SheetCheck = { kind: 'freeze', rows: 1, cols: 0 };

  it('passes on exactly the frozen rows and columns', () => {
    expect(grade(check, { kind: 'freeze', rows: 1, cols: 0 }).status).toBe('pass');
    expect(grade({ kind: 'freeze', rows: 1, cols: 1 }, { kind: 'freeze', rows: 1, cols: 1 }).status).toBe('pass');
  });

  it('says where to freeze when nothing is frozen', () => {
    const item = failed(grade(check, { kind: 'freeze', rows: 0, cols: 0 }));
    expect(item.detail).toBe('Nothing is frozen yet. Select A2, then View › Freeze Panes to keep row 1 in view.');
    expect(item.focus).toBe('A2');
    expect(failed(grade({ kind: 'freeze', rows: 1, cols: 1 }, { kind: 'freeze', rows: 0, cols: 0 })).detail).toContain('Select B2');
  });

  it('names what is frozen when it is the wrong amount', () => {
    const item = failed(grade(check, { kind: 'freeze', rows: 2, cols: 0 }));
    expect(item.detail).toBe('Rows 1–2 are frozen instead of row 1. Unfreeze from View › Freeze Panes, select A2, then freeze again.');
    expect(failed(grade(check, { kind: 'freeze', rows: 1, cols: 1 })).detail).toMatch(/^Row 1 and column A are frozen instead of row 1\./);
    expect(failed(grade({ kind: 'freeze', rows: 2, cols: 3 }, { kind: 'freeze', rows: 0, cols: 1 })).detail).toContain('Column A is frozen instead of rows 1–2 and columns A–C');
  });

  it('asks to unfreeze when nothing should be frozen', () => {
    const item = failed(grade({ kind: 'freeze', rows: 0, cols: 0 }, { kind: 'freeze', rows: 1, cols: 0 }));
    expect(item.detail).toBe('Row 1 is frozen. Unfreeze from View › Freeze Panes.');
    expect(item.focus).toBeUndefined();
  });

  it('reads Excel’s frozen location: entire rows mean no columns, entire columns mean no rows', () => {
    expect(frozenPanes(null)).toEqual({ rows: 0, cols: 0 });
    expect(frozenPanes({ rowCount: 1, columnCount: 16384 })).toEqual({ rows: 1, cols: 0 });
    expect(frozenPanes({ rowCount: 1048576, columnCount: 2 })).toEqual({ rows: 0, cols: 2 });
    expect(frozenPanes({ rowCount: 3, columnCount: 1 })).toEqual({ rows: 3, cols: 1 });
  });
});

describe('filter', () => {
  it('passes when the buttons match', () => {
    expect(grade({ kind: 'filter', on: true }, { kind: 'filter', on: true }).status).toBe('pass');
    expect(grade({ kind: 'filter', on: false }, { kind: 'filter', on: false }).status).toBe('pass');
  });

  it('explains how to turn them on or off', () => {
    expect(failed(grade({ kind: 'filter', on: true }, { kind: 'filter', on: false })).detail).toBe('There are no filter buttons yet. Click in the data, then Data › Filter.');
    expect(failed(grade({ kind: 'filter', on: false }, { kind: 'filter', on: true })).detail).toContain('Turn them off with Data › Filter.');
  });
});

describe('numberFormat', () => {
  const check: SheetCheck = { kind: 'numberFormat', range: 'D2:D5', matches: CURRENCY, describe: 'currency with 2 decimals' };
  const facts = (formats: string[]): SheetFacts => ({ kind: 'numberFormat', address: 'D2:D5', formats: formats.map((f) => [f]) });

  it('accepts every way Excel writes the format', () => {
    expect(grade(check, facts(['$#,##0.00', '\\$#,##0.00_);[Red]\\(\\$#,##0.00\\)', '[$$-409]#,##0.00', '"$"#,##0.00'])).status).toBe('pass');
  });

  it('names the first cell that doesn’t match', () => {
    const item = failed(grade(check, facts(['$#,##0.00', 'General', '$#,##0.00', '$#,##0.00'])));
    expect(item.detail).toBe('D3 uses General; it should be currency with 2 decimals.');
    expect(item.focus).toBe('D3');
  });

  it('counts several and quotes the format found', () => {
    const item = failed(grade(check, facts(['$#,##0.00', '0.00', '#,##0', '$#,##0.00'])));
    expect(item.detail).toBe('2 cells in D2:D5 aren’t currency with 2 decimals, starting at D3 (“0.00”).');
  });

  it('treats a format Office.js reports as blank as General', () => {
    expect(failed(grade(check, facts(['$#,##0.00', '', '$#,##0.00', '$#,##0.00']))).detail).toContain('uses General');
  });

  it('isn’t thrown off by a global pattern', () => {
    const global: SheetCheck = { ...check, matches: /#,##0\.00/g } as SheetCheck;
    expect(grade(global, facts(['$#,##0.00', '$#,##0.00', '$#,##0.00', '$#,##0.00'])).status).toBe('pass');
  });

  it('works across a block of cells in reading order', () => {
    const block: SheetCheck = { kind: 'numberFormat', range: 'B2:C3', matches: /0\.0%/, describe: 'a percentage with 1 decimal' };
    const item = failed(grade(block, { kind: 'numberFormat', address: 'B2:C3', formats: [['0.0%', '0.0%'], ['0.0%', 'General']] }));
    expect(item.focus).toBe('C3');
  });

  it('handles whole columns, read from where the used part starts', () => {
    const columns: SheetCheck = { kind: 'numberFormat', range: 'F:G', matches: CURRENCY, describe: 'currency with 2 decimals' };
    const item = failed(grade(columns, { kind: 'numberFormat', address: 'F3:G4', formats: [['$#,##0.00', '$#,##0.00'], ['$#,##0.00', 'General']] }));
    expect(item.detail).toBe('G4 uses General; it should be currency with 2 decimals.');
    expect(item.focus).toBe('G4');
    // The host's read when nothing in the columns is used, and Excel's own whole-column address.
    expect(grade(columns, { kind: 'numberFormat', address: 'F:G', formats: [] }).status).toBe('pass');
    expect(failed(grade({ ...columns, range: 'D:D' } as SheetCheck, { kind: 'numberFormat', address: '$D:$D', formats: [['General']] })).focus).toBe('D1');
  });
});

describe('bold', () => {
  const facts = (bold: boolean[][]): SheetFacts => ({ kind: 'bold', address: 'A1:C1', bold });

  it('passes when every cell is bold, or none is', () => {
    expect(grade({ kind: 'bold', range: 'A1:C1', bold: true }, facts([[true, true, true]])).status).toBe('pass');
    expect(grade({ kind: 'bold', range: 'A1:C1', bold: false }, facts([[false, false, false]])).status).toBe('pass');
  });

  it('names the first cell that isn’t bold', () => {
    const one = failed(grade({ kind: 'bold', range: 'A1:C1', bold: true }, facts([[true, false, true]])));
    expect(one.detail).toBe('B1 isn’t bold.');
    expect(one.focus).toBe('B1');
    expect(failed(grade({ kind: 'bold', range: 'A1:C1', bold: true }, facts([[true, false, false]]))).detail).toBe('2 cells in A1:C1 aren’t bold, starting at B1.');
  });

  it('flags bold cells when none should be', () => {
    expect(failed(grade({ kind: 'bold', range: 'A1:C1', bold: false }, facts([[false, false, true]]))).detail).toBe('C1 is bold; it shouldn’t be.');
  });

  it('handles whole rows', () => {
    const row: SheetCheck = { kind: 'bold', range: '1:1', bold: true };
    expect(failed(grade(row, { kind: 'bold', address: '1:1', bold: [[true, false]] })).detail).toBe('B1 isn’t bold.');
    expect(failed(grade(row, { kind: 'bold', address: 'C1:F1', bold: [[true, true, false, false]] })).detail).toBe('2 cells in 1:1 aren’t bold, starting at E1.');
    expect(grade(row, { kind: 'bold', address: '1:1', bold: [] }).status).toBe('pass');
  });
});

describe('filled', () => {
  const facts = (colors: (string | null)[]): SheetFacts => ({ kind: 'filled', address: 'A1:D1', colors: [colors] });

  it('counts white, blank and null as no fill', () => {
    expect(isNoFill('#FFFFFF')).toBe(true);
    expect(isNoFill('#ffffff')).toBe(true);
    expect(isNoFill('FFFFFF')).toBe(true);
    expect(isNoFill('')).toBe(true);
    expect(isNoFill(null)).toBe(true);
    expect(isNoFill('#FFF4CE')).toBe(false);
    expect(isNoFill('#000000')).toBe(false);
  });

  it('passes when every cell has a real fill', () => {
    expect(grade({ kind: 'filled', range: 'A1:D1', filled: true }, facts(['#D9E1F2', '#D9E1F2', '#FFF4CE', '#000000'])).status).toBe('pass');
    expect(grade({ kind: 'filled', range: 'A1:D1', filled: false }, facts(['#FFFFFF', '#ffffff', null, ''])).status).toBe('pass');
  });

  it('treats a white fill as missing', () => {
    const item = failed(grade({ kind: 'filled', range: 'A1:D1', filled: true }, facts(['#D9E1F2', '#ffffff', '#D9E1F2', '#D9E1F2'])));
    expect(item.detail).toBe('B1 has no fill color.');
    expect(item.focus).toBe('B1');
    expect(failed(grade({ kind: 'filled', range: 'A1:D1', filled: true }, facts([null, '', '#D9E1F2', '#FFFFFF']))).detail).toBe('3 cells in A1:D1 have no fill color, starting at A1.');
  });

  it('asks for No Fill when cells should be plain', () => {
    expect(failed(grade({ kind: 'filled', range: 'A1:D1', filled: false }, facts(['#FFFFFF', '#FFF4CE', null, '']))).detail).toBe('B1 has a fill color. Set it to No Fill.');
  });

  it('handles whole columns', () => {
    const item = failed(grade({ kind: 'filled', range: 'B:B', filled: false }, { kind: 'filled', address: 'B1:B3', colors: [[null], ['#FFFFFF'], ['#FFF4CE']] }));
    expect(item).toMatchObject({ detail: 'B3 has a fill color. Set it to No Fill.', focus: 'B3' });
  });
});

describe('hidden', () => {
  const cols: SheetCheck = { kind: 'hidden', columns: 'D:E', hidden: true };

  it('targets whole columns or rows', () => {
    expect(hiddenTarget({ kind: 'hidden', columns: '$d:$e', hidden: true })).toEqual({ columns: true, first: 'D', last: 'E', address: 'D:E' });
    expect(hiddenTarget({ kind: 'hidden', columns: 'D', hidden: true }).address).toBe('D:D');
    expect(hiddenTarget({ kind: 'hidden', rows: '5:9', hidden: false })).toEqual({ columns: false, first: '5', last: '9', address: '5:9' });
    expect(() => hiddenTarget({ kind: 'hidden', hidden: true })).toThrow();
  });

  it('passes when hidden or visible as asked', () => {
    expect(grade(cols, { kind: 'hidden', hidden: true }).status).toBe('pass');
    expect(grade({ kind: 'hidden', rows: '5:9', hidden: false }, { kind: 'hidden', hidden: false }).status).toBe('pass');
  });

  it('says which columns still show', () => {
    const item = failed(grade(cols, { kind: 'hidden', hidden: false }));
    expect(item.detail).toBe('Columns D–E are still showing. Select them, right-click, then Hide.');
    expect(item.focus).toBe('D1');
    expect(failed(grade({ kind: 'hidden', rows: '7', hidden: true }, { kind: 'hidden', hidden: false })).detail).toBe('Row 7 is still showing. Select it, right-click, then Hide.');
  });

  it('explains a mix of hidden and showing', () => {
    expect(failed(grade(cols, { kind: 'hidden', hidden: null })).detail).toBe('Only some of columns D–E are hidden. Select all of them, right-click, then Hide.');
    expect(failed(grade({ kind: 'hidden', rows: '5:9', hidden: false }, { kind: 'hidden', hidden: null })).detail).toBe(
      'Some of rows 5–9 are hidden. Select the rows on both sides, right-click, then Unhide.',
    );
  });

  it('asks to unhide when they should show', () => {
    const item = failed(grade({ kind: 'hidden', columns: 'D', hidden: false }, { kind: 'hidden', hidden: true }));
    expect(item.detail).toBe('Column D is hidden. Select the columns on both sides, right-click, then Unhide.');
    expect(item.focus).toBeUndefined();
  });

  it('unhides column A and row 1 from Select All, since nothing is on their other side', () => {
    expect(failed(grade({ kind: 'hidden', columns: 'A', hidden: false }, { kind: 'hidden', hidden: true })).detail).toBe(
      'Column A is hidden. Click the Select All button at the top-left corner of the sheet, then Home › Format › Hide & Unhide › Unhide Columns.',
    );
    expect(failed(grade({ kind: 'hidden', rows: '1:3', hidden: false }, { kind: 'hidden', hidden: null })).detail).toBe(
      'Some of rows 1–3 are hidden. Click the Select All button at the top-left corner of the sheet, then Home › Format › Hide & Unhide › Unhide Rows.',
    );
    // Row 10 starts with a 1 but has a row above it.
    expect(failed(grade({ kind: 'hidden', rows: '10', hidden: false }, { kind: 'hidden', hidden: true })).detail).toContain('Select the rows on both sides');
  });
});

describe('minWidth', () => {
  const check: SheetCheck = { kind: 'minWidth', range: 'B1:D1', points: 60 };

  it('passes wide columns, allowing a pixel of rounding', () => {
    expect(grade(check, { kind: 'minWidth', address: 'B1:D1', widths: [64, 60, 59.25] }).status).toBe('pass');
  });

  it('names the narrow column and how to fit it', () => {
    const item = failed(grade(check, { kind: 'minWidth', address: 'B1:D1', widths: [64, 38.25, 70] }));
    expect(item.detail).toBe('Column C is 38.25 points wide; it needs at least 60. Double-click the right edge of its column heading to fit the contents.');
    expect(item.focus).toBe('C1');
    expect(failed(grade(check, { kind: 'minWidth', address: 'B1:D1', widths: [48, 64, 30] })).detail).toMatch(/^2 columns are narrower than 60 points, starting with column B \(48\)\./);
  });

  it('calls out a hidden column', () => {
    expect(failed(grade(check, { kind: 'minWidth', address: 'B1:D1', widths: [64, 0, 70] })).detail).toBe('Column C is hidden. Unhide it and make it at least 60 points wide.');
  });

  it('handles whole columns', () => {
    const item = failed(grade({ kind: 'minWidth', range: 'B:D', points: 60 }, { kind: 'minWidth', address: 'B:D', widths: [64, 20, 70] }));
    expect(item.focus).toBe('C1');
  });
});

describe('conditionalFormat', () => {
  const check: SheetCheck = { kind: 'conditionalFormat', range: 'E2:E30' };
  const rules = (...appliesTo: string[][]): SheetFacts => ({ kind: 'conditionalFormat', count: appliesTo.length, appliesTo });

  it('passes a rule that covers the range', () => {
    expect(grade(check, rules(['E2:E30'])).status).toBe('pass');
    expect(grade(check, rules(['A1:H40'])).status).toBe('pass');
    expect(grade(check, rules(['$E:$E'])).status).toBe('pass');
    expect(grade(check, rules(['2:30'])).status).toBe('pass');
    // Excel splits a rule's Applies to into areas after rows are inserted or formats pasted.
    expect(grade(check, rules(['E2:E15', 'E16:E30'])).status).toBe('pass');
    expect(grade(check, rules(['E2'], ['E1:E31'])).status).toBe('pass');
  });

  it('checks coverage of a tall block without walking every row', () => {
    const tall: SheetCheck = { kind: 'conditionalFormat', range: 'E1:F1048576' };
    expect(grade(tall, rules(['E1:F500000', 'E400000:E1048576', 'F500001:F1048576'])).status).toBe('pass');
    expect(failed(grade(tall, rules(['E1:F10', 'E12:F1048576']))).detail).toContain('only part of E1:F1048576');
  });

  it('falls back to any rule when this Excel can’t list where rules apply', () => {
    expect(grade(check, { kind: 'conditionalFormat', count: 2 }).status).toBe('pass');
  });

  it('says where to add one', () => {
    const item = failed(grade(check, { kind: 'conditionalFormat', count: 0 }));
    expect(item.detail).toBe('There’s no conditional formatting on E2:E30 yet. Select it, then Home › Conditional Formatting.');
    expect(item.focus).toBe('E2');
  });

  it('fails a rule that touches only part of the range', () => {
    const item = failed(grade(check, rules(['E2'])));
    expect(item.detail).toBe(
      'The conditional formatting rule applies to E2, only part of E2:E30. Open Home › Conditional Formatting › Manage Rules and change Applies to so it covers all of E2:E30.',
    );
    expect(item.focus).toBe('E2');
    expect(failed(grade(check, rules(['E2:E15', 'E17:E30']))).detail).toContain('applies to E2:E15, E17:E30, only part of E2:E30.');
    expect(failed(grade(check, rules(['E2:E15'], ['E16:E30']))).detail).toMatch(/^None of the 2 conditional formatting rules here covers all of E2:E30\./);
    expect(failed(grade({ kind: 'conditionalFormat', range: 'D2:E3' }, rules(['D2:D3', 'E2']))).detail).toContain('only part of D2:E3');
  });

  it('counts any rule in whole columns or rows', () => {
    expect(grade({ kind: 'conditionalFormat', range: 'E:E' }, rules(['E2:E30'])).status).toBe('pass');
    expect(failed(grade({ kind: 'conditionalFormat', range: 'E:E' }, { kind: 'conditionalFormat', count: 0 })).focus).toBe('E1');
  });
});

describe('values', () => {
  const check: SheetCheck = { kind: 'values', range: 'B12:D12', expected: [['Total', 455, 1200.5]], describe: 'the totals row' };
  const facts = (values: Grid): SheetFacts => ({ kind: 'values', address: 'B12:D12', values });

  it('passes matching values, within float noise and with matchers', () => {
    expect(grade(check, facts([['Total', 455.00000000001, 1200.5]])).status).toBe('pass');
    const matcher: SheetCheck = { ...check, expected: [[{ match: /^total$/i, describe: 'a Total label' }, 455, 1200.5]] } as SheetCheck;
    expect(grade(matcher, facts([['TOTAL', 455, 1200.5]])).status).toBe('pass');
  });

  it('describes the first mismatch like the answer checks do', () => {
    const one = failed(grade(check, facts([['Total', 410, 1200.5]])));
    expect(one.detail).toBe('C12 shows 410; expected 455.');
    expect(one.focus).toBe('C12');
    expect(failed(grade(check, facts([['', 410, 1200.5]]))).detail).toBe('2 cells in the totals row don’t match. B12 shows a blank cell; expected “Total”.');
  });

  it('explains errors', () => {
    expect(failed(grade(check, facts([['Total', '#DIV/0!', 1200.5]]))).detail).toBe('C12 shows #DIV/0!. The formula divides by zero or by an empty cell.');
  });

  it('lines a whole-column read up with the expected values', () => {
    const column: SheetCheck = { kind: 'values', range: 'F:F', expected: [['Amount'], [10], [20]], describe: 'column F' };
    expect(grade(column, { kind: 'values', address: 'F1:F3', values: [['Amount'], [10], [20]] }).status).toBe('pass');
    // The used part starts at row 2, so the read starts there too.
    expect(grade({ ...column, expected: [[''], [10], [20]] } as SheetCheck, { kind: 'values', address: 'F2:F3', values: [[10], [20]] }).status).toBe('pass');
    const item = failed(grade(column, { kind: 'values', address: 'F2:F4', values: [[10], [25], ['extra']] }));
    expect(item.detail).toBe('3 cells in column F don’t match. F1 shows a blank cell; expected “Amount”.');
    expect(failed(grade(column, { kind: 'values', address: 'F:F', values: [] })).focus).toBe('F1');
    const row: SheetCheck = { kind: 'values', range: '5:5', expected: [['', '', 'Q1', 'Q2']], describe: 'row 5' };
    expect(grade(row, { kind: 'values', address: 'C5:D5', values: [['Q1', 'Q2']] }).status).toBe('pass');
    expect(failed(grade(row, { kind: 'values', address: 'C5:D5', values: [['Q1', 'Q3']] })).detail).toBe('D5 shows “Q3”; expected “Q2”.');
  });
});

describe('formulas', () => {
  const check: SheetCheck = { kind: 'formulas', range: 'F2:F4', formulas: true };
  const facts = (formulas: Grid): SheetFacts => ({ kind: 'formulas', address: 'F2:F4', formulas });

  it('passes when every cell has a formula', () => {
    expect(grade(check, facts([['=D2*E2'], ['=D3*E3'], ['=D4*E4']])).status).toBe('pass');
  });

  it('tells a typed value from an empty cell', () => {
    const typed = failed(grade(check, facts([['=D2*E2'], [125], ['=D4*E4']])));
    expect(typed.detail).toBe('F3 holds a typed value, not a formula.');
    expect(typed.focus).toBe('F3');
    expect(failed(grade(check, facts([['=D2*E2'], ['=D3*E3'], ['']]))).detail).toBe('F4 is empty; it needs a formula.');
    expect(failed(grade(check, facts([[1], [''], ['=D4*E4']]))).detail).toBe('2 cells in F2:F4 don’t have formulas, starting at F2.');
  });

  it('checks the formula against a pattern', () => {
    const subtotal: SheetCheck = { kind: 'formulas', range: 'F2:F4', formulas: true, pattern: /^=SUBTOTAL\(109,/i };
    expect(grade(subtotal, facts([['=SUBTOTAL(109,B2:B9)'], ['=subtotal(109,C2:C9)'], ['=SUBTOTAL(109,D2:D9)']])).status).toBe('pass');
    const item = failed(grade(subtotal, facts([['=SUBTOTAL(109,B2:B9)'], ['=SUM(C2:C9)'], ['=SUBTOTAL(109,D2:D9)']])));
    expect(item.detail).toBe('F3 has =SUM(C2:C9), which isn’t the formula this step asks for.');
    expect(item.focus).toBe('F3');
  });

  it('can require plain values instead', () => {
    const values: SheetCheck = { kind: 'formulas', range: 'F2:F4', formulas: false };
    expect(grade(values, facts([[1], ['x'], ['']])).status).toBe('pass');
    expect(failed(grade(values, facts([[1], ['=A3'], ['']]))).detail).toBe('F3 still holds a formula; it should hold a value.');
  });

  it('handles whole columns', () => {
    const column: SheetCheck = { kind: 'formulas', range: 'G:G', formulas: false };
    expect(failed(grade(column, { kind: 'formulas', address: 'G2:G4', formulas: [[1], [2], ['=SUM(G2:G3)']] })).detail).toBe('G4 still holds a formula; it should hold a value.');
    expect(failed(grade({ ...column, formulas: true } as SheetCheck, { kind: 'formulas', address: 'G1:G2', formulas: [['=A1'], [5]] })).focus).toBe('G2');
    expect(grade(column, { kind: 'formulas', address: 'G:G', formulas: [] }).status).toBe('pass');
  });
});

describe('advice and contract', () => {
  it('appends advice to a failing detail only', () => {
    const advice = 'Edit the macro so it finds the last row instead of stopping at row 20.';
    const item = failed(grade({ kind: 'bold', range: 'A1:B1', bold: true }, { kind: 'bold', address: 'A1:B1', bold: [[true, false]] }, advice));
    expect(item.detail).toBe(`B1 isn’t bold. ${advice}`);
    expect(grade({ kind: 'filter', on: true }, { kind: 'filter', on: true }, advice)).toEqual({ id: 'sheet', label: 'A sheet fact', status: 'pass' });
  });

  it('refuses facts read for a different check', () => {
    expect(() => grade({ kind: 'filter', on: true }, { kind: 'freeze', rows: 1, cols: 0 })).toThrow();
  });
});

// ---------- data validation ----------

const OPTIONS = ['Base', 'Upside', 'Downside'];
const dropdown = (cell = 'C2', options = OPTIONS): ValidationInspection => ({ kind: 'validationList', cell, options, label: 'A scenario dropdown' });
const list = (source: string, sourceValues?: Grid, inCellDropDown = true): ValidationFacts => ({ type: 'List', source, sourceValues, inCellDropDown });
const gradeList = (facts: ValidationFacts, insp = dropdown()) => gradeValidationList(insp, facts);

describe('list sources', () => {
  it('splits typed lists on commas or semicolons', () => {
    expect(parseListSource('Base,Upside,Downside')).toEqual({ kind: 'items', items: OPTIONS });
    expect(parseListSource(' Base ; Upside;Downside ')).toEqual({ kind: 'items', items: OPTIONS });
    expect(parseListSource('Base,,Upside,')).toEqual({ kind: 'items', items: ['Base', 'Upside'] });
  });

  it('reads references, names and spills', () => {
    expect(parseListSource('=$H$2:$H$4')).toEqual({ kind: 'range', sheet: undefined, address: 'H2:H4' });
    expect(parseListSource('=Inputs!$H$2:$H$4')).toEqual({ kind: 'range', sheet: 'Inputs', address: 'H2:H4' });
    expect(parseListSource("='My Sheet'!h2:h4")).toEqual({ kind: 'range', sheet: 'My Sheet', address: 'H2:H4' });
    expect(parseListSource("='Team''s inputs'!$A$1")).toEqual({ kind: 'range', sheet: "Team's inputs", address: 'A1' });
    expect(parseListSource('=$H:$H')).toEqual({ kind: 'range', sheet: undefined, address: 'H:H' });
    expect(parseListSource('=$H$2#')).toEqual({ kind: 'range', sheet: undefined, address: 'H2', spill: true });
    expect(parseListSource('=Scenarios')).toEqual({ kind: 'name', sheet: undefined, name: 'Scenarios' });
    expect(parseListSource('=Inputs!Scenarios')).toEqual({ kind: 'name', sheet: 'Inputs', name: 'Scenarios' });
  });

  it('leaves formulas alone', () => {
    expect(parseListSource('=INDIRECT("Inputs!H2:H4")')).toEqual({ kind: 'formula' });
    expect(parseListSource('=OFFSET($H$2,0,0,3,1)')).toEqual({ kind: 'formula' });
    expect(parseListSource('=Scenarios[Name]')).toEqual({ kind: 'formula' });
  });
});

describe('validationList', () => {
  it('passes a typed list, in any order, case or spacing', () => {
    expect(gradeList(list('Base,Upside,Downside')).status).toBe('pass');
    expect(gradeList(list('Downside;Base;Upside')).status).toBe('pass');
    expect(gradeList(list(' base ,UPSIDE,  downside ')).status).toBe('pass');
  });

  it('passes a referenced range, ignoring blank cells and spacing', () => {
    expect(gradeList(list('=$H$2:$H$6', [['Base'], [' Upside '], ['downside'], [''], [null]])).status).toBe('pass');
    expect(gradeList(list('=Scenarios', [['Base', 'Upside', 'Downside']])).status).toBe('pass');
  });

  it('compares numbers by their text', () => {
    expect(gradeList(list('=$H$2:$H$3', [[2025], [2026]]), dropdown('C2', ['2025', '2026'])).status).toBe('pass');
  });

  it('counts a repeated option once', () => {
    expect(gradeList(list('Base,Upside,Downside,base')).status).toBe('pass');
  });

  it('names a missing option', () => {
    const item = failed(gradeList(list('Base,Upside')));
    expect(item.detail).toBe('The list on C2 is missing “Downside”. Edit Source in Data › Data Validation.');
    expect(item.focus).toBe('C2');
    expect(failed(gradeList(list('Base'))).detail).toContain('is missing “Upside” and “Downside”.');
  });

  it('names an extra option, such as a header picked up with the range', () => {
    const item = failed(gradeList(list('=$H$1:$H$4', [['Scenario'], ['Base'], ['Upside'], ['Downside']])));
    expect(item.detail).toBe('The list on C2 has an extra option: “Scenario”. Its options come from H1:H4. Fix those cells, or point Source at the right ones.');
    expect(failed(gradeList(list("='My Sheet'!$H$1:$H$5", [['Scenario'], ['Base'], ['Upside'], ['Downside'], ['Stretch']]))).detail).toContain(
      'extra options: “Scenario” and “Stretch”. Its options come from \'My Sheet\'!H1:H5.',
    );
  });

  it('shows both lists when options are missing and extra', () => {
    expect(failed(gradeList(list('Base,Up,Down'))).detail).toBe(
      'The list on C2 offers “Base”, “Up” and “Down”; it should offer “Base”, “Upside” and “Downside”. Edit Source in Data › Data Validation.',
    );
  });

  it('flags an empty list', () => {
    expect(failed(gradeList(list('=Scenarios', [[''], ['']]))).detail).toBe(
      'The list on C2 is empty. Its options come from the named range Scenarios. Fix those cells, or point Source at the right ones.',
    );
    expect(failed(gradeList(list(''))).detail).toContain('has no source');
  });

  it('says what a reference points at that isn’t there', () => {
    const missing = (source: string, sourceMissing: ValidationFacts['sourceMissing']) => gradeList({ ...list(source), sourceMissing });
    const name = failed(missing('=Scenario', 'name'));
    expect(name.detail).toBe('The list on C2 uses =Scenario, but this workbook has no range named Scenario. Check the spelling, or point Source at the cells that hold the options.');
    expect(name.focus).toBe('C2');
    expect(failed(missing('=Inputs!Scenario', 'name')).detail).toContain('but Inputs has no range named Scenario.');
    expect(failed(missing("='Old inputs'!$H$2:$H$4", 'sheet')).detail).toBe(
      "The list on C2 uses ='Old inputs'!$H$2:$H$4, but there’s no sheet named 'Old inputs' in this workbook. Check the sheet name, or point Source at the cells that hold the options.",
    );
    expect(failed(missing('=$H$2#', 'spill')).detail).toBe(
      'The list on C2 uses =$H$2#, but H2 isn’t spilling anything. Enter a formula in H2 whose results spill, or point Source at the cells that hold the options.',
    );
    expect(failed(missing('=Inputs!$H$2#', 'spill')).detail).toContain('but Inputs!H2 isn’t spilling anything. Enter a formula in Inputs!H2');
  });

  it('asks for the cells or typed options when the source can’t be read', () => {
    const advice = 'Point Source at the cells that hold the options, or type them separated by commas.';
    expect(failed(gradeList(list('=INDIRECT("Inputs!H2:H4")'))).detail).toBe(`The list on C2 uses =INDIRECT("Inputs!H2:H4"), which the coach can’t read. ${advice}`);
    expect(failed(gradeList(list('=Missing'))).detail).toContain(advice);
  });

  it('requires the in-cell dropdown', () => {
    const item = failed(gradeList(list('Base,Upside,Downside', undefined, false)));
    expect(item.detail).toBe('The list on C2 has no dropdown arrow. Open Data › Data Validation and select In-cell dropdown.');
    expect(failed(gradeList(list('Base,Upside', undefined, false))).detail).toMatch(/Also select In-cell dropdown so the arrow shows\.$/);
  });

  it('explains cells without a list rule', () => {
    expect(failed(gradeList({ type: 'None', source: '' })).detail).toBe('There’s no dropdown on C2 yet. Select it, then Data › Data Validation, Allow: List.');
    expect(failed(gradeList({ type: 'WholeNumber', source: '' })).detail).toBe('The rule on C2 is set to Allow: Whole number. Change it to List in Data › Data Validation.');
    expect(failed(gradeList({ type: 'Inconsistent', source: '' })).detail).toContain('Clear All');
    expect(failed(gradeList({ type: 'MixedCriteria', source: '' })).detail).toMatch(/^Only part of C2 has a validation rule\./);
    expect(gradeList({ type: 'None', source: '' }).focus).toBe('C2');
  });
});

// ---------- host reads, against a stand-in for the Office.js object model ----------

interface FakeObject {
  isNullObject: boolean;
  load(): FakeObject;
  [key: string]: unknown;
}

const obj = (props: Record<string, unknown> = {}): FakeObject => ({ isNullObject: false, load() { return this; }, ...props });
const NULL_OBJECT = obj({ isNullObject: true });

/** Defined names, each with the displayed text of its range, or null for a name that isn't a range. */
function fakeNames(names: Record<string, Grid | null>) {
  return {
    getItemOrNullObject: (name: string) => (name in names ? obj({ getRangeOrNullObject: () => (names[name] ? obj({ text: names[name] }) : NULL_OBJECT) }) : NULL_OBJECT),
  };
}

function fakeSheet(name: string, ranges: Record<string, FakeObject>, extra: Record<string, unknown> = {}) {
  return obj({
    name,
    getRange: (address: string) => {
      if (!ranges[address]) throw new Error(`The test sheet has no range ${address}`);
      return ranges[address];
    },
    tables: obj({ items: [] }),
    names: fakeNames({}),
    ...extra,
  });
}

/** A request context whose syncs throw the queued errors in order. */
function fakeCtx(sheets: FakeObject[], names: Record<string, Grid | null> = {}, syncErrors: unknown[] = []) {
  const ctx = {
    workbook: { worksheets: { getItemOrNullObject: (n: string) => sheets.find((s) => s.name === n) ?? NULL_OBJECT }, names: fakeNames(names) },
    sync: async () => {
      const err = syncErrors.shift();
      if (err) throw err;
    },
  };
  return ctx as unknown as Excel.RequestContext;
}

const asSheet = (s: FakeObject) => s as unknown as Excel.Worksheet;

/** Pretends this Excel supports ExcelApi up to 1.<minor>. */
function excelApiUpTo(minor: number) {
  vi.stubGlobal('Office', { context: { requirements: { isSetSupported: (_set: string, version: string) => Number(version.split('.')[1]) <= minor } } });
}

describe('host: inspectSheet', () => {
  afterEach(() => vi.unstubAllGlobals());

  it('reads frozen panes and assigns the item id', async () => {
    const ws = fakeSheet('Coach-x', {}, { freezePanes: { getLocationOrNullObject: () => obj({ rowCount: 1, columnCount: 16384 }) } });
    const item = await inspectSheet(fakeCtx([ws]), asSheet(ws), sheet({ kind: 'freeze', rows: 1, cols: 0 }), 3);
    expect(item).toEqual({ id: 'sheet-3', label: 'A sheet fact', status: 'pass' });
    const none = fakeSheet('Coach-x', {}, { freezePanes: { getLocationOrNullObject: () => NULL_OBJECT } });
    failed(await inspectSheet(fakeCtx([none]), asSheet(none), sheet({ kind: 'freeze', rows: 1, cols: 0 }), 0));
  });

  it('skips when this Excel lacks the API, before or during the read', async () => {
    excelApiUpTo(6);
    const ws = fakeSheet('Coach-x', {}, { freezePanes: { getLocationOrNullObject: () => NULL_OBJECT } });
    expect(await inspectSheet(fakeCtx([ws]), asSheet(ws), sheet({ kind: 'freeze', rows: 1, cols: 0 }), 1)).toEqual({
      id: 'sheet-1',
      label: 'A sheet fact',
      status: 'skip',
      detail: 'This version of Excel can’t report frozen panes to add-ins.',
    });
    vi.unstubAllGlobals();
    const ctx = fakeCtx([ws], {}, [{ code: 'ApiNotFound' }]);
    expect((await inspectSheet(ctx, asSheet(ws), sheet({ kind: 'freeze', rows: 1, cols: 0 }), 1)).status).toBe('skip');
  });

  it('lets other Excel errors through', async () => {
    const ws = fakeSheet('Coach-x', {}, { freezePanes: { getLocationOrNullObject: () => NULL_OBJECT } });
    await expect(inspectSheet(fakeCtx([ws], {}, [{ code: 'GeneralException' }]), asSheet(ws), sheet({ kind: 'freeze', rows: 1, cols: 0 }), 0)).rejects.toEqual({ code: 'GeneralException' });
  });

  it('counts a Table’s filter buttons or the sheet’s AutoFilter', async () => {
    const check = sheet({ kind: 'filter', on: true });
    const table = fakeSheet('Coach-x', {}, { tables: obj({ items: [{ showFilterButton: false }, { showFilterButton: true }] }) });
    expect((await inspectSheet(fakeCtx([table]), asSheet(table), check, 0)).status).toBe('pass');
    const auto = fakeSheet('Coach-x', {}, { autoFilter: obj({ enabled: true }) });
    expect((await inspectSheet(fakeCtx([auto]), asSheet(auto), check, 0)).status).toBe('pass');
    const off = fakeSheet('Coach-x', {}, { autoFilter: obj({ enabled: false }) });
    failed(await inspectSheet(fakeCtx([off]), asSheet(off), check, 0));
    // The AutoFilter read fails on its own sync: nothing settles it, so skip.
    expect((await inspectSheet(fakeCtx([off], {}, [undefined, { code: 'GeneralException' }]), asSheet(off), check, 0)).status).toBe('skip');
  });

  it('settles “filter buttons off” from Tables alone when this Excel can’t read the AutoFilter', async () => {
    excelApiUpTo(8);
    const check = sheet({ kind: 'filter', on: false });
    const plain = fakeSheet('Coach-x', {}, { autoFilter: obj({ enabled: true }) });
    expect(await inspectSheet(fakeCtx([plain]), asSheet(plain), check, 4)).toEqual({
      id: 'sheet-4',
      label: 'A sheet fact',
      status: 'skip',
      detail: 'This version of Excel can’t report filter buttons to add-ins.',
    });
    const table = fakeSheet('Coach-x', {}, { tables: obj({ items: [{ showFilterButton: true }] }) });
    expect(failed(await inspectSheet(fakeCtx([table]), asSheet(table), check, 4)).detail).toBe('The filter buttons are still on. Turn them off with Data › Filter.');
  });

  it('maps cell properties to bold and fill', async () => {
    const cellProps = obj({
      address: "'Coach-x'!A1:B1",
      getCellProperties: () => ({ value: [[{ format: { font: { bold: true }, fill: { color: '#FFFFFF' } } }, { format: { font: { bold: false }, fill: { color: '#D9E1F2' } } }]] }),
    });
    const ws = fakeSheet('Coach-x', { 'A1:B1': cellProps });
    const bold = failed(await inspectSheet(fakeCtx([ws]), asSheet(ws), sheet({ kind: 'bold', range: 'A1:B1', bold: true }), 0));
    expect(bold).toMatchObject({ detail: 'B1 isn’t bold.', focus: 'B1' });
    const fill = failed(await inspectSheet(fakeCtx([ws]), asSheet(ws), sheet({ kind: 'filled', range: 'A1:B1', filled: true }), 0));
    expect(fill.detail).toBe('A1 has no fill color.');
  });

  it('reads column widths whether Excel returns a list or a grid', async () => {
    const props = [{ format: { columnWidth: 64 } }, { format: { columnWidth: 30 } }];
    const check = sheet({ kind: 'minWidth', range: 'B2:C9', points: 60 });
    for (const value of [props, [props, props]]) {
      const ws = fakeSheet('Coach-x', { 'B2:C9': obj({ address: 'Coach-x!B2:C9', getColumnProperties: () => ({ value }) }) });
      expect(failed(await inspectSheet(fakeCtx([ws]), asSheet(ws), check, 0)).focus).toBe('C2');
    }
  });

  it('reads hidden columns as whole columns', async () => {
    const ws = fakeSheet('Coach-x', { 'D:D': obj({ columnHidden: true }), '5:9': obj({ rowHidden: null }) });
    expect((await inspectSheet(fakeCtx([ws]), asSheet(ws), sheet({ kind: 'hidden', columns: 'D', hidden: true }), 0)).status).toBe('pass');
    expect(failed(await inspectSheet(fakeCtx([ws]), asSheet(ws), sheet({ kind: 'hidden', rows: '5:9', hidden: true }), 0)).detail).toContain('Only some of rows 5–9');
  });

  it('reads number formats, values and formulas', async () => {
    const ws = fakeSheet('Coach-x', {
      'D2:D3': obj({ address: 'Coach-x!D2:D3', numberFormat: [['$#,##0.00'], ['General']], values: [[5], [6]], formulas: [['=B2'], [6]] }),
    });
    const run = (check: SheetCheck) => inspectSheet(fakeCtx([ws]), asSheet(ws), sheet(check), 0);
    expect(failed(await run({ kind: 'numberFormat', range: 'D2:D3', matches: CURRENCY, describe: 'currency' })).focus).toBe('D3');
    expect((await run({ kind: 'values', range: 'D2:D3', expected: [[5], [6]], describe: 'the totals' })).status).toBe('pass');
    expect(failed(await run({ kind: 'formulas', range: 'D2:D3', formulas: true })).detail).toBe('D3 holds a typed value, not a formula.');
  });

  it('reads whole columns and rows from the part of the sheet that holds anything', async () => {
    const used = obj({ address: "'Coach-x'!A1:G4" });
    const cell = (bold: boolean, color: string) => ({ format: { font: { bold }, fill: { color } } });
    const amounts = obj({
      address: "'Coach-x'!F2:G4",
      numberFormat: [['$#,##0.00', '$#,##0.00'], ['$#,##0.00', '$#,##0.00'], ['$#,##0.00', '0.00']],
      values: [[10, 20], [30, 40], [50, 60]],
      formulas: [['=B2*2', 20], ['=B3*2', '=C3*2'], ['=B4*2', '=C4*2']],
    });
    const header = obj({ address: "'Coach-x'!A1:C1", getCellProperties: () => ({ value: [[cell(true, '#D9E1F2'), cell(true, '#D9E1F2'), cell(false, '#FFFFFF')]] }) });
    // Only the values-only used range clips: formatting alone doesn't widen the read.
    const clip = (part: FakeObject) => obj({ getIntersectionOrNullObject: (u: FakeObject) => (u === used ? part : NULL_OBJECT) });
    const ws = fakeSheet('Coach-x', { 'F:G': clip(amounts), '1:1': clip(header) }, { getUsedRange: (valuesOnly?: boolean) => (valuesOnly ? used : obj()) });
    const run = (check: SheetCheck) => inspectSheet(fakeCtx([ws]), asSheet(ws), sheet(check), 0);

    expect(failed(await run({ kind: 'numberFormat', range: 'F:G', matches: CURRENCY, describe: 'currency' }))).toMatchObject({ detail: 'G4 uses “0.00”; it should be currency.', focus: 'G4' });
    expect(failed(await run({ kind: 'bold', range: '1:1', bold: true })).detail).toBe('C1 isn’t bold.');
    expect(failed(await run({ kind: 'filled', range: '1:1', filled: true })).detail).toBe('C1 has no fill color.');
    expect((await run({ kind: 'values', range: 'F:G', expected: [['', ''], [10, 20], [30, 40], [50, 60]], describe: 'the amounts' })).status).toBe('pass');
    expect(failed(await run({ kind: 'values', range: 'F:G', expected: [['Amount', 'Tax'], [10, 20], [30, 40], [50, 60]], describe: 'the amounts' })).detail).toBe(
      '2 cells in the amounts don’t match. F1 shows a blank cell; expected “Amount”.',
    );
    expect(failed(await run({ kind: 'formulas', range: 'F:G', formulas: true })).detail).toBe('G2 holds a typed value, not a formula.');
  });

  it('treats whole columns with nothing used as empty', async () => {
    const ws = fakeSheet('Coach-x', { 'H:H': obj({ getIntersectionOrNullObject: () => NULL_OBJECT }) }, { getUsedRange: () => obj() });
    const run = (check: SheetCheck) => inspectSheet(fakeCtx([ws]), asSheet(ws), sheet(check), 0);
    expect((await run({ kind: 'numberFormat', range: 'H:H', matches: CURRENCY, describe: 'currency' })).status).toBe('pass');
    expect((await run({ kind: 'bold', range: 'H:H', bold: false })).status).toBe('pass');
    expect((await run({ kind: 'formulas', range: 'H:H', formulas: false })).status).toBe('pass');
    expect(failed(await run({ kind: 'values', range: 'H:H', expected: [['Total']], describe: 'the label' }))).toMatchObject({ detail: 'H1 shows a blank cell; expected “Total”.', focus: 'H1' });
  });

  it('checks that one conditional format covers the range', async () => {
    const rule = (...areas: string[]) => ({ getRanges: () => ({ areas: obj({ items: areas.map((address) => ({ address })) }) }) });
    const withRules = (...rules: unknown[]) => fakeSheet('Coach-x', { 'E2:E30': obj({ conditionalFormats: obj({ items: rules }) }) });
    const run = (ws: FakeObject, syncErrors: unknown[] = []) => inspectSheet(fakeCtx([ws], {}, syncErrors), asSheet(ws), sheet({ kind: 'conditionalFormat', range: 'E2:E30' }), 0);

    expect(failed(await run(withRules())).detail).toContain('There’s no conditional formatting on E2:E30 yet.');
    expect(failed(await run(withRules(rule("'Coach-x'!E2")))).detail).toMatch(/^The conditional formatting rule applies to E2, only part of E2:E30\./);
    expect((await run(withRules(rule("'Coach-x'!E2"), rule("'Coach-x'!E2:E15", "'Coach-x'!E16:E30")))).status).toBe('pass');
    // Can't list where the rules apply: any rule that touches the range counts.
    expect((await run(withRules(rule("'Coach-x'!E2")), [undefined, { code: 'GeneralException' }])).status).toBe('pass');
    excelApiUpTo(8);
    expect((await run(withRules(rule("'Coach-x'!E2")))).status).toBe('pass');
  });
});

describe('host: inspectValidation', () => {
  afterEach(() => vi.unstubAllGlobals());

  const cellWith = (type: string, source = '', inCellDropDown = true) => obj({ dataValidation: obj({ type, rule: type === 'List' ? { list: { source, inCellDropDown } } : {} }) });
  const run = (ws: FakeObject, sheets: FakeObject[] = [ws], names: Record<string, Grid | null> = {}, insp = dropdown()) => inspectValidation(fakeCtx(sheets, names), asSheet(ws), insp, 2);
  const SCENARIOS: Grid = [['Base'], ['Upside'], ['Downside']];

  it('grades a typed list and assigns the item id', async () => {
    const ws = fakeSheet('Coach-x', { C2: cellWith('List', 'Base,Upside,Downside') });
    expect(await run(ws)).toEqual({ id: 'validation-2', label: 'A scenario dropdown', status: 'pass' });
  });

  it('resolves a range on the practice sheet', async () => {
    const ws = fakeSheet('Coach-x', { C2: cellWith('List', '=$H$2:$H$5'), 'H2:H5': obj({ text: [['Base'], ['Upside'], ['Downside'], ['']] }) });
    expect((await run(ws)).status).toBe('pass');
  });

  it('compares options as the cells display them, like the dropdown', async () => {
    const ws = fakeSheet('Coach-x', { C2: cellWith('List', '=$H$2:$H$3'), 'H2:H3': obj({ values: [[0.05], [0.1]], text: [['5%'], ['10%']] }) });
    expect((await run(ws, [ws], {}, dropdown('C2', ['5%', '10%']))).status).toBe('pass');
  });

  it('resolves a range on another sheet, quoted or not', async () => {
    const inputs = fakeSheet("Team's inputs", { 'H2:H4': obj({ text: [['Base'], ['Upside'], ['Stretch']] }) });
    const ws = fakeSheet('Coach-x', { C2: cellWith('List', "='Team''s inputs'!$H$2:$H$4") });
    expect(failed(await run(ws, [ws, inputs])).detail).toContain('Its options come from \'Team\'\'s inputs\'!H2:H4.');
    const plain = fakeSheet('Inputs', { 'H2:H4': obj({ text: SCENARIOS }) });
    const ws2 = fakeSheet('Coach-x', { C2: cellWith('List', '=Inputs!$H$2:$H$4') });
    expect((await run(ws2, [ws2, plain])).status).toBe('pass');
  });

  it('resolves the sheet’s own names before the workbook’s, as Excel does', async () => {
    const ws = fakeSheet('Coach-x', { C2: cellWith('List', '=Scenarios') });
    expect((await run(ws, [ws], { Scenarios: SCENARIOS })).status).toBe('pass');
    const scoped = fakeSheet('Coach-x', { C2: cellWith('List', '=Scenarios') }, { names: fakeNames({ Scenarios: SCENARIOS }) });
    expect((await run(scoped)).status).toBe('pass');
    expect((await run(scoped, [scoped], { Scenarios: [['Old'], ['Stale']] })).status).toBe('pass');
    const stale = fakeSheet('Coach-x', { C2: cellWith('List', '=Scenarios') }, { names: fakeNames({ Scenarios: [['Old'], ['Stale']] }) });
    expect(failed(await run(stale, [stale], { Scenarios: SCENARIOS })).detail).toContain('“Old” and “Stale”');
  });

  it('reads a sheet-qualified name from that sheet only', async () => {
    const inputs = fakeSheet('Inputs', {}, { names: fakeNames({ Scenarios: SCENARIOS }) });
    const ws = fakeSheet('Coach-x', { C2: cellWith('List', '=Inputs!Scenarios') });
    expect((await run(ws, [ws, inputs], { Scenarios: [['Old']] })).status).toBe('pass');
    const other = fakeSheet('Coach-x', { C2: cellWith('List', '=Inputs!Cases') });
    expect(failed(await run(other, [other, inputs], { Cases: SCENARIOS })).detail).toBe(
      'The list on C2 uses =Inputs!Cases, but Inputs has no range named Cases. Check the spelling, or point Source at the cells that hold the options.',
    );
  });

  it('says when a name isn’t defined, and when it isn’t a range', async () => {
    const ws = fakeSheet('Coach-x', { C2: cellWith('List', '=Scenario') });
    expect(failed(await run(ws, [ws], { Scenarios: SCENARIOS })).detail).toContain('but this workbook has no range named Scenario.');
    // A name for a constant or a formula has no cells to read.
    expect(failed(await run(ws, [ws], { Scenario: null })).detail).toContain('which the coach can’t read');
  });

  it('reads a spill, and says when the anchor doesn’t spill', async () => {
    const spilled = obj({ text: SCENARIOS });
    const ws = fakeSheet('Coach-x', { C2: cellWith('List', '=$H$2#'), H2: obj({ getSpillingToRangeOrNullObject: () => spilled }) });
    expect((await run(ws)).status).toBe('pass');
    const flat = fakeSheet('Coach-x', { C2: cellWith('List', '=$H$2#'), H2: obj({ getSpillingToRangeOrNullObject: () => NULL_OBJECT }) });
    expect(failed(await run(flat)).detail).toContain('but H2 isn’t spilling anything.');
    excelApiUpTo(11);
    expect(failed(await run(ws)).detail).toBe(
      'The list on C2 uses =$H$2#, which the coach can’t read. Point Source at the cells that hold the options, or type them separated by commas.',
    );
  });

  it('reads the used part of a whole column or row', async () => {
    const column = fakeSheet(
      'Coach-x',
      { C2: cellWith('List', '=$H:$H'), 'H:H': obj({ getIntersectionOrNullObject: () => obj({ text: [['Scenario'], ['Base'], ['Upside'], ['Downside'], ['']] }) }) },
      { getUsedRange: () => obj() },
    );
    expect(failed(await run(column)).detail).toContain('extra option: “Scenario”');
    const row = fakeSheet(
      'Coach-x',
      { C2: cellWith('List', '=$3:$3'), '3:3': obj({ getIntersectionOrNullObject: () => obj({ text: [['', 'Base', 'Upside', 'Downside']] }) }) },
      { getUsedRange: () => obj() },
    );
    expect((await run(row)).status).toBe('pass');
    const unused = fakeSheet('Coach-x', { C2: cellWith('List', '=$J:$J'), 'J:J': obj({ getIntersectionOrNullObject: () => NULL_OBJECT }) }, { getUsedRange: () => obj() });
    expect(failed(await run(unused)).detail).toMatch(/^The list on C2 is empty\. Its options come from J:J\./);
  });

  it('says when a referenced sheet isn’t there, and can’t read a formula source', async () => {
    const ws = fakeSheet('Coach-x', { C2: cellWith('List', '=Gone!$H$2:$H$4') });
    expect(failed(await run(ws)).detail).toBe(
      'The list on C2 uses =Gone!$H$2:$H$4, but there’s no sheet named Gone in this workbook. Check the sheet name, or point Source at the cells that hold the options.',
    );
    const formula = fakeSheet('Coach-x', { C2: cellWith('List', '=OFFSET($H$2,0,0,3,1)') });
    expect(failed(await run(formula)).detail).toContain('which the coach can’t read');
  });

  it('reports cells without a list, and no dropdown arrow', async () => {
    expect(failed(await run(fakeSheet('Coach-x', { C2: cellWith('None') }))).detail).toBe('There’s no dropdown on C2 yet. Select it, then Data › Data Validation, Allow: List.');
    expect(failed(await run(fakeSheet('Coach-x', { C2: cellWith('List', 'Base,Upside,Downside', false) }))).detail).toContain('no dropdown arrow');
  });

  it('skips when this Excel can’t report data validation', async () => {
    excelApiUpTo(7);
    const ws = fakeSheet('Coach-x', { C2: cellWith('List', 'Base,Upside,Downside') });
    expect(await run(ws)).toEqual({ id: 'validation-2', label: 'A scenario dropdown', status: 'skip', detail: 'This version of Excel can’t report data validation to add-ins.' });
    vi.unstubAllGlobals();
    const ctx = fakeCtx([ws], {}, [{ code: 'ApiNotFound' }]);
    expect((await inspectValidation(ctx, asSheet(ws), dropdown(), 0)).status).toBe('skip');
  });
});

describe('copy', () => {
  it('follows the copy rules in every failing detail', () => {
    expect(details.length).toBeGreaterThan(40);
    for (const d of details) {
      expect(d, d).not.toMatch(/\b(please|simply|just|easy|easily|successfully|leverage|seamless)\b/i);
      expect(d, d).not.toMatch(/!(\s|$)/);
      expect(d, d).not.toContain("'t ");
      expect(d, d).toMatch(/[.)]$/);
    }
  });
});

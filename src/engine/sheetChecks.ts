import { cellAddress, colToNumber, numberToCol, parseRange, sheetRef, stripSheet, type CellRef, type RangeRef } from './address';
import { compareGrids, describeValue, explainError, isBlank, isErrorValue } from './compare';
import { isFormula } from './formula';
import type { CheckItem, Grid, Inspection, SheetCheck } from './types';

/**
 * Pure grading for the 'sheet' and 'validationList' inspections. The host (excel/sheetInspect.ts)
 * reads these facts from Excel; everything here is testable without Excel.
 */

/** What the host read for one SheetCheck. One shape per SheetCheck kind. */
export type SheetFacts =
  | { kind: 'freeze'; rows: number; cols: number }
  | { kind: 'filter'; on: boolean }
  /**
   * range.numberFormat for the checked range. In every grid kind, `address` is where the grid was
   * read: its top-left cell is grid[0][0]. For whole columns or rows the host reads only the part of
   * the sheet that holds anything, so the grid can start below or right of the check's range, or be empty.
   */
  | { kind: 'numberFormat'; address: string; formats: Grid }
  /** Per-cell bold, from range.getCellProperties. */
  | { kind: 'bold'; address: string; bold: boolean[][] }
  /** Per-cell fill color, from range.getCellProperties. null, '' or '#FFFFFF' mean no fill. */
  | { kind: 'filled'; address: string; colors: (string | null)[][] }
  /** range.columnHidden / rowHidden: null when some are hidden and some aren't. */
  | { kind: 'hidden'; hidden: boolean | null }
  /** Width in points of each column in the range, left to right. */
  | { kind: 'minWidth'; address: string; widths: number[] }
  /**
   * `count`: the rules that touch the range. `appliesTo`: each of those rules' Applies to areas, as
   * A1 ranges without the sheet. Absent when this Excel can't list them; then any rule counts.
   */
  | { kind: 'conditionalFormat'; count: number; appliesTo?: string[][] }
  | { kind: 'values'; address: string; values: Grid }
  | { kind: 'formulas'; address: string; formulas: Grid };

type SheetInspection = Extract<Inspection, { kind: 'sheet' }>;
type ValidationInspection = Extract<Inspection, { kind: 'validationList' }>;

/** Excel snaps column widths to whole pixels, so a width set by hand can read back up to a pixel (0.75 pt) short. */
const PIXEL_POINTS = 0.75;

const widthFormat = new Intl.NumberFormat('en-US', { maximumFractionDigits: 2 });

/** RegExp.test without the lastIndex a /g or /y pattern carries between calls. */
function test(pattern: RegExp, text: string): boolean {
  pattern.lastIndex = 0;
  return pattern.test(text);
}

const SHEET_ROWS = 1_048_576;
const SHEET_COLS = 16_384;

const WHOLE_LINES = /^(?:\$?[A-Z]{1,3}:\$?[A-Z]{1,3}|\$?\d+:\$?\d+)$/i;

/** True for whole columns ("F:G", "$H:$H") or whole rows ("1:1"), with or without a sheet. */
export function isWholeLines(address: string): boolean {
  return WHOLE_LINES.test(stripSheet(address).trim());
}

/** The cells a range spans, including whole columns ("B:D" is B1:D1048576) and whole rows ("5:9" is A5:XFD9). */
function bounds(address: string): RangeRef {
  const [a, b = a] = stripSheet(address).replace(/\$/g, '').trim().split(':');
  const span = (x: number, y: number) => [Math.min(x, y), Math.max(x, y)];
  if (/^[A-Z]{1,3}$/i.test(a) && /^[A-Z]{1,3}$/i.test(b)) {
    const [first, last] = span(colToNumber(a), colToNumber(b));
    return { start: { row: 1, col: first }, end: { row: SHEET_ROWS, col: last } };
  }
  if (/^\d+$/.test(a) && /^\d+$/.test(b)) {
    const [first, last] = span(Number(a), Number(b));
    return { start: { row: first, col: 1 }, end: { row: last, col: SHEET_COLS } };
  }
  return parseRange(address);
}

/** Top-left cell of a range, including whole columns ("B:D" is B1) and whole rows ("5:9" is A5). */
const topLeft = (address: string): CellRef => bounds(address).start;

/** The address `rows` down and `cols` right of `start`. */
const offset = (start: CellRef, rows: number, cols: number) => cellAddress({ row: start.row + rows, col: start.col + cols });

/** The cells in `grid` (read at `address`) where `bad` holds, in reading order. */
function offenders<T>(address: string, grid: T[][], bad: (value: T) => boolean): { address: string; value: T }[] {
  const start = topLeft(address);
  const out: { address: string; value: T }[] = [];
  grid.forEach((row, r) =>
    row.forEach((value, c) => {
      if (bad(value)) out.push({ address: offset(start, r, c), value });
    }),
  );
  return out;
}

/**
 * `grid`, read with its top-left at `at`, shifted so its top-left is `origin`: blank rows above and
 * blank cells to the left fill the gap. A whole-column read starts where the used part does.
 */
function alignGrid(grid: Grid, at: CellRef, origin: CellRef): Grid {
  const rows = Math.max(0, at.row - origin.row);
  const cols = Math.max(0, at.col - origin.col);
  if (!rows && !cols) return grid;
  const pad = Array.from({ length: cols }, () => '');
  return [...Array.from({ length: rows }, () => []), ...grid.map((row) => [...pad, ...row])];
}

/** True when the union of `areas` contains every cell of `target`. */
function covers(areas: readonly RangeRef[], target: RangeRef): boolean {
  // The areas crossing a row change only where one starts or ends, so one row per band is enough.
  const bands = new Set([target.start.row]);
  for (const a of areas) for (const row of [a.start.row, a.end.row + 1]) if (row > target.start.row && row <= target.end.row) bands.add(row);
  return [...bands].every((row) => {
    // Walk right across the areas on this row; a gap before the target's last column leaves it uncovered.
    const spans = areas.filter((a) => a.start.row <= row && row <= a.end.row).sort((x, y) => x.start.col - y.start.col);
    let col = target.start.col;
    for (const a of spans) {
      if (a.start.col > col) break;
      col = Math.max(col, a.end.col + 1);
    }
    return col > target.end.col;
  });
}

/** “A”, “B” and “C”. */
function quoteList(items: readonly string[]): string {
  const quoted = items.map((s) => `“${s}”`);
  return quoted.length > 1 ? `${quoted.slice(0, -1).join(', ')} and ${quoted[quoted.length - 1]}` : (quoted[0] ?? '');
}

const capitalize = (s: string) => s.charAt(0).toUpperCase() + s.slice(1);

// ---------- sheet checks ----------

/** No fill: Office.js reports an unfilled cell as #FFFFFF, and white looks the same as none. */
export function isNoFill(color: string | null | undefined): boolean {
  return !color || /^#?FFFFFF$/i.test(color.trim());
}

/** Rows and columns frozen, as the learner would say it: "row 1 and column A", "rows 1–2". */
function describeFrozen(rows: number, cols: number): { text: string; plural: boolean } {
  const parts: string[] = [];
  if (rows) parts.push(rows === 1 ? 'row 1' : `rows 1–${rows}`);
  if (cols) parts.push(cols === 1 ? 'column A' : `columns A–${numberToCol(cols)}`);
  return { text: parts.join(' and ') || 'nothing', plural: parts.length > 1 || rows > 1 || cols > 1 };
}

/** The whole columns or rows a 'hidden' check targets: "D:E" or "5:9" (a single "D" or "5" works too). */
export function hiddenTarget(check: Extract<SheetCheck, { kind: 'hidden' }>): { columns: boolean; first: string; last: string; address: string } {
  const spec = check.columns ?? check.rows;
  if (!spec) throw new Error('A hidden check needs columns or rows');
  const [first, last = first] = spec.replace(/\$/g, '').toUpperCase().split(':');
  return { columns: check.columns !== undefined, first, last, address: `${first}:${last}` };
}

/** Grades one 'sheet' inspection from the facts the host read. Item id: `sheet-<index>` is assigned by the host. */
export function gradeSheetCheck(insp: SheetInspection, facts: SheetFacts): CheckItem {
  const { check, label } = insp;
  if (check.kind !== facts.kind) throw new Error(`Facts for ${facts.kind} can’t grade a ${check.kind} check`);
  const pass: CheckItem = { id: 'sheet', label, status: 'pass' };
  const fail = (detail: string, focus?: string): CheckItem => ({ id: 'sheet', label, status: 'fail', detail: insp.advice ? `${detail} ${insp.advice}` : detail, focus });

  switch (check.kind) {
    case 'freeze': {
      const have = facts as Extract<SheetFacts, { kind: 'freeze' }>;
      if (have.rows === check.rows && have.cols === check.cols) return pass;
      const want = describeFrozen(check.rows, check.cols);
      const got = describeFrozen(have.rows, have.cols);
      // Freeze Panes freezes everything above and left of the selected cell.
      const cell = cellAddress({ row: check.rows + 1, col: check.cols + 1 });
      if (!have.rows && !have.cols) return fail(`Nothing is frozen yet. Select ${cell}, then View › Freeze Panes to keep ${want.text} in view.`, cell);
      if (!check.rows && !check.cols) return fail(`${capitalize(got.text)} ${got.plural ? 'are' : 'is'} frozen. Unfreeze from View › Freeze Panes.`);
      return fail(`${capitalize(got.text)} ${got.plural ? 'are' : 'is'} frozen instead of ${want.text}. Unfreeze from View › Freeze Panes, select ${cell}, then freeze again.`, cell);
    }

    case 'filter': {
      const { on } = facts as Extract<SheetFacts, { kind: 'filter' }>;
      if (on === check.on) return pass;
      return check.on ? fail('There are no filter buttons yet. Click in the data, then Data › Filter.') : fail('The filter buttons are still on. Turn them off with Data › Filter.');
    }

    case 'numberFormat': {
      const { address, formats } = facts as Extract<SheetFacts, { kind: 'numberFormat' }>;
      const bad = offenders(address, formats, (f) => !test(check.matches, String(f ?? '')));
      if (!bad.length) return pass;
      const format = String(bad[0].value ?? '') || 'General';
      const shown = format === 'General' ? 'General' : `“${format}”`;
      return bad.length === 1
        ? fail(`${bad[0].address} uses ${shown}; it should be ${check.describe}.`, bad[0].address)
        : fail(`${bad.length} cells in ${check.range} aren’t ${check.describe}, starting at ${bad[0].address} (${shown}).`, bad[0].address);
    }

    case 'bold': {
      const { address, bold } = facts as Extract<SheetFacts, { kind: 'bold' }>;
      const bad = offenders(address, bold, (b) => b !== check.bold);
      if (!bad.length) return pass;
      const first = bad[0].address;
      if (check.bold) return fail(bad.length === 1 ? `${first} isn’t bold.` : `${bad.length} cells in ${check.range} aren’t bold, starting at ${first}.`, first);
      return fail(bad.length === 1 ? `${first} is bold; it shouldn’t be.` : `${bad.length} cells in ${check.range} are bold, starting at ${first}. None should be.`, first);
    }

    case 'filled': {
      const { address, colors } = facts as Extract<SheetFacts, { kind: 'filled' }>;
      const bad = offenders(address, colors, (c) => isNoFill(c) === check.filled);
      if (!bad.length) return pass;
      const first = bad[0].address;
      if (check.filled) return fail(bad.length === 1 ? `${first} has no fill color.` : `${bad.length} cells in ${check.range} have no fill color, starting at ${first}.`, first);
      return fail(
        bad.length === 1 ? `${first} has a fill color. Set it to No Fill.` : `${bad.length} cells in ${check.range} have a fill color, starting at ${first}. Set them to No Fill.`,
        first,
      );
    }

    case 'hidden': {
      const { hidden } = facts as Extract<SheetFacts, { kind: 'hidden' }>;
      if (hidden === check.hidden) return pass;
      const t = hiddenTarget(check);
      const single = t.first === t.last;
      const noun = t.columns ? 'column' : 'row';
      const what = single ? `${noun} ${t.first}` : `${noun}s ${t.first}–${t.last}`;
      // Column A and row 1 have nothing on one side to select with them.
      const atEdge = t.first === (t.columns ? 'A' : '1');
      const unhide = atEdge
        ? `Click the Select All button at the top-left corner of the sheet, then Home › Format › Hide & Unhide › Unhide ${t.columns ? 'Columns' : 'Rows'}.`
        : `Select the ${noun}s on both sides, right-click, then Unhide.`;
      if (hidden === null) return check.hidden ? fail(`Only some of ${what} are hidden. Select all of them, right-click, then Hide.`) : fail(`Some of ${what} are hidden. ${unhide}`);
      if (!check.hidden) return fail(`${capitalize(what)} ${single ? 'is' : 'are'} hidden. ${unhide}`);
      const firstCell = t.columns ? `${t.first}1` : `A${t.first}`;
      return fail(`${capitalize(what)} ${single ? 'is' : 'are'} still showing. Select ${single ? 'it' : 'them'}, right-click, then Hide.`, firstCell);
    }

    case 'minWidth': {
      const { address, widths } = facts as Extract<SheetFacts, { kind: 'minWidth' }>;
      const start = topLeft(address);
      const narrow = widths.map((w, i) => ({ w, col: start.col + i })).filter((x) => x.w < check.points - PIXEL_POINTS);
      if (!narrow.length) return pass;
      const { w, col } = narrow[0];
      const letter = numberToCol(col);
      const focus = cellAddress({ row: start.row, col });
      if (w <= 0) return fail(`Column ${letter} is hidden. Unhide it and make it at least ${check.points} points wide.`, focus);
      return narrow.length === 1
        ? fail(`Column ${letter} is ${widthFormat.format(w)} points wide; it needs at least ${check.points}. Double-click the right edge of its column heading to fit the contents.`, focus)
        : fail(
            `${narrow.length} columns are narrower than ${check.points} points, starting with column ${letter} (${widthFormat.format(w)}). Double-click the right edge of each column heading to fit the contents.`,
            focus,
          );
    }

    case 'conditionalFormat': {
      const { count, appliesTo } = facts as Extract<SheetFacts, { kind: 'conditionalFormat' }>;
      const focus = cellAddress(topLeft(check.range));
      if (!count) return fail(`There’s no conditional formatting on ${check.range} yet. Select it, then Home › Conditional Formatting.`, focus);
      // A block needs one rule that covers all of it. For whole columns or rows, a rule anywhere in them counts.
      if (!appliesTo || isWholeLines(check.range)) return pass;
      const target = bounds(check.range);
      if (appliesTo.some((areas) => covers(areas.map(bounds), target))) return pass;
      const manage = `Open Home › Conditional Formatting › Manage Rules and change Applies to so it covers all of ${check.range}.`;
      return appliesTo.length === 1
        ? fail(`The conditional formatting rule applies to ${appliesTo[0].join(', ')}, only part of ${check.range}. ${manage}`, focus)
        : fail(`None of the ${appliesTo.length} conditional formatting rules here covers all of ${check.range}. ${manage}`, focus);
    }

    case 'values': {
      const { address, values } = facts as Extract<SheetFacts, { kind: 'values' }>;
      // `expected` starts at the range's top-left; a whole-column read may start lower down.
      const origin = topLeft(check.range);
      const mismatches = compareGrids(alignGrid(values, topLeft(address), origin), check.expected);
      if (!mismatches.length) return pass;
      const m = mismatches[0];
      const at = offset(origin, m.row, m.col);
      const text = isErrorValue(m.actual) ? `${at} shows ${m.actual}. ${explainError(m.actual)}` : `${at} shows ${describeValue(m.actual)}; expected ${describeValue(m.expected)}.`;
      return fail(mismatches.length === 1 ? text : `${mismatches.length} cells in ${check.describe} don’t match. ${text}`, at);
    }

    case 'formulas': {
      const { address, formulas } = facts as Extract<SheetFacts, { kind: 'formulas' }>;
      if (!check.formulas) {
        const bad = offenders(address, formulas, isFormula);
        if (!bad.length) return pass;
        const first = bad[0].address;
        return fail(
          bad.length === 1 ? `${first} still holds a formula; it should hold a value.` : `${bad.length} cells in ${check.range} still hold formulas, starting at ${first}. They should hold values.`,
          first,
        );
      }
      const missing = offenders(address, formulas, (f) => !isFormula(f));
      if (missing.length) {
        const first = missing[0];
        if (missing.length > 1) return fail(`${missing.length} cells in ${check.range} don’t have formulas, starting at ${first.address}.`, first.address);
        return fail(isBlank(first.value) ? `${first.address} is empty; it needs a formula.` : `${first.address} holds a typed value, not a formula.`, first.address);
      }
      const pattern = check.pattern;
      const off = pattern ? offenders(address, formulas, (f) => !test(pattern, String(f))) : [];
      if (!off.length) return pass;
      const first = off[0];
      return off.length === 1
        ? fail(`${first.address} has ${String(first.value)}, which isn’t the formula this step asks for.`, first.address)
        : fail(`${off.length} formulas in ${check.range} aren’t what this step asks for, starting at ${first.address} (${String(first.value)}).`, first.address);
    }
  }
}

// ---------- data-validation lists ----------

/** What the host read for a 'validationList' inspection. */
export interface ValidationFacts {
  /** dataValidation.type, e.g. "List", "None", "Inconsistent". */
  type: string;
  /** rule.list.source as Excel reports it: "Base,Upside,Downside", "=$H$2:$H$4", "=Scenarios". Empty when not a list. */
  source: string;
  /**
   * When `source` is a reference or a name, the cells it points at as they display (range.text), since
   * that's what the dropdown offers: 5% rather than 0.05. The host resolves it.
   */
  sourceValues?: Grid;
  /** Set when the reference points at something that isn't there: no such name, no such sheet, or an anchor that doesn't spill. */
  sourceMissing?: 'name' | 'sheet' | 'spill';
  /** rule.list.inCellDropDown. */
  inCellDropDown?: boolean;
}

/** A list rule's source, parsed. The host resolves 'range' and 'name' to cell values. */
export type ListSource =
  /** Options typed into the rule: "Base,Upside,Downside" (or ; between them). */
  | { kind: 'items'; items: string[] }
  /** A range, without $: "=$H$2:$H$4", "=Inputs!$H$2:$H$4", "='My Sheet'!H2:H4", "=$H:$H". With spill: the spill from an anchor, "=$H$2#". */
  | { kind: 'range'; sheet?: string; address: string; spill?: boolean }
  /** A defined name: "=Scenarios". */
  | { kind: 'name'; sheet?: string; name: string }
  /** Anything else, such as =INDIRECT(…) or =OFFSET(…). The coach doesn't evaluate these. */
  | { kind: 'formula' };

const SHEET_QUALIFIED = /^(?:'((?:[^']|'')+)'|([^\s'!"(),;:=#]+))!(.+)$/;
const A1_RANGE = /^\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?$/i;
const SPILL_REF = /^(\$?[A-Z]{1,3}\$?\d+)#$/i;
const DEFINED_NAME = /^[A-Z_\\][\w.\\]*$/i;

export function parseListSource(source: string): ListSource {
  const text = source.trim();
  if (!text.startsWith('=')) {
    return {
      kind: 'items',
      items: text
        .split(/[,;]/)
        .map((s) => s.trim())
        .filter(Boolean),
    };
  }
  let ref = text.slice(1).trim();
  let sheet: string | undefined;
  const q = SHEET_QUALIFIED.exec(ref);
  if (q) {
    sheet = q[1] !== undefined ? q[1].replace(/''/g, "'") : q[2];
    ref = q[3];
  }
  const spill = SPILL_REF.exec(ref);
  if (spill) return { kind: 'range', sheet, address: spill[1].replace(/\$/g, '').toUpperCase(), spill: true };
  if (A1_RANGE.test(ref) || WHOLE_LINES.test(ref)) return { kind: 'range', sheet, address: ref.replace(/\$/g, '').toUpperCase() };
  if (DEFINED_NAME.test(ref)) return { kind: 'name', sheet, name: ref };
  return { kind: 'formula' };
}

/** Where a referenced list reads its options from, as the learner would write it. */
function describeSource(src: Extract<ListSource, { kind: 'range' | 'name' }>): string {
  const prefix = src.sheet !== undefined ? `${sheetRef(src.sheet)}!` : '';
  return src.kind === 'name' ? `the named range ${prefix}${src.name}` : `${prefix}${src.address}${src.spill ? '#' : ''}`;
}

/** Excel's Allow: names for the other validation types. */
const ALLOW: Record<string, string> = {
  WholeNumber: 'Whole number',
  Decimal: 'Decimal',
  Date: 'Date',
  Time: 'Time',
  TextLength: 'Text length',
  Custom: 'Custom',
};

/**
 * Grades a dropdown: a List rule on `cell` offering exactly `options`, compared as a set ignoring
 * case and surrounding spaces. Blank cells in a referenced range don't count as options.
 */
export function gradeValidationList(insp: ValidationInspection, facts: ValidationFacts): CheckItem {
  const { cell, label } = insp;
  const fail = (detail: string): CheckItem => ({ id: 'validation', label, status: 'fail', detail, focus: cell });

  switch (facts.type) {
    case 'List':
      break;
    case 'None':
      return fail(`There’s no dropdown on ${cell} yet. Select it, then Data › Data Validation, Allow: List.`);
    case 'Inconsistent':
      return fail(`${cell} has conflicting validation rules. Select it, open Data › Data Validation, select Clear All, then set Allow: List.`);
    case 'MixedCriteria':
      return fail(`Only part of ${cell} has a validation rule. Select it, open Data › Data Validation, and select Yes when Excel offers to extend the rule.`);
    default:
      return fail(`The rule on ${cell} is set to Allow: ${ALLOW[facts.type] ?? facts.type}. Change it to List in Data › Data Validation.`);
  }

  if (!facts.source.trim()) return fail(`The list on ${cell} has no source. Open Data › Data Validation and type the options in Source, separated by commas.`);

  const src = parseListSource(facts.source);
  if (facts.sourceMissing && (src.kind === 'range' || src.kind === 'name')) {
    // The learner's reference is wrong, not the coach: say what isn't there.
    const uses = `The list on ${cell} uses ${facts.source.trim()}`;
    const pointAt = 'point Source at the cells that hold the options.';
    const sheet = src.sheet !== undefined ? sheetRef(src.sheet) : undefined;
    if (facts.sourceMissing === 'sheet' && sheet) return fail(`${uses}, but there’s no sheet named ${sheet} in this workbook. Check the sheet name, or ${pointAt}`);
    if (facts.sourceMissing === 'name' && src.kind === 'name') return fail(`${uses}, but ${sheet ?? 'this workbook'} has no range named ${src.name}. Check the spelling, or ${pointAt}`);
    if (facts.sourceMissing === 'spill' && src.kind === 'range') {
      const anchor = describeSource({ ...src, spill: false });
      return fail(`${uses}, but ${anchor} isn’t spilling anything. Enter a formula in ${anchor} whose results spill, or ${pointAt}`);
    }
  }

  let offered: string[];
  let fix: string;
  if (src.kind === 'items') {
    offered = src.items;
    fix = 'Edit Source in Data › Data Validation.';
  } else if (src.kind !== 'formula' && facts.sourceValues) {
    offered = facts.sourceValues
      .flat()
      .filter((v) => !isBlank(v))
      .map((v) => String(v).trim())
      .filter(Boolean);
    fix = `Its options come from ${describeSource(src)}. Fix those cells, or point Source at the right ones.`;
  } else {
    return fail(`The list on ${cell} uses ${facts.source.trim()}, which the coach can’t read. Point Source at the cells that hold the options, or type them separated by commas.`);
  }

  // Compare as sets: a repeated option still counts once.
  const key = (s: string) => s.trim().toLowerCase();
  const wanted = new Set(insp.options.map(key));
  const offeredKeys = new Set<string>();
  const have: string[] = [];
  for (const o of offered) {
    if (offeredKeys.has(key(o))) continue;
    offeredKeys.add(key(o));
    have.push(o);
  }
  const missing = insp.options.filter((o) => !offeredKeys.has(key(o)));
  const extra = have.filter((o) => !wanted.has(key(o)));
  const noArrow = facts.inCellDropDown === false;

  let mismatch: string | undefined;
  if (!have.length) mismatch = `The list on ${cell} is empty.`;
  else if (missing.length && extra.length) mismatch = `The list on ${cell} offers ${quoteList(have)}; it should offer ${quoteList(insp.options)}.`;
  else if (missing.length) mismatch = `The list on ${cell} is missing ${quoteList(missing)}.`;
  else if (extra.length) mismatch = `The list on ${cell} has ${extra.length === 1 ? 'an extra option' : 'extra options'}: ${quoteList(extra)}.`;

  if (mismatch) return fail(`${mismatch} ${fix}${noArrow ? ' Also select In-cell dropdown so the arrow shows.' : ''}`);
  if (noArrow) return fail(`The list on ${cell} has no dropdown arrow. Open Data › Data Validation and select In-cell dropdown.`);
  return { id: 'validation', label, status: 'pass' };
}

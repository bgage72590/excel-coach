import { cellAddress, colToNumber, parseRange, type CellRef } from './address';
import { isFormula, normalizeFormula, scanFormula } from './formula';
import type { Finding, FindingCell, ScanOptions, SheetFormulas } from './scan';

/**
 * Detectors behind "My Work". Everything is local and pure: the host hands over each sheet's
 * used range (A1 formulas plus R1C1), and this returns upgrade opportunities.
 *
 * Speed: formula text is analysed once per distinct R1C1 formula (a column filled down 5,000 rows
 * is one analysis), and the structural checks are a single pass over each column.
 */

type DetectorId =
  | 'vlookup-column-number'
  | 'vlookup-approximate'
  | 'iferror-lookup'
  | 'nested-if'
  | 'sumproduct-conditions'
  | 'concatenated-keys'
  | 'fixed-long-ranges'
  | 'inconsistent-column'
  | 'typed-over-formula'
  | 'volatile-references'
  | 'hardcoded-constants';

interface DetectorInfo {
  severity: Finding['severity'];
  title: string;
  why: string;
  fix: string;
  exerciseId?: string;
}

const DETECTORS: Record<DetectorId, DetectorInfo> = {
  'vlookup-column-number': {
    severity: 'medium',
    title: 'Lookups with a typed column number',
    why: 'VLOOKUP and HLOOKUP count columns by position, so inserting or deleting a column in the lookup range quietly returns the wrong field.',
    fix: 'Switch to XLOOKUP, which points at the return column itself and keeps working when columns move.',
    exerciseId: 'xlookup-lead-time',
  },
  'vlookup-approximate': {
    severity: 'high',
    title: 'Lookups that may match approximately',
    why: 'With the last argument left out or set to TRUE, VLOOKUP returns the closest smaller match instead of an error. That suits a sorted rate table, but on unsorted data it gives wrong answers with no warning.',
    fix: 'Add FALSE as the last argument for an exact match, or for tiered rates use XLOOKUP with match mode -1 so the intent is explicit.',
    exerciseId: 'xlookup-tiered',
  },
  'iferror-lookup': {
    severity: 'medium',
    title: 'IFERROR or IFNA wrapped around a lookup',
    why: 'The wrapper turns problems into the same fallback value, so a mistyped key or a number stored as text looks exactly like a genuinely missing item.',
    fix: 'Use XLOOKUP’s if_not_found argument so only a missing key gets the fallback and other errors still show.',
    exerciseId: 'xlookup-not-found',
  },
  'nested-if': {
    severity: 'medium',
    title: 'IF formulas nested three or more levels deep',
    why: 'Long IF chains are hard to read, and changing a threshold means finding and editing every copy of the formula.',
    fix: 'Move the tiers into a small lookup table and use XLOOKUP with match mode -1, or use IFS for a short list.',
    exerciseId: 'xlookup-tiered',
  },
  'sumproduct-conditions': {
    severity: 'low',
    title: 'SUMPRODUCT used as a conditional sum',
    why: 'Multiplying comparisons inside SUMPRODUCT works, but it’s harder to read and slower than the function built for the job.',
    fix: 'Use SUMIFS (or COUNTIFS) with one range-and-criteria pair per condition.',
    exerciseId: 'sumifs-warehouse',
  },
  'concatenated-keys': {
    severity: 'low',
    title: 'Helper columns that join fields into a key',
    why: 'Columns like =A2&B2 are usually built so a lookup can match on two fields; they widen the data and are one more thing to keep filled down.',
    fix: 'Match on both fields directly with XLOOKUP and multiplied conditions, then remove the helper column.',
    exerciseId: 'xlookup-two-keys',
  },
  'fixed-long-ranges': {
    severity: 'medium',
    title: 'Fixed ranges over long lists',
    why: 'References like A2:A500 stop at a fixed row, so rows added below the range are left out of totals and lookups.',
    fix: 'Convert the data to a Table and refer to its columns by name, so formulas grow with the data.',
    exerciseId: 'tables-convert',
  },
  'inconsistent-column': {
    severity: 'high',
    title: 'Formulas that break the pattern of their column',
    why: 'Most cells in these runs share one formula but a few differ, which usually means a fill was interrupted or a cell was edited by hand.',
    fix: 'Check the odd cells, then fill the correct formula down the whole run.',
    exerciseId: 'sumifs-grid',
  },
  'typed-over-formula': {
    severity: 'high',
    title: 'Typed values in a column of formulas',
    why: 'A number typed over a formula looks like a result but never updates, so the column goes stale without any sign.',
    fix: 'Restore the column’s formula in these cells, or move intended overrides to a separate, labeled input column.',
  },
  'volatile-references': {
    severity: 'medium',
    title: 'INDIRECT or OFFSET references',
    why: 'INDIRECT and OFFSET build references on the fly, so renaming or moving things breaks them without warning. They also recalculate on every edit, which slows large workbooks.',
    fix: 'Refer to Table columns, or use INDEX or XLOOKUP to pick the range you need.',
    exerciseId: 'tables-convert',
  },
  'hardcoded-constants': {
    severity: 'low',
    title: 'The same number typed into many formulas',
    why: 'The same number is typed into many formulas. When it changes, every copy has to be found and edited.',
    fix: 'Put the number in one labeled input cell (or a named cell) and point the formulas at it.',
  },
};

const DETECTOR_IDS = Object.keys(DETECTORS) as DetectorId[];
const SEVERITY_RANK: Record<Finding['severity'], number> = { high: 0, medium: 1, low: 2 };

const MAX_EXAMPLES = 8;
/** Rows past this many cells per sheet are not scanned (a guard against runaway used ranges). */
export const MAX_CELLS_PER_SHEET = 250_000;
/** Distinct formulas analysed per workbook; past this only the structural checks run. */
export const MAX_UNIQUE_FORMULAS = 25_000;
const LONG_RANGE_ROWS = 20;
const MIN_RUN = 3;
/** Distinct formulas (per sheet) one number has to be typed into to count as hard-coded. */
export const MIN_CONSTANT_FORMULAS = 3;
/** Cells one number has to reach, through a single filled formula or a few, to count as hard-coded. */
export const MIN_CONSTANT_CELLS = 5;
/** Numbers that are units rather than assumptions: 0 and 1, halving and quartering (2, 4), weeks, months, hours, minutes, days in a year, percent, thousands, millions. */
const TRIVIAL_NUMBERS = new Set([0, 1, 2, 4, 7, 12, 24, 52, 60, 100, 365, 365.25, 1000, 1440, 3600, 86400, 1_000_000]);

/**
 * Numbers that only count across different formulas. Repeated down one filled formula they're
 * usually structure, not an assumption: single digits (LEN(A2)-3, ROW()-5, MONTH(A2)/3,
 * WEEKDAY(A2,2)>5), halves and quarters (B2^0.5, INT(B2+0.5), TRUNC(B2/0.25)*0.25).
 */
const isFillStructure = (n: number): boolean => (Number.isInteger(n) && n < 10) || n === 0.5 || n === 0.25;

/**
 * Whether a number typed into `formulas` distinct formulas, filling `cells` cells in all, counts as
 * hard-coded. `n` comes from arithmeticConstants, so units are already left out. Exported so Fix
 * plans pick the same numbers the scan flags.
 */
export function isRepeatedConstant(n: number, formulas: number, cells: number): boolean {
  return formulas >= MIN_CONSTANT_FORMULAS || (cells >= MIN_CONSTANT_CELLS && !isFillStructure(n));
}

const LOOKUP_FUNCTIONS = new Set(['VLOOKUP', 'HLOOKUP', 'LOOKUP', 'XLOOKUP', 'MATCH', 'XMATCH']);
const ERROR_WRAPPERS = new Set(['IFERROR', 'IFNA']);

// ---------- Formula text helpers ----------

export interface Span {
  start: number;
  /** Exclusive. */
  end: number;
}

/**
 * Blanks out the inside of text literals, quoted sheet names and structured-reference brackets,
 * keeping the delimiters and the length, so positions line up with the original text.
 */
export function maskFormula(body: string): string {
  const n = body.length;
  let out = '';
  let i = 0;
  while (i < n) {
    const ch = body[i];
    if (ch === '"' || ch === "'") {
      out += ch;
      let j = i + 1;
      while (j < n) {
        if (body[j] === ch) {
          if (body[j + 1] === ch) {
            out += '  ';
            j += 2;
            continue;
          }
          break;
        }
        out += ' ';
        j++;
      }
      if (j < n) out += ch;
      i = j + 1;
    } else if (ch === '[') {
      out += '[';
      let depth = 1;
      let j = i + 1;
      while (j < n) {
        const c = body[j];
        if (c === "'") {
          // Escape character inside a structured reference: the next character is literal.
          out += j + 1 < n ? '  ' : ' ';
          j += 2;
          continue;
        }
        if (c === '[') depth++;
        else if (c === ']' && --depth === 0) break;
        out += ' ';
        j++;
      }
      if (j < n) out += ']';
      i = j + 1;
    } else {
      out += ch;
      i++;
    }
  }
  return out;
}

/**
 * Splits `text[from, to)` on `sep` wherever it sits outside parentheses, braces, brackets and
 * quotes. Works on raw or masked text.
 */
export function topLevelSpans(text: string, from: number, to: number, sep: string): Span[] {
  const spans: Span[] = [];
  let depth = 0;
  let start = from;
  for (let i = from; i < to; i++) {
    const ch = text[i];
    if (ch === '"' || ch === "'") {
      i++;
      while (i < to) {
        if (text[i] === ch) {
          if (text[i + 1] === ch) {
            i += 2;
            continue;
          }
          break;
        }
        i++;
      }
    } else if (ch === '(' || ch === '{' || ch === '[') depth++;
    else if (ch === ')' || ch === '}' || ch === ']') depth = Math.max(0, depth - 1);
    else if (depth === 0 && ch === sep) {
      spans.push({ start, end: i });
      start = i + 1;
    }
  }
  spans.push({ start, end: to });
  return spans;
}

/** Top-level comma split of an argument list, e.g. `"a,b",{1,2},MAX(1,2)` gives three parts. */
export function splitArgs(argsText: string): string[] {
  const code = maskFormula(argsText);
  return topLevelSpans(code, 0, code.length, ',').map((s) => argsText.slice(s.start, s.end).trim());
}

interface Call {
  name: string;
  /** Index of the "(". */
  open: number;
  /** Index of the matching ")" (or the text length when unclosed). */
  close: number;
  /** Index of the nearest enclosing call, or -1. */
  parent: number;
  args: Span[];
}

const isNameChar = (code: number): boolean =>
  (code >= 65 && code <= 90) || (code >= 97 && code <= 122) || (code >= 48 && code <= 57) || code === 95 || code === 46;
const isNameStart = (code: number): boolean => (code >= 65 && code <= 90) || (code >= 97 && code <= 122) || code === 95;

/** Finds function calls in masked formula text, outermost first. */
function parseCalls(code: string): Call[] {
  const calls: Call[] = [];
  const stack: number[] = [];
  for (let i = 0; i < code.length; i++) {
    const ch = code[i];
    if (ch === '(') {
      let k = i;
      while (k > 0 && isNameChar(code.charCodeAt(k - 1))) k--;
      if (k < i && isNameStart(code.charCodeAt(k))) {
        let parent = -1;
        for (let s = stack.length - 1; s >= 0; s--) {
          if (stack[s] >= 0) {
            parent = stack[s];
            break;
          }
        }
        calls.push({ name: code.slice(k, i).toUpperCase(), open: i, close: code.length, parent, args: [] });
        stack.push(calls.length - 1);
      } else {
        stack.push(-1);
      }
    } else if (ch === ')') {
      const top = stack.pop();
      if (top !== undefined && top >= 0) calls[top].close = i;
    }
  }
  for (const call of calls) {
    const spans = topLevelSpans(code, call.open + 1, call.close, ',');
    const empty = spans.length === 1 && code.slice(spans[0].start, spans[0].end).trim() === '';
    call.args = empty ? [] : spans;
  }
  return calls;
}

/** Index of the ")" matching the "(" at `open`, or -1. */
function matchingParen(code: string, open: number, end: number): number {
  let depth = 0;
  for (let i = open; i < end; i++) {
    const ch = code[i];
    if (ch === '"' || ch === "'") {
      i++;
      while (i < end && code[i] !== ch) i++;
    } else if (ch === '(') depth++;
    else if (ch === ')' && --depth === 0) return i;
  }
  return -1;
}

/** The single top-level comparison operator in `code[from, to)`: null when none, 'many' when several. */
function topLevelComparison(code: string, from: number, to: number): Span | 'many' | null {
  let found: Span | null = null;
  let depth = 0;
  for (let i = from; i < to; i++) {
    const ch = code[i];
    if (ch === '"' || ch === "'") {
      i++;
      while (i < to && code[i] !== ch) i++;
    } else if (ch === '(' || ch === '{' || ch === '[') depth++;
    else if (ch === ')' || ch === '}' || ch === ']') depth = Math.max(0, depth - 1);
    else if (depth === 0 && (ch === '<' || ch === '>' || ch === '=')) {
      if (found) return 'many';
      const next = code[i + 1];
      const two = (ch === '<' && (next === '>' || next === '=')) || (ch === '>' && next === '=');
      found = { start: i, end: two ? i + 2 : i + 1 };
      if (two) i++;
    }
  }
  return found;
}

const SHEET_PREFIX = /^(?:'[^']*'|[A-Za-z_][\w.]*)!/;
const CELL_REF = /^\$?[A-Za-z]{1,3}\$?\d+$/;
const A1_RANGE = /^\$?[A-Za-z]{1,3}\$?\d+:\$?[A-Za-z]{1,3}\$?\d+$/;
const COLUMN_RANGE = /^\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}$/;
const ROW_RANGE = /^\$?\d+:\$?\d+$/;
const MASKED_STRUCTURED = /^(?:[A-Za-z_\\][\w.]*)?\[ *\]$/;
const MASKED_STRING = /^" *"$/;
const NAME = /^[A-Za-z_\\][\w.]*$/;
const NUMBER_LITERAL = /^\d+(?:\.\d+)?(?:[eE][+-]?\d+)?%?$/;
const THIS_ROW = /@|#this row/i;

const stripSheetPrefix = (masked: string): string => masked.replace(SHEET_PREFIX, '');

function isCellRef(masked: string): boolean {
  return CELL_REF.test(stripSheetPrefix(masked));
}

/** A reference to one row of a Table, e.g. [@Region] or Sales[[#This Row],[Region]]. */
function isThisRowRef(masked: string, raw: string): boolean {
  return MASKED_STRUCTURED.test(masked) && THIS_ROW.test(raw);
}

/** A multi-cell reference: A2:A50, A:A, a Table column, or a defined name. */
function isRangeish(masked: string, raw: string): boolean {
  const m = stripSheetPrefix(masked.trim());
  if (A1_RANGE.test(m) || COLUMN_RANGE.test(m) || ROW_RANGE.test(m)) return true;
  if (MASKED_STRUCTURED.test(m)) return !THIS_ROW.test(raw);
  return NAME.test(m) && !CELL_REF.test(m) && !/^(TRUE|FALSE)$/i.test(m);
}

/** A single value: a literal, one cell, a this-row reference or an expression over those. */
function isScalarish(masked: string, raw: string): boolean {
  const m = masked.trim();
  if (!m || m.includes(':')) return false;
  return !m.includes('[') || THIS_ROW.test(raw);
}

/** SUMPRODUCT whose arguments are comparisons of a range against a value, times at most one range. */
function isConditionalSumproduct(code: string, body: string, call: Call): boolean {
  let comparisons = 0;
  let ranges = 0;
  const comparisonOk = (from: number, to: number, op: Span): boolean => {
    const lhs = [code.slice(from, op.start), body.slice(from, op.start)] as const;
    const rhs = [code.slice(op.end, to), body.slice(op.end, to)] as const;
    return (isRangeish(...lhs) && isScalarish(...rhs)) || (isRangeish(...rhs) && isScalarish(...lhs));
  };
  const visit = (from: number, to: number): boolean => {
    for (const factor of topLevelSpans(code, from, to, '*')) {
      let s = factor.start;
      let e = factor.end;
      while (s < e && (code[s] === ' ' || code[s] === '-' || code[s] === '+')) s++;
      while (e > s && code[e - 1] === ' ') e--;
      if (s >= e) return false;
      if (code[s] === '(' && matchingParen(code, s, e) === e - 1) {
        const inner = topLevelComparison(code, s + 1, e - 1);
        if (inner === 'many') return false;
        if (inner) {
          if (!comparisonOk(s + 1, e - 1, inner)) return false;
          comparisons++;
        } else if (!visit(s + 1, e - 1)) return false;
        continue;
      }
      const cmp = topLevelComparison(code, s, e);
      if (cmp === 'many') return false;
      if (cmp) {
        if (!comparisonOk(s, e, cmp)) return false;
        comparisons++;
        continue;
      }
      const masked = code.slice(s, e);
      if (isRangeish(masked, body.slice(s, e))) ranges++;
      else if (!NUMBER_LITERAL.test(masked) && !isCellRef(masked)) return false;
    }
    return true;
  };
  for (const arg of call.args) if (!visit(arg.start, arg.end)) return false;
  if (comparisons === 0 || ranges > 1) return false;
  // SUMIFS needs every range to be the same one-dimensional shape. A row of month headers compared
  // against a block of figures (a two-way lookup-and-sum) is beyond it.
  const shapes = new Set<string>();
  for (const m of code.slice(call.open, call.close).matchAll(A1_RANGE_SHAPE)) {
    const rows = Math.abs(Number(m[4]) - Number(m[2])) + 1;
    const cols = Math.abs(colToNumber(m[3]) - colToNumber(m[1])) + 1;
    if (rows > 1 && cols > 1) return false;
    shapes.add(`${rows}x${cols}`);
  }
  return shapes.size <= 1;
}

const A1_RANGE_SHAPE = /(?<![A-Za-z0-9_.$])\$?([A-Za-z]{1,3})\$?(\d+):\$?([A-Za-z]{1,3})\$?(\d+)(?![A-Za-z0-9_(!])/g;

/**
 * A key separator: nothing, or up to three symbols such as | - _ / : with no spaces or letters.
 * Text with spaces or words ("Smith, J", "Order 1042") is a display label, not a lookup key.
 */
const KEY_SEPARATOR = /^"[^\sA-Za-z0-9"]{0,3}"$/;

/** =A2&B2, =A2&"|"&B2, =CONCAT(A2,"-",B2) and friends: only cell refs and key separators, two or more refs. */
function isConcatenatedKey(code: string, body: string, calls: Call[]): boolean {
  const refsAndText = (spans: Span[]): boolean => {
    let refs = 0;
    for (const span of spans) {
      const masked = code.slice(span.start, span.end).trim();
      const raw = body.slice(span.start, span.end).trim();
      if (isCellRef(masked) || isThisRowRef(masked, raw)) refs++;
      else if (!MASKED_STRING.test(masked) || !KEY_SEPARATOR.test(raw)) return false;
    }
    return refs >= 2;
  };
  const parts = topLevelSpans(code, 0, code.length, '&');
  if (parts.length >= 2) return refsAndText(parts);
  // Otherwise the whole formula must be one CONCAT/CONCATENATE/TEXTJOIN call.
  const top = calls[0];
  if (!top || code.slice(0, top.open).trim().toUpperCase() !== top.name) return false;
  if (code.slice(top.close + 1).trim() !== '') return false;
  if (top.name === 'CONCAT' || top.name === 'CONCATENATE') return refsAndText(top.args);
  if (top.name === 'TEXTJOIN' && top.args.length >= 4) {
    const delimiter = code.slice(top.args[0].start, top.args[0].end).trim();
    const ignore = code.slice(top.args[1].start, top.args[1].end).trim();
    const rawDelimiter = body.slice(top.args[0].start, top.args[0].end).trim();
    if (!(MASKED_STRING.test(delimiter) && KEY_SEPARATOR.test(rawDelimiter)) && !isCellRef(delimiter)) return false;
    if (!/^(TRUE|FALSE|0|1|)$/i.test(ignore)) return false;
    return refsAndText(top.args.slice(2));
  }
  return false;
}

const A1_RANGE_IN_TEXT = /(?<![A-Za-z0-9_.$])\$?[A-Za-z]{1,3}\$?(\d+):\$?[A-Za-z]{1,3}\$?(\d+)(?![A-Za-z0-9_(!])/g;

function longestRangeRows(code: string): number {
  let longest = 0;
  for (const m of code.matchAll(A1_RANGE_IN_TEXT)) longest = Math.max(longest, Math.abs(Number(m[2]) - Number(m[1])) + 1);
  return longest;
}

/** True when the masked text holds a structured reference (not an external-workbook prefix). */
function hasStructuredRef(code: string): boolean {
  let i = code.indexOf('[');
  while (i >= 0) {
    const close = code.indexOf(']', i);
    if (close < 0) return true;
    if (!/^[A-Za-z0-9_.]*!/.test(code.slice(close + 1, close + 40))) return true;
    i = code.indexOf('[', close);
  }
  return false;
}

const NUMBER_IN_TEXT = /(?<![A-Za-z0-9_.$:])(\d+(?:\.\d+)?(?:[eE][+-]?\d+)?)(%?)(?![A-Za-z0-9_(.:])/g;
const ARITHMETIC = '*/+-^<>=';
const OPERAND_END = /[A-Za-z0-9_)\]}"'%$#.]/;

const roundKey = (n: number): number => Number(n.toPrecision(12));

/**
 * Numbers typed into arithmetic or comparisons (B2*1.08, C2>5000), as opposed to structural
 * arguments such as ROUND(x,2), LEFT(A2,3) or DATE(2026,1,1). Trivial values are left out.
 */
export function arithmeticConstants(code: string): number[] {
  const found = new Set<number>();
  for (const m of code.matchAll(NUMBER_IN_TEXT)) {
    const start = m.index;
    const end = start + m[0].length;
    let p = start - 1;
    while (p >= 0 && code[p] === ' ') p--;
    let before = p >= 0 ? code[p] : '';
    if (before === '-' || before === '+') {
      let q = p - 1;
      while (q >= 0 && code[q] === ' ') q--;
      // A sign after an operand is subtraction or addition; otherwise it's a unary sign.
      if (q < 0 || !OPERAND_END.test(code[q])) before = q >= 0 ? code[q] : '';
    }
    let a = end;
    while (a < code.length && code[a] === ' ') a++;
    const after = a < code.length ? code[a] : '';
    const inArithmetic = (before !== '' && ARITHMETIC.includes(before)) || (after !== '' && ARITHMETIC.includes(after));
    if (!inArithmetic) continue;
    const value = roundKey(Number(m[1]) / (m[2] ? 100 : 1));
    if (Number.isFinite(value) && !TRIVIAL_NUMBERS.has(value)) found.add(value);
  }
  return [...found];
}

// ---------- Per-formula analysis ----------

export interface FormulaFacts {
  vlookupColumnNumber: boolean;
  vlookupApproximate: boolean;
  iferrorLookup: boolean;
  /** Deepest chain of IFs inside IFs (1 = a single IF). */
  ifDepth: number;
  sumproductConditions: boolean;
  concatKey: boolean;
  /** Rows in the longest explicit A1 range, e.g. 499 for A2:A500. */
  longestRange: number;
  structured: boolean;
  volatile: boolean;
  constants: number[];
}

const TYPED_INDEX = /^\{?\s*\d[\d\s.,;]*\}?$/;

function isTruthyLiteral(text: string): boolean {
  if (/^TRUE$/i.test(text)) return true;
  return /^\d+(?:\.\d+)?$/.test(text) && Number(text) !== 0;
}

export function analyzeFormula(formula: string): FormulaFacts {
  const body = normalizeFormula(formula);
  const code = maskFormula(body);
  const calls = parseCalls(code);
  const argText = (span: Span): string => code.slice(span.start, span.end).trim();

  let vlookupColumnNumber = false;
  let vlookupApproximate = false;
  let iferrorLookup = false;
  let sumproductConditions = false;
  let ifDepth = 0;
  const ifDepths: number[] = new Array(calls.length).fill(0);

  calls.forEach((call, index) => {
    if (call.name === 'VLOOKUP' || call.name === 'HLOOKUP') {
      const args = call.args;
      if (args.length >= 3 && TYPED_INDEX.test(argText(args[2]))) vlookupColumnNumber = true;
      // An empty fourth argument (VLOOKUP(a,b,2,)) counts as FALSE, an exact match.
      if (args.length === 3 || (args.length >= 4 && isTruthyLiteral(argText(args[3])))) vlookupApproximate = true;
    } else if (call.name === 'IF') {
      let depth = 1;
      for (let p = call.parent; p >= 0; p = calls[p].parent) {
        if (calls[p].name === 'IF') {
          depth = ifDepths[p] + 1;
          break;
        }
      }
      ifDepths[index] = depth;
      ifDepth = Math.max(ifDepth, depth);
    } else if (call.name === 'SUMPRODUCT' && !sumproductConditions) {
      sumproductConditions = isConditionalSumproduct(code, body, call);
    }
    if (!iferrorLookup && LOOKUP_FUNCTIONS.has(call.name)) {
      for (let p = call.parent; p >= 0; p = calls[p].parent) {
        const wrapper = calls[p];
        const first = wrapper.args[0];
        if (ERROR_WRAPPERS.has(wrapper.name) && first && call.open >= first.start && call.open < first.end) {
          iferrorLookup = true;
          break;
        }
      }
    }
  });

  const functions = scanFormula(formula).functions;
  return {
    vlookupColumnNumber,
    vlookupApproximate,
    iferrorLookup,
    ifDepth,
    sumproductConditions,
    concatKey: isConcatenatedKey(code, body, calls),
    longestRange: longestRangeRows(code),
    structured: hasStructuredRef(code),
    volatile: functions.includes('INDIRECT') || functions.includes('OFFSET'),
    constants: arithmeticConstants(code),
  };
}

// ---------- Collecting hits ----------

class Hits {
  count = 0;
  private readonly groups = new Set<string>();
  private readonly distinct: FindingCell[] = [];
  private readonly others: FindingCell[] = [];

  /** `group` keeps the examples spread out: one per group first, then the rest in order. */
  add(sheet: string, row: number, col: number, group: string): void {
    this.count++;
    if (this.distinct.length < MAX_EXAMPLES && !this.groups.has(group)) {
      this.groups.add(group);
      this.distinct.push({ sheet, address: cellAddress({ row, col }) });
    } else if (this.others.length < MAX_EXAMPLES) {
      this.others.push({ sheet, address: cellAddress({ row, col }) });
    }
  }

  examples(): FindingCell[] {
    return [...this.distinct, ...this.others].slice(0, MAX_EXAMPLES);
  }
}

interface ConstantUse {
  sheet: string;
  row: number;
  col: number;
  /** Sheet and formula key: every cell filled from one formula shares it. */
  group: string;
  numbers: number[];
}

/** Where one typed number turns up: the distinct formulas (per sheet) and how many cells they fill. */
interface ConstantSpread {
  formulas: Set<string>;
  cells: number;
}

class Scan {
  readonly hits: Record<DetectorId, Hits>;
  private readonly cache = new Map<string, FormulaFacts | null>();
  private analysed = 0;
  readonly constantUses: ConstantUse[] = [];

  constructor() {
    this.hits = Object.fromEntries(DETECTOR_IDS.map((id) => [id, new Hits()])) as Record<DetectorId, Hits>;
  }

  facts(key: string, formula: string): FormulaFacts | null {
    let facts = this.cache.get(key);
    if (facts === undefined) {
      facts = this.analysed < MAX_UNIQUE_FORMULAS ? analyzeFormula(formula) : null;
      this.analysed++;
      this.cache.set(key, facts);
    }
    return facts;
  }

  cachedFacts(key: string): FormulaFacts | null {
    return this.cache.get(key) ?? null;
  }

  sheet(input: SheetFormulas): void {
    const grid = input.formulas ?? [];
    const r1c1 = input.r1c1 ?? [];
    let width = 0;
    for (const row of grid) width = Math.max(width, row?.length ?? 0);
    if (width === 0) return;
    const height = Math.min(grid.length, Math.max(1, Math.floor(MAX_CELLS_PER_SHEET / width)));
    let origin: CellRef = { row: 1, col: 1 };
    try {
      origin = parseRange(input.address || 'A1').start;
    } catch {
      // Unreadable address: report positions relative to A1.
    }
    const sheet = input.sheet;
    const keys: (string | null)[] = new Array(height * width).fill(null);
    const longRangeCells: number[] = [];
    let sheetHasStructured = false;

    for (let r = 0; r < height; r++) {
      const row = grid[r] ?? [];
      const rowR1C1 = r1c1[r] ?? [];
      for (let c = 0; c < width; c++) {
        const value = row[c];
        if (!isFormula(value)) continue;
        const alt = rowR1C1[c];
        // Cells filled from one formula share their R1C1 text, so they share one analysis.
        const key = isFormula(alt) ? `r${alt}` : `a${value}`;
        keys[r * width + c] = key;
        const facts = this.facts(key, value);
        if (!facts) continue;
        const absRow = origin.row + r;
        const absCol = origin.col + c;
        const group = `${sheet}\u0001${key}`;
        if (facts.vlookupColumnNumber) this.hits['vlookup-column-number'].add(sheet, absRow, absCol, group);
        if (facts.vlookupApproximate) this.hits['vlookup-approximate'].add(sheet, absRow, absCol, group);
        if (facts.iferrorLookup) this.hits['iferror-lookup'].add(sheet, absRow, absCol, group);
        if (facts.ifDepth >= 3) this.hits['nested-if'].add(sheet, absRow, absCol, group);
        if (facts.sumproductConditions) this.hits['sumproduct-conditions'].add(sheet, absRow, absCol, group);
        if (facts.volatile) this.hits['volatile-references'].add(sheet, absRow, absCol, group);
        if (facts.longestRange >= LONG_RANGE_ROWS) longRangeCells.push(r * width + c);
        if (facts.structured) sheetHasStructured = true;
        if (facts.constants.length) this.constantUses.push({ sheet, row: absRow, col: absCol, group, numbers: facts.constants });
      }
    }

    // Long fixed ranges only matter on sheets that don't use Tables at all.
    if (!sheetHasStructured) {
      for (const idx of longRangeCells) {
        const r = Math.floor(idx / width);
        const c = idx % width;
        this.hits['fixed-long-ranges'].add(sheet, origin.row + r, origin.col + c, `${sheet}\u0001${keys[idx]}`);
      }
    }

    // Column runs.
    const isFilled = (v: unknown): boolean => v !== '' && v !== null && v !== undefined;
    for (let c = 0; c < width; c++) {
      let formulaStart = -1;
      let filledStart = -1;
      for (let r = 0; r <= height; r++) {
        const inside = r < height;
        const key = inside ? keys[r * width + c] : null;
        if (key !== null) {
          if (formulaStart < 0) formulaStart = r;
        } else if (formulaStart >= 0) {
          if (r - formulaStart >= MIN_RUN) this.formulaRun(sheet, origin, keys, width, c, formulaStart, r);
          formulaStart = -1;
        }
        const filled = inside && isFilled(grid[r]?.[c]);
        if (filled) {
          if (filledStart < 0) filledStart = r;
        } else if (filledStart >= 0) {
          if (r - filledStart > MIN_RUN) this.filledRun(sheet, origin, grid, keys, width, c, filledStart, r);
          filledStart = -1;
        }
      }
    }
  }

  /** A vertical run of formula cells, rows [r0, r1). */
  private formulaRun(sheet: string, origin: CellRef, keys: (string | null)[], width: number, c: number, r0: number, r1: number): void {
    const keyAt = (r: number): string => keys[r * width + c]!;
    const len = r1 - r0;
    const counts = new Map<string, number>();
    let majority = '';
    let majorityCount = 0;
    for (let r = r0; r < r1; r++) {
      const k = keyAt(r);
      const n = (counts.get(k) ?? 0) + 1;
      counts.set(k, n);
      if (n > majorityCount) {
        majority = k;
        majorityCount = n;
      }
    }
    const group = `${sheet}\u0001${c}\u0001${r0}`;
    const absCol = origin.col + c;

    const majorityIsFill = majority.startsWith('r') && majorityCount >= 2 && majorityCount * 2 > len;
    const majorityFacts = majorityIsFill ? this.runFacts(majority, absCol) : null;
    // After a total or subtotal, the cells up to the next regular row are a summary block (tax, margin, net).
    let inSummary = false;

    // Segments of identical formulas.
    for (let s = r0; s < r1; ) {
      const k = keyAt(s);
      let e = s + 1;
      while (e < r1 && keyAt(e) === k) e++;
      const segment = e - s;

      if (segment >= MIN_RUN && this.cachedFacts(k)?.concatKey) {
        for (let r = s; r < e; r++) this.hits['concatenated-keys'].add(sheet, origin.row + r, absCol, `${group}\u0001${s}`);
      }

      if (k === majority) inSummary = false;
      // A different formula repeated over several rows is a deliberate section, not a slip.
      if (majorityFacts && k !== majority && k.startsWith('r') && segment < MIN_RUN) {
        const facts = this.runFacts(k, absCol);
        if (facts.ownRange || facts.summary) inSummary = true;
        for (let r = s; r < e; r++) {
          if (!inSummary && !excusedOutlier(facts, majorityFacts, r === r0, r === r1 - 1)) {
            this.hits['inconsistent-column'].add(sheet, origin.row + r, absCol, group);
          }
        }
      }
      s = e;
    }
  }

  private readonly runFactsCache = new Map<string, RunRefFacts>();

  private runFacts(key: string, col: number): RunRefFacts {
    const cacheKey = `${col}\u0001${key}`;
    let facts = this.runFactsCache.get(cacheKey);
    if (!facts) {
      facts = runRefFacts(key.slice(1), col);
      this.runFactsCache.set(cacheKey, facts);
    }
    return facts;
  }

  /** A vertical run of non-empty cells, rows [r0, r1): numbers typed among filled-down formulas. */
  private filledRun(
    sheet: string,
    origin: CellRef,
    grid: SheetFormulas['formulas'],
    keys: (string | null)[],
    width: number,
    c: number,
    r0: number,
    r1: number,
  ): void {
    const counts = new Map<string, number>();
    let majority = '';
    let majorityCount = 0;
    for (let r = r0; r < r1; r++) {
      const k = keys[r * width + c];
      if (k === null || !k.startsWith('r')) continue;
      const n = (counts.get(k) ?? 0) + 1;
      counts.set(k, n);
      if (n > majorityCount) {
        majority = k;
        majorityCount = n;
      }
    }
    if (majorityCount < MIN_RUN) return;
    let first = -1;
    let last = -1;
    for (let r = r0; r < r1; r++) {
      if (keys[r * width + c] === majority) {
        if (first < 0) first = r;
        last = r;
      }
    }
    if (majorityCount * 2 <= last - first + 1) return;
    const candidates: number[] = [];
    for (let r = first + 1; r < last; r++) candidates.push(r);
    // A number right after the last formula, ending the run, was most likely typed over it too.
    if (last + 1 === r1 - 1) candidates.push(last + 1);
    const group = `${sheet}\u0001${c}\u0001${r0}`;
    for (const r of candidates) {
      if (typeof grid[r]?.[c] === 'number') this.hits['typed-over-formula'].add(sheet, origin.row + r, origin.col + c, group);
    }
  }

  findings(): Finding[] {
    const findings: Finding[] = [];
    const constants = this.constants();
    for (const id of DETECTOR_IDS) {
      const hits = this.hits[id];
      if (hits.count === 0) continue;
      const info = DETECTORS[id];
      findings.push({
        id,
        severity: info.severity,
        title: info.title,
        why: id === 'hardcoded-constants' && constants ? constants.why : info.why,
        fix: info.fix,
        ...(info.exerciseId ? { exerciseId: info.exerciseId } : {}),
        count: hits.count,
        cells: hits.examples(),
        ...(id === 'hardcoded-constants' && constants ? { numbers: constants.numbers } : {}),
      });
    }
    return findings.sort(
      (a, b) => SEVERITY_RANK[a.severity] - SEVERITY_RANK[b.severity] || b.count - a.count || a.id.localeCompare(b.id),
    );
  }

  /**
   * Flags numbers typed into 3+ distinct formulas, or into 5+ cells (a rate filled down a column is
   * one formula but many copies); see isRepeatedConstant. Returns the "why" sentence naming them,
   * and the numbers themselves for the guided fix.
   */
  private constants(): { why: string; numbers: number[] } | null {
    const spread = new Map<number, ConstantSpread>();
    for (const use of this.constantUses) {
      for (const n of use.numbers) {
        let s = spread.get(n);
        if (!s) spread.set(n, (s = { formulas: new Set(), cells: 0 }));
        s.formulas.add(use.group);
        s.cells++;
      }
    }
    const flagged = [...spread]
      .filter(([n, s]) => isRepeatedConstant(n, s.formulas.size, s.cells))
      .sort((a, b) => b[1].formulas.size - a[1].formulas.size || b[1].cells - a[1].cells || a[0] - b[0]);
    if (!flagged.length) return null;
    const flaggedSet = new Set(flagged.map(([n]) => n));
    for (const use of this.constantUses) {
      if (use.numbers.some((n) => flaggedSet.has(n))) this.hits['hardcoded-constants'].add(use.sheet, use.row, use.col, use.group);
    }
    const numbers = flagged.map(([n]) => n);
    // Not "different formulas": the same formula on two sheets counts as two.
    if (flagged.length === 1) {
      const [n, { formulas, cells }] = flagged[0];
      const where =
        formulas.size === 1
          ? `${countText(cells)} cells`
          : cells === formulas.size
            ? `${countText(formulas.size)} formulas`
            : `${countText(cells)} cells across ${countText(formulas.size)} formulas`;
      return { why: `${formatNumber(n)} is typed into ${where}. When it changes, every copy has to be found and edited.`, numbers };
    }
    const shown = flagged.slice(0, 3).map(([n]) => formatNumber(n));
    const extra = flagged.length - shown.length;
    const list = extra > 0 ? `${shown.join(', ')} and ${countText(extra)} other number${extra === 1 ? '' : 's'}` : `${shown.slice(0, -1).join(', ')} and ${shown[shown.length - 1]}`;
    // Every flagged number is in at least three cells, but only some may be in three formulas.
    const where = flagged.every(([, s]) => s.formulas.size >= MIN_CONSTANT_FORMULAS)
      ? `${countText(MIN_CONSTANT_FORMULAS)} or more formulas`
      : 'several cells';
    return { why: `${list} are each typed into ${where}. When one changes, every copy has to be found and edited.`, numbers };
  }
}

function formatNumber(n: number): string {
  return String(roundKey(n));
}

const COUNT_WORDS = ['zero', 'one', 'two', 'three', 'four', 'five', 'six', 'seven', 'eight', 'nine'];
const countFormat = new Intl.NumberFormat('en-US');

/**
 * A count in running text: zero to nine in words, then numerals with thousands separators. The
 * words also keep a count like five from reading as a typed 5 when a Fix plan looks for the flagged
 * numbers in the "why" sentence (single digits are the numbers a long fill doesn't flag).
 */
const countText = (n: number): string => COUNT_WORDS[n] ?? countFormat.format(n);

// ---------- R1C1 reference facts (for column-run checks) ----------

// R1C1 offsets are written R[-1] and C[2]; maskR1C1 swaps those brackets for ‹ › so that masking
// (which blanks Table-reference brackets) leaves them intact.
const R1C1_PART = String.raw`(?:‹-?\d+›|\d+)`;
const R1C1_OFFSET_REF = /(?<![A-Za-z0-9_.])(?:R(?:\[-?\d+\]|\d+)?C(?:\[-?\d+\]|\d+)?|R\[-?\d+\]|C\[-?\d+\])(?![A-Za-z0-9_(])/g;

function maskR1C1(r1c1: string): string {
  const safe = normalizeFormula(r1c1).replace(R1C1_OFFSET_REF, (ref) => ref.replace(/\[/g, '‹').replace(/\]/g, '›'));
  return maskFormula(safe);
}
/** An R1C1 cell or range reference: R[-1]C, RC[2], R2C4:R40C4, R[-20]C:R[-1]C. */
const R1C1_REF = new RegExp(
  String.raw`(?<![A-Za-z0-9_.])R(${R1C1_PART})?C(${R1C1_PART})?(?::R(${R1C1_PART})?C(${R1C1_PART})?)?(?![A-Za-z0-9_(\[])`,
  'g',
);
const R1C1_WHOLE = new RegExp(String.raw`(?<![A-Za-z0-9_.])(?:R${R1C1_PART}(?::R${R1C1_PART})?|C${R1C1_PART}(?::C${R1C1_PART})?)(?![A-Za-z0-9_(\[])`, 'g');
const AGGREGATE_CALL = /(?<![A-Za-z0-9_.])(?:SUBTOTAL|AGGREGATE)\s*\(/i;

export interface RunRefFacts {
  /** Sums or otherwise spans a range of its own column (a total or subtotal). */
  ownRange: boolean;
  /** Refers to another cell of its own column on the same sheet. */
  ownCell: boolean;
  /** Refers to a row above, in any column of the same sheet. */
  rowAbove: boolean;
  /** Refers to a row below, in any column of the same sheet. */
  rowBelow: boolean;
  /** SUBTOTAL, AGGREGATE or a whole Table column (a totals-row style formula). */
  summary: boolean;
  /** Has any reference at all (cells, ranges, Table columns or names). */
  hasRef: boolean;
  /** Operator shape with every operand replaced by X, e.g. "X*X" or "SUMIFS(X,X,X)". */
  shape: string;
}

const rowOffset = (part: string | undefined): number | null => (part === undefined ? 0 : part.startsWith('‹') ? Number(part.slice(1, -1)) : null);
const isOwnCol = (part: string | undefined, col: number): boolean => part === undefined || part === '‹0›' || part === String(col);

/** Facts about an R1C1 formula as seen from column `col` (1-based, absolute). */
export function runRefFacts(r1c1: string, col: number): RunRefFacts {
  const code = maskR1C1(r1c1);
  let ownRange = false;
  let ownCell = false;
  let rowAbove = false;
  let rowBelow = false;
  let refs = 0;
  for (const m of code.matchAll(R1C1_REF)) {
    refs++;
    if (code[m.index - 1] === '!') continue; // Another sheet.
    const [, r1, c1, r2, c2] = m;
    const isRange = m[0].includes(':');
    for (const r of isRange ? [r1, r2] : [r1]) {
      const off = rowOffset(r);
      if (off !== null && off < 0) rowAbove = true;
      if (off !== null && off > 0) rowBelow = true;
    }
    if (isRange) {
      if (isOwnCol(c1, col) && isOwnCol(c2, col) && r1 !== r2) ownRange = true;
    } else if (isOwnCol(c1, col) && r1 !== undefined && r1 !== '‹0›') {
      ownCell = true;
    }
  }
  const wholeTableColumn = /\[/.test(code) && !THIS_ROW.test(r1c1);
  const summary = AGGREGATE_CALL.test(code) || wholeTableColumn;
  const shape = operatorShape(code);
  const hasRef = refs > 0 || R1C1_WHOLE.test(code) || /\[/.test(code) || hasName(code);
  R1C1_WHOLE.lastIndex = 0;
  return { ownRange, ownCell, rowAbove, rowBelow, summary, hasRef, shape };
}

/** Defined names left in masked R1C1 text once function names, refs and literals are gone. */
function hasName(code: string): boolean {
  const rest = code
    .replace(/"[^"]*"/g, ' ')
    .replace(/'[^']*'!|(?<![A-Za-z0-9_.])[A-Za-z_][\w.]*!/g, ' ')
    .replace(R1C1_REF, ' ')
    .replace(R1C1_WHOLE, ' ')
    .replace(/(?<![A-Za-z0-9_.])[A-Za-z_][\w.]*\s*\(/g, '(')
    .replace(/(?<![A-Za-z0-9_.])(?:TRUE|FALSE)(?![\w.])/gi, ' ');
  R1C1_REF.lastIndex = 0;
  R1C1_WHOLE.lastIndex = 0;
  return /(?<![0-9.])[A-Za-z_\\]/.test(rest.replace(/\d+(?:\.\d+)?(?:[eE][+-]?\d+)?/g, ' '));
}

/** Replaces every operand with X so two formulas can be compared by structure alone. */
function operatorShape(code: string): string {
  const out = code
    .replace(/'[^']*'!|(?<![A-Za-z0-9_.])[A-Za-z_][\w.]*!/g, '')
    .replace(/(?:[A-Za-z_\\][\w.]*)?\[[^\]]*\]/g, 'X')
    .replace(/"[^"]*"/g, 'X')
    .replace(R1C1_REF, 'X')
    .replace(R1C1_WHOLE, 'X')
    .replace(/(?<![A-Za-z0-9_.])\d+(?:\.\d+)?(?:[eE][+-]?\d+)?%?/g, 'X')
    .replace(/(?<![A-Za-z0-9_.])[A-Za-z_\\][\w.]*(?![\w.]*\s*\()/g, 'X')
    .replace(/\s+/g, '')
    .replace(/X(?::X)+/g, 'X')
    .toUpperCase();
  R1C1_REF.lastIndex = 0;
  R1C1_WHOLE.lastIndex = 0;
  return out;
}

/**
 * A cell that differs from its run looks like a broken fill only when it keeps the run's structure
 * (references drifted, or the run's formula was extended in place) or holds no references at all
 * (a hard-coded result typed as a formula). A structurally different formula is a deliberate line.
 */
function looksLikeBrokenFill(cell: RunRefFacts, majority: RunRefFacts): boolean {
  if (!cell.hasRef) return true;
  return cell.shape === majority.shape || cell.shape.includes(majority.shape);
}

/** Cells that legitimately differ from their run: totals, calculation rows, seeds and deliberate lines. */
function excusedOutlier(cell: RunRefFacts, majority: RunRefFacts, first: boolean, last: boolean): boolean {
  // Totals, subtotals and Table totals rows.
  if (cell.ownRange || cell.summary) return true;
  // A calculation row (gross profit = revenue - COGS) works from its own column; a fill that doesn't.
  if (cell.ownCell && !majority.ownCell) return true;
  // The first row of a fill that looks at the row above has no row above (a running balance seed,
  // a blank first % change); the last row of one that looks below has no row below.
  if (first && majority.rowAbove) return true;
  if (last && (majority.rowBelow || cell.ownCell)) return true;
  return !looksLikeBrokenFill(cell, majority);
}

// ---------- Entry point ----------

export function scanWorkbook(sheets: SheetFormulas[], options: ScanOptions = {}): Finding[] {
  const scan = new Scan();
  for (const sheet of sheets) {
    if (options.skipPrefix && sheet.sheet.startsWith(options.skipPrefix)) continue;
    scan.sheet(sheet);
  }
  return scan.findings();
}

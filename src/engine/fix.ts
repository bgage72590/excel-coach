import { cellAddress, colToNumber, numberToCol, parseCell, parseRange, rangeAddress, rangeSize, type CellRef } from './address';
import { cellMatches, compareGrids, describeValue, explainError, isBlank, isErrorValue, type Mismatch } from './compare';
import { isFormula, normalizeFormula } from './formula';
import type { Finding, SheetFormulas } from './scan';
import { analyzeFormula, arithmeticConstants, isRepeatedConstant, maskFormula, runRefFacts, topLevelSpans, type FormulaFacts, type RunRefFacts, type Span } from './scanDetectors';
import type { Cell, CheckItem, Grid } from './types';

/**
 * "Fix it in My Work": turns a scan finding into a guided fix done on a copy of the learner's
 * sheet, then compares the copy's results with the original's. Pure; the host copies and reads.
 *
 * A fix covers one block: the vertical run of formulas around the finding's first example cell.
 * The plan writes the exact formula to use where it can (the column's own VLOOKUP rewritten as
 * XLOOKUP, say), and the check proves the copy still shows the same results as the original.
 * Findings whose fix can legitimately change results (approximate lookups) get no plan.
 */

export interface FixRule {
  pattern: RegExp;
  /**
   * Decides the rule instead of `pattern`, for checks a regular expression can't express, such as
   * how deeply IFs nest. `pattern` is then a readable approximation. Without it, `pattern` is tested
   * against the formula with its text literals blanked, so "VLOOKUP(" inside quotes never counts.
   */
  test?(formula: string): boolean;
  /** Check label, e.g. "No VLOOKUP left in the column". */
  label: string;
  /** Shown when it fails. */
  advice: string;
}

/**
 * The column's own formula, which every repaired cell must hold after the fix. The other cells in
 * `valuesMayChange` (rows that build on the repaired ones) must keep the formulas they had.
 */
export interface FixRestore {
  /** The typed-over or odd cells the fix puts the column's formula back into. */
  cells: string[];
  /**
   * A cell that holds the column's formula, e.g. "D8". The check reads the formula there on the copy,
   * so a Table or sheet name Excel changed when it made the copy changes the expected formula too.
   */
  model: string;
  /** relativeKey() of the model's formula on the original, for when the copy's model isn't a formula. */
  key: string;
  /** Check label, e.g. "Typed values replaced with the column’s formula". */
  label: string;
  /** Shown when it fails. */
  advice: string;
}

export interface FixPlan {
  /** Detector id from the finding, e.g. "vlookup-column-number". */
  findingId: string;
  /** e.g. "Replace VLOOKUP with XLOOKUP". */
  title: string;
  /** One or two sentences: what changes and why it's safer. */
  intro: string;
  /** Guided steps. `backticks` render as cell references; addresses refer to the copy. */
  steps: string[];
  /** The learner's original sheet. */
  sheet: string;
  /** The block to fix, e.g. "E2:E240" (same address on the original and the copy). */
  range: string;
  /** Cells whose values are allowed to differ after the fix (e.g. stale typed-over values being restored). */
  valuesMayChange: string[];
  /** Every formula in `range` on the copy must avoid each of these... */
  forbid: FixRule[];
  /** ...and match each of these. */
  require: FixRule[];
  /** Every cell in `range` must hold a formula after the fix. */
  allFormulas: boolean;
  /** Typed-over and inconsistent-column fixes: the formula the repaired cells must hold. */
  restore?: FixRestore;
  /**
   * Replaces the usual "copy the new formulas back" next step when something else has to come
   * along: an input cell, or a Table the formulas refer to. `backticks` render as cell references.
   */
  copyBack?: string;
  /**
   * Set when the column runs on past `range` (a block cut at MAX_FIX_ROWS, or a sheet read in part):
   * one or two sentences for the end of the copy-back step. The fill step already says it.
   */
  continues?: string;
  /** IFERROR fixes: what the old formula showed in place of an error, when its fallback is a plain value. */
  fallback?: Cell;
}

export interface FixOptions {
  /**
   * The sheet goes on below the rows that were scanned (a very large sheet read in part), so a
   * block that reaches the last scanned row may continue. See scanEdgeBelow.
   */
  partial?: boolean;
}

/** One read of a range: values and formulas, top-left-anchored at `address` (no sheet name). */
export interface RangeRead {
  address: string;
  values: Grid;
  formulas: Grid;
}

/** Rows in one fix at most; a longer run is fixed in blocks. */
export const MAX_FIX_ROWS = 5000;
/** As in the scanner: an explicit range this many rows long is a long list. */
const LONG_RANGE_ROWS = 20;
/** As in the scanner: a different formula repeated over this many rows is a deliberate section. */
const MIN_SECTION = 3;

// ---------- Formula text ----------

/**
 * The formula with every A1 reference written as an offset from `cell`, spaces dropped and case
 * folded outside text literals. Copies of one filled formula share a key, like their R1C1 text.
 */
export function relativeKey(formula: string, cell: string): string {
  const { row, col } = parseCell(cell);
  const body = normalizeFormula(formula);
  const code = maskFormula(body);
  const r = (abs: string, n: string) => (abs ? `R${n}` : `R[${Number(n) - row}]`);
  const c = (abs: string, letters: string) => (abs ? `C${colToNumber(letters)}` : `C[${colToNumber(letters) - col}]`);
  let out = '';
  let last = 0;
  for (const m of code.matchAll(A1_REF)) {
    const [, ca, cl, ra, rn, c1a, c1, c2a, c2, r1a, r1, r2a, r2] = m;
    const ref = cl !== undefined ? r(ra, rn) + c(ca, cl) : c1 !== undefined ? `${c(c1a, c1)}:${c(c2a, c2)}` : `${r(r1a, r1)}:${r(r2a, r2)}`;
    out += body.slice(last, m.index) + ref;
    last = m.index + m[0].length;
  }
  return squash(out + body.slice(last));
}

/** One cell, one whole-column range or one whole-row range, in masked formula text. */
const A1_REF =
  /(?<![A-Za-z0-9_.$\]])(?:(\$?)([A-Za-z]{1,3})(\$?)(\d+)|(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})|(\$?)(\d+):(\$?)(\d+))(?![A-Za-z0-9_(])/g;

function squash(text: string): string {
  let out = '';
  let quoted = false;
  for (const ch of text) {
    if (ch === '"') quoted = !quoted;
    if (quoted || ch === '"') out += ch;
    else if (!/\s/.test(ch)) out += ch.toUpperCase();
  }
  return out;
}

interface Call {
  /** Upper-cased function name. */
  name: string;
  /** Index of the name's first character. */
  start: number;
  /** Index of the "(". */
  open: number;
  /** Index of the matching ")" (the text length when unclosed). */
  close: number;
  args: Span[];
}

const CALL = /(?<![A-Za-z0-9_.])([A-Za-z_][A-Za-z0-9_.]*)\s*\(/g;

/** Function calls in masked formula text, in order of their names: outer calls before inner ones. */
function findCalls(code: string): Call[] {
  const calls: Call[] = [];
  for (const m of code.matchAll(CALL)) {
    const open = m.index + m[0].length - 1;
    const close = closingParen(code, open);
    const spans = topLevelSpans(code, open + 1, close, ',');
    const empty = spans.length === 1 && code.slice(spans[0].start, spans[0].end).trim() === '';
    calls.push({ name: m[1].toUpperCase(), start: m.index, open, close, args: empty ? [] : spans });
  }
  return calls;
}

/** Index of the ")" matching the "(" at `open` in masked text (where quotes hold no parentheses). */
function closingParen(code: string, open: number): number {
  let depth = 0;
  for (let i = open; i < code.length; i++) {
    if (code[i] === '(') depth++;
    else if (code[i] === ')' && --depth === 0) return i;
  }
  return code.length;
}

const textOf = (body: string, span: Span): string => body.slice(span.start, span.end).trim();

/** The call that makes up the whole of `span`, if one does. */
function soleCall(calls: Call[], code: string, span: Span): Call | undefined {
  return calls.find((c) => c.start >= span.start && c.close < span.end && code.slice(span.start, c.start).trim() === '' && code.slice(c.close + 1, span.end).trim() === '');
}

type Rewrite = (call: Call, body: string, code: string) => string | undefined | null;

/**
 * Rewrites the calls named in `names`, outermost first. `rewrite` returns the replacement text,
 * undefined to keep the call (its arguments are still visited), or null when the formula can't be
 * rewritten exactly. Returns the new formula with its "=", or null.
 */
function rewriteCalls(formula: string, names: ReadonlySet<string>, rewrite: Rewrite): string | null {
  let body = normalizeFormula(formula);
  let from = 0;
  for (let guard = 0; guard < 200; guard++) {
    const code = maskFormula(body);
    const call = findCalls(code).find((c) => c.start >= from && names.has(c.name));
    if (!call) return `=${body}`;
    const next = rewrite(call, body, code);
    if (next === null) return null;
    if (next !== undefined) body = body.slice(0, call.start) + next + body.slice(call.close + 1);
    from = call.start + 1;
  }
  return null;
}

// ---------- A1 areas ----------

interface Edge {
  /** "$" or "". */
  abs: string;
  n: number;
}

/** An A1 range as typed: both corners, or whole columns (no rows), or whole rows (no cols). */
interface Area {
  /** Sheet prefix with its "!", or "". */
  prefix: string;
  cols?: [Edge, Edge];
  rows?: [Edge, Edge];
}

const SHEET_PREFIX = String.raw`(?:'(?:[^']|'')+'|[A-Za-z_][\w.]*)!`;
const AREA_TEXT = new RegExp(
  String.raw`^(${SHEET_PREFIX})?(?:(\$?)([A-Za-z]{1,3})(\$?)(\d+):(\$?)([A-Za-z]{1,3})(\$?)(\d+)|(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})|(\$?)(\d+):(\$?)(\d+))$`,
);
const CELL_TEXT = new RegExp(String.raw`^(${SHEET_PREFIX})?(\$?)([A-Za-z]{1,3})(\$?)(\d+)$`);

const edge = (abs: string | undefined, n: number): Edge => ({ abs: abs ?? '', n });

function parseArea(raw: string): Area | null {
  const m = AREA_TEXT.exec(raw.trim());
  if (!m) return null;
  const prefix = m[1] ?? '';
  if (m[3] !== undefined) {
    return {
      prefix,
      cols: [edge(m[2], colToNumber(m[3])), edge(m[6], colToNumber(m[7]))],
      rows: [edge(m[4], Number(m[5])), edge(m[8], Number(m[9]))],
    };
  }
  if (m[11] !== undefined) return { prefix, cols: [edge(m[10], colToNumber(m[11])), edge(m[12], colToNumber(m[13]))] };
  return { prefix, rows: [edge(m[14], Number(m[15])), edge(m[16], Number(m[17]))] };
}

function areaText(a: Area): string {
  const col = (e: Edge) => e.abs + numberToCol(e.n);
  const row = (e: Edge) => e.abs + e.n;
  if (a.cols && a.rows) return `${a.prefix}${col(a.cols[0])}${row(a.rows[0])}:${col(a.cols[1])}${row(a.rows[1])}`;
  if (a.cols) return `${a.prefix}${col(a.cols[0])}:${col(a.cols[1])}`;
  return `${a.prefix}${row(a.rows![0])}:${row(a.rows![1])}`;
}

/** The sheet a prefix such as `'Q1 Ops'!` names. */
const prefixSheet = (prefix: string): string => prefix.slice(0, -1).replace(/^'(.*)'$/, '$1').replace(/''/g, "'");

// ---------- Rewrites, one per kind of fix ----------

const LOOKUP_CALLS = new Set(['VLOOKUP', 'HLOOKUP']);
const ANY_LOOKUP = new Set(['VLOOKUP', 'HLOOKUP', 'LOOKUP', 'XLOOKUP', 'MATCH', 'XMATCH']);
const WRAPPERS = new Set(['IFERROR', 'IFNA']);

/** The XLOOKUP for an exact-match VLOOKUP or HLOOKUP with a typed index into a plain range, or null. */
function xlookupFor(call: Call, body: string, fallback?: string): string | null {
  if (!LOOKUP_CALLS.has(call.name)) return null;
  const args = call.args.map((s) => textOf(body, s));
  if (args.length !== 4 || !/^(FALSE|0|)$/i.test(args[3]) || !/^\d+$/.test(args[2])) return null;
  const index = Number(args[2]);
  const area = parseArea(args[1]);
  const vertical = call.name === 'VLOOKUP';
  const span = vertical ? area?.cols : area?.rows;
  if (!area || !span || index < 1 || index > span[1].n - span[0].n + 1) return null;
  // The first column (or row) is where the key is found; the index counts from it.
  const line = (e: Edge): Area => (vertical ? { ...area, cols: [e, e] } : { ...area, rows: [e, e] });
  const lookup = areaText(line(span[0]));
  const result = areaText(line({ abs: span[1].abs, n: span[0].n + index - 1 }));
  return `XLOOKUP(${args[0]},${lookup},${result}${fallback !== undefined ? `,${fallback}` : ''})`;
}

/** XLOOKUP(a,b,c) with `fallback` as its if_not_found argument, or null when it already has one. */
function withNotFound(call: Call, body: string, fallback: string): string | null {
  const args = call.args.map((s) => textOf(body, s));
  if (args.length < 3 || args.length > 6 || (args.length >= 4 && args[3] !== '')) return null;
  return `XLOOKUP(${[args[0], args[1], args[2], fallback, ...args.slice(4)].join(',')})`;
}

/** True when a lookup sits in the first argument of an IFERROR (or, with `orIfna`, an IFNA). */
function wrapsLookup(formula: string, orIfna: boolean): boolean {
  const calls = findCalls(maskFormula(normalizeFormula(formula)));
  return calls.some((w) => {
    if (w.name !== 'IFERROR' && !(orIfna && w.name === 'IFNA')) return false;
    const first = w.args[0];
    return !!first && calls.some((c) => ANY_LOOKUP.has(c.name) && c.start >= first.start && c.start < first.end);
  });
}

const iferrorRewrite: Rewrite = (call, body, code) => {
  const calls = findCalls(code);
  const first = call.args[0];
  if (!first || !calls.some((c) => ANY_LOOKUP.has(c.name) && c.start >= first.start && c.start < first.end)) return undefined;
  if (call.args.length !== 2) return null;
  const value = textOf(body, call.args[0]);
  const fallback = textOf(body, call.args[1]);
  const inner = soleCall(calls, code, call.args[0]);
  if (fallback && inner) {
    const x = inner.name === 'XLOOKUP' ? withNotFound(inner, body, fallback) : xlookupFor(inner, body, fallback);
    if (x) return x;
  }
  return call.name === 'IFERROR' ? `IFNA(${value},${fallback})` : null;
};

/** IF(a,x,IF(b,y,IF(c,z,w))) as IFS(a,x,b,y,c,z,TRUE,w), when the chain is three or more deep. */
const ifsRewrite: Rewrite = (call, body, code) => {
  const calls = findCalls(code);
  const parts: string[] = [];
  let otherwise = 'FALSE';
  for (let cur: Call | undefined = call; cur; ) {
    const args: Span[] = cur.args;
    const test = args[0] ? textOf(body, args[0]) : '';
    if (args.length < 2 || args.length > 3 || !test) return undefined;
    // An empty result argument returns 0, and a missing third argument returns FALSE.
    parts.push(test, textOf(body, args[1]) || '0');
    if (args.length === 2) break;
    const next = soleCall(calls, code, args[2]);
    if (next?.name === 'IF') {
      cur = next;
      continue;
    }
    otherwise = textOf(body, args[2]) || '0';
    cur = undefined;
  }
  return parts.length >= 6 ? `IFS(${parts.join(',')},TRUE,${otherwise})` : undefined;
};

interface Part {
  raw: string;
  code: string;
}

const FLIP: Record<string, string> = { '=': '=', '<>': '<>', '<': '>', '>': '<', '<=': '>=', '>=': '<=' };
const MASKED_PREFIX = /^(?:'[^']*'|[A-Za-z_][\w.]*)!/;
const THIS_ROW = /@|#this row/i;

/** A multi-cell reference: an A1 range, a Table column or a defined name. */
function isRange(p: Part): boolean {
  if (parseArea(p.raw)) return true;
  const m = p.code.replace(MASKED_PREFIX, '');
  if (/^(?:[A-Za-z_\\][\w.]*)?\[ *\]$/.test(m)) return !THIS_ROW.test(p.raw);
  return /^[A-Za-z_\\][\w.]*$/.test(m) && !/^\$?[A-Za-z]{1,3}\$?\d+$/.test(m) && !/^(TRUE|FALSE)$/i.test(m);
}

/** A single value: a literal, one cell, a this-row reference or an expression over those. */
const isScalar = (p: Part): boolean => p.code !== '' && !p.code.includes(':') && (!p.code.includes('[') || THIS_ROW.test(p.raw));

/** The single top-level comparison in `code[from, to)`: null when none, 'many' when several. */
function comparison(code: string, from: number, to: number): Span | 'many' | null {
  let found: Span | null = null;
  let depth = 0;
  for (let i = from; i < to; i++) {
    const ch = code[i];
    if (ch === '(' || ch === '{' || ch === '[') depth++;
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

/** SUMIFS criteria for "range <op> value": "East", G2, ">=5000" or ">="&H$1. */
function criteria(op: string, p: Part): string {
  if (op === '=') return p.raw;
  if (/^\d+(?:\.\d+)?$/.test(p.raw)) return `"${op}${p.raw}"`;
  if (/^"(?:[^"]|"")*"$/.test(p.raw)) return `"${op}${p.raw.slice(1)}`;
  return `"${op}"&${p.raw}`;
}

/** SUMPRODUCT((A=x)*(B>=y)*C) as SUMIFS(C,A,x,B,">="&y), or COUNTIFS when nothing is summed. */
const sumifsRewrite: Rewrite = (call, body, code) => {
  let sum: string | undefined;
  const pairs: string[] = [];
  const shapes = new Set<string>();
  const shape = (raw: string) => {
    const a = parseArea(raw);
    if (a) shapes.add(`${a.rows ? a.rows[1].n - a.rows[0].n : '*'}x${a.cols ? a.cols[1].n - a.cols[0].n : '*'}`);
  };
  for (const arg of call.args) {
    for (const factor of topLevelSpans(code, arg.start, arg.end, '*')) {
      let s = factor.start;
      let e = factor.end;
      let minus = 0;
      for (;;) {
        while (s < e && (code[s] === ' ' || code[s] === '+' || code[s] === '-')) if (code[s++] === '-') minus++;
        while (e > s && code[e - 1] === ' ') e--;
        if (code[s] !== '(' || closingParen(code, s) !== e - 1) break;
        s++;
        e--;
      }
      if (s >= e || minus % 2) return undefined;
      const cmp = comparison(code, s, e);
      if (cmp === 'many') return undefined;
      if (cmp) {
        const left = { raw: body.slice(s, cmp.start).trim(), code: code.slice(s, cmp.start).trim() };
        const right = { raw: body.slice(cmp.end, e).trim(), code: code.slice(cmp.end, e).trim() };
        const op = code.slice(cmp.start, cmp.end);
        if (isRange(left) && isScalar(right)) pairs.push(left.raw, criteria(op, right));
        else if (isRange(right) && isScalar(left)) pairs.push(right.raw, criteria(FLIP[op], left));
        else return undefined;
        shape(isRange(left) ? left.raw : right.raw);
        continue;
      }
      const part = { raw: body.slice(s, e).trim(), code: code.slice(s, e).trim() };
      if (isRange(part) && sum === undefined) {
        sum = part.raw;
        shape(part.raw);
      } else if (part.code !== '1') return undefined;
    }
  }
  // SUMIFS needs every range to be the same shape.
  if (!pairs.length || shapes.size > 1) return undefined;
  return sum !== undefined ? `SUMIFS(${sum},${pairs.join(',')})` : `COUNTIFS(${pairs.join(',')})`;
};

/** True when `code[from, to)` is plain arithmetic: no text joins or comparisons at the top level. */
function arithmeticOnly(code: string, from: number, to: number): boolean {
  return comparison(code, from, to) === null && topLevelSpans(code, from, to, '&').length === 1;
}

/** `expr + n`, folding a trailing "- k" so COUNTA(C:C)-1 plus 1 reads COUNTA(C:C). */
function plus(expr: string, n: number): string {
  if (n === 0) return expr;
  if (/^\d+$/.test(expr)) return String(Number(expr) + n);
  const m = /^([\s\S]*\S)\s*-\s*(\d+)$/.exec(expr);
  const masked = m ? maskFormula(m[1]) : '';
  if (m && (masked.match(/\(/g)?.length ?? 0) === (masked.match(/\)/g)?.length ?? 0)) {
    const k = n - Number(m[2]);
    return k === 0 ? m[1] : k > 0 ? `${m[1]}+${k}` : `${m[1]}-${-k}`;
  }
  return `${expr}+${n}`;
}

/** INDEX equivalents for the common OFFSET and INDIRECT shapes. */
const indexRewrite: Rewrite = (call, body, code) => {
  const args = call.args.map((s) => textOf(body, s));
  if (call.name === 'INDIRECT') {
    // INDIRECT("B"&n) is column B, row n. The text never moves, so the column is absolute.
    if (args.length < 1 || args.length > 2 || (args.length === 2 && !/^(TRUE|1)$/i.test(args[1]))) return undefined;
    const span = call.args[0];
    const m = /^"\$?([A-Za-z]{1,3})"\s*&/.exec(args[0]);
    const joins = topLevelSpans(code, span.start, span.end, '&');
    if (!m || joins.length !== 2) return undefined;
    const row = textOf(body, joins[1]);
    if (!row || !arithmeticOnly(code, joins[1].start, joins[1].end)) return undefined;
    const col = m[1].toUpperCase();
    return `INDEX($${col}:$${col},${row})`;
  }
  if (args.length < 3 || args.length > 5) return undefined;
  const anchor = CELL_TEXT.exec(args[0]);
  const [, rows, cols, height = '', width = ''] = args;
  if (!anchor || cols !== '0' || !call.args.slice(1).every((s) => arithmeticOnly(code, s.start, s.end))) return undefined;
  const [, prefix = '', colAbs, letters, rowAbs, rowText] = anchor;
  const column = `${prefix}${colAbs}${letters}:${colAbs}${letters}`;
  // Row numbers: literal when the anchor's row is fixed, ROW(anchor) when it moves as it fills.
  const rowPlus = (n: string, k: number) => (rowAbs ? plus(n, Number(rowText) + k) : `ROW(${args[0]})+${plus(n, k)}`);
  if (height === '' && width === '') return rows === '0' ? args[0] : `INDEX(${column},${rowPlus(rows, 0)})`;
  if (rows === '0' && height !== '' && (width === '' || width === '1')) return `${args[0]}:INDEX(${column},${rowPlus(height, -1)})`;
  return undefined;
};

const NUMBER_TOKEN = /(?<![A-Za-z0-9_.$:])(\d+(?:\.\d+)?(?:[eE][+-]?\d+)?)(%?)(?![A-Za-z0-9_(.:])/g;
const ARITHMETIC = '*/+-^<>=';
const OPERAND_END = /[A-Za-z0-9_)\]}"'%$#.]/;
const sameNumber = (a: number, b: number): boolean => Number(a.toPrecision(12)) === Number(b.toPrecision(12));
const formatNumber = (n: number): string => String(Number(n.toPrecision(12)));

/** The formula with `target` replaced by `ref` wherever it's typed into arithmetic (as arithmeticConstants sees it). */
function replaceConstant(formula: string, target: number, ref: string): string {
  const body = normalizeFormula(formula);
  const code = maskFormula(body);
  let out = '';
  let last = 0;
  for (const m of code.matchAll(NUMBER_TOKEN)) {
    if (!sameNumber(Number(m[1]) / (m[2] ? 100 : 1), target)) continue;
    let p = m.index - 1;
    while (p >= 0 && code[p] === ' ') p--;
    let before = p >= 0 ? code[p] : '';
    if (before === '-' || before === '+') {
      let q = p - 1;
      while (q >= 0 && code[q] === ' ') q--;
      if (q < 0 || !OPERAND_END.test(code[q])) before = q >= 0 ? code[q] : '';
    }
    let a = m.index + m[0].length;
    while (a < code.length && code[a] === ' ') a++;
    const after = a < code.length ? code[a] : '';
    if (!(before && ARITHMETIC.includes(before)) && !(after && ARITHMETIC.includes(after))) continue;
    out += body.slice(last, m.index) + ref;
    last = m.index + m[0].length;
  }
  return `=${out}${body.slice(last)}`;
}

const constantsIn = (formula: string): number[] => arithmeticConstants(maskFormula(normalizeFormula(formula)));

// ---------- The block to fix ----------

/** A vertical run of cells in one column of the scanned sheet. */
interface Run {
  col: number;
  top: number;
  /** A1 formulas or typed values, top to bottom. */
  cells: Cell[];
  /** R1C1 formulas for the same cells; blank when the host gave none. */
  r1c1: Cell[];
  /** A short text label directly above the run. */
  header?: string;
  /** The column goes on above or below the run: it was cut at MAX_FIX_ROWS or at the last scanned row. */
  cut: { above: boolean; below: boolean };
}

/**
 * The contiguous run of formulas in the column through `cell`, stopping at headers, blanks and
 * totals. With `withTyped`, numbers sandwiched between formulas belong to the run (typed over a
 * formula), as does one number ending the column right after its last formula, as in the scanner.
 * With `partial`, a run that reaches the last scanned row is taken to go on below it.
 */
function findRun(sheet: SheetFormulas, cell: CellRef, withTyped: boolean, partial = false): Run | null {
  let origin: CellRef;
  try {
    origin = parseRange(sheet.address || 'A1').start;
  } catch {
    return null;
  }
  const grid = sheet.formulas ?? [];
  const r1c1 = sheet.r1c1 ?? [];
  const c = cell.col - origin.col;
  const start = cell.row - origin.row;
  if (start < 0 || c < 0 || start >= grid.length) return null;
  const at = (r: number): Cell => grid[r]?.[c] ?? '';
  const keyAt = (r: number): string | undefined => {
    const alt = r1c1[r]?.[c];
    return isFormula(alt) ? alt : undefined;
  };
  // A total or subtotal under (or above) the column isn't part of the fill. It's told apart from the
  // column's own formula (the one at the cell, or beside a typed value): copies of that never count,
  // and when the column itself reads whole Table columns (XLOOKUP into Items[SKU]), only a range of
  // its own column marks a total.
  const anchor = keyAt(start) ?? keyAt(start - 1) ?? keyAt(start + 1);
  const anchorSummary = anchor !== undefined && runRefFacts(anchor, origin.col + c).summary;
  const total = (r: number): boolean => {
    const alt = keyAt(r);
    if (alt === undefined || alt === anchor) return false;
    const facts = runRefFacts(alt, origin.col + c);
    return facts.ownRange || (facts.summary && !anchorSummary);
  };
  const inside = (r: number): boolean => {
    const v = at(r);
    return isFormula(v) ? !total(r) : withTyped && typeof v === 'number';
  };
  if (!inside(start)) return null;

  let lo = start;
  let hi = start;
  while (lo > 0 && inside(lo - 1)) lo--;
  while (hi + 1 < grid.length && inside(hi + 1)) hi++;
  if (withTyped) {
    while (lo < start && !isFormula(at(lo))) lo++;
    let last = hi;
    while (last > lo && !isFormula(at(last))) last--;
    const endsColumn = hi + 1 >= grid.length || !isFormula(at(hi + 1));
    if (!(hi - last === 1 && endsColumn)) hi = Math.max(last, start);
  }
  const cut = { above: false, below: partial && hi + 1 >= grid.length };
  if (hi - lo + 1 > MAX_FIX_ROWS) {
    if (start - lo >= MAX_FIX_ROWS) {
      lo = start;
      cut.above = true;
    }
    if (hi > lo + MAX_FIX_ROWS - 1) {
      hi = lo + MAX_FIX_ROWS - 1;
      cut.below = true;
    }
  }

  const cells: Cell[] = [];
  const alts: Cell[] = [];
  for (let r = lo; r <= hi; r++) {
    cells.push(at(r));
    alts.push(r1c1[r]?.[c] ?? '');
  }
  if (!cells.some(isFormula)) return null;

  // The text right above is the column's header unless it interrupts a column of formulas.
  const above = lo > 0 ? at(lo - 1) : '';
  const label = typeof above === 'string' && !above.startsWith('=') ? above.replace(/`/g, '').trim() : '';
  const header = label && label.length <= 40 && !(lo > 1 && isFormula(at(lo - 2))) ? label : undefined;
  return { col: origin.col + c, top: origin.row + lo, cells, r1c1: alts, header, cut };
}

interface RunCell {
  row: number;
  address: string;
  value: Cell;
}

interface FormulaCell extends RunCell {
  formula: string;
  /** R1C1 text (or an equivalent built from A1): the same for every copy of one filled formula. */
  rc: string;
  facts: FormulaFacts;
}

interface FixContext {
  finding: Finding;
  sheet: SheetFormulas;
  run: Run;
  range: string;
  cells: (RunCell | FormulaCell)[];
  formulas: FormulaCell[];
  /** Says the column runs past the block, when it does. */
  continues?: string;
}

const hasFormula = (c: RunCell | FormulaCell): c is FormulaCell => 'formula' in c;

function contextFor(finding: Finding, sheet: SheetFormulas, run: Run): FixContext {
  // Filled-down copies share their R1C1 text, so a 5,000-row column is one analysis.
  const analysed = new Map<string, FormulaFacts>();
  const cells = run.cells.map((value, i): RunCell | FormulaCell => {
    const row = run.top + i;
    const address = cellAddress({ row, col: run.col });
    if (!isFormula(value)) return { row, address, value };
    const alt = run.r1c1[i];
    const rc = isFormula(alt) ? alt : `=${relativeKey(value, address)}`;
    let facts = analysed.get(rc);
    if (!facts) analysed.set(rc, (facts = analyzeFormula(value)));
    return { row, address, value, formula: value, rc, facts };
  });
  return {
    finding,
    sheet,
    run,
    range: rangeAddress({ row: run.top, col: run.col }, run.cells.length, 1),
    cells,
    formulas: cells.filter(hasFormula),
    continues: continuesText(run),
  };
}

/** The sentence that says the column runs past the block, for the fill step and the copy-back. */
function continuesText(run: Run): string | undefined {
  const top = mono(cellAddress({ row: run.top, col: run.col }));
  const bottom = mono(cellAddress({ row: run.top + run.cells.length - 1, col: run.col }));
  if (run.cut.above && run.cut.below) return `The column continues above ${top} and below ${bottom}. Fill the new formula through all of it.`;
  if (run.cut.above) return `The column continues above ${top}. Fill the new formula up to the top of it.`;
  if (run.cut.below) return `The column continues below ${bottom}. Fill the new formula to the end of it.`;
  return undefined;
}

/** The formula most of the run shares, when it's more than half of `span` cells. */
function majority(formulas: FormulaCell[], span: number): FormulaCell[] | null {
  const groups = new Map<string, FormulaCell[]>();
  for (const c of formulas) {
    const group = groups.get(c.rc);
    if (group) group.push(c);
    else groups.set(c.rc, [c]);
  }
  let best: FormulaCell[] = [];
  for (const group of groups.values()) if (group.length > best.length) best = group;
  return best.length >= 2 && best.length * 2 > span ? best : null;
}

/** The cell of `group` closest above `row`, else closest below. */
function nearest(group: FormulaCell[], row: number): FormulaCell {
  let above: FormulaCell | undefined;
  for (const c of group) if (c.row < row) above = c;
  return above ?? group.find((c) => c.row > row) ?? group[0];
}

/**
 * Cells whose values follow from the repaired ones: in a running column (each row builds on the
 * row above), everything below the first repair; when rows build on the row below, everything above.
 */
function downstream(ctx: FixContext, fill: RunRefFacts, repaired: number[]): string[] {
  if (!fill.ownCell || !repaired.length) return [];
  const first = Math.min(...repaired);
  const last = Math.max(...repaired);
  return ctx.cells.filter((c) => (fill.rowAbove && c.row > first) || (fill.rowBelow && c.row < last)).map((c) => c.address);
}

// ---------- Copy helpers ----------

const mono = (text: string): string => `\`${text}\``;

/** `D9`, `D9` and `D14`, or `D9`, `D14` and 3 more. */
function codeList(items: string[], max = 3): string {
  const shown = items.length > max ? items.slice(0, max - 1) : items;
  const rest = items.length - shown.length;
  const parts = shown.map(mono);
  if (rest > 0) return `${parts.join(', ')} and ${rest} more`;
  return parts.length > 1 ? `${parts.slice(0, -1).join(', ')} and ${parts[parts.length - 1]}` : parts[0];
}

const plural = (n: number, one: string, many = `${one}s`): string => (n === 1 ? one : many);

function columnName(run: Run): string {
  return run.header ? `the ${run.header} column` : 'the column';
}

/**
 * The fill step: copy the new formula down when every formula below `source` is a copy of it,
 * otherwise rewrite the other affected cells by hand. Null when nothing else needs changing.
 */
function fillStep(ctx: FixContext, source: FormulaCell, affected: FormulaCell[], what: string): string | null {
  if (affected.length <= 1) return null;
  const below = ctx.formulas.filter((c) => c.row > source.row);
  const last = ctx.run.top + ctx.run.cells.length - 1;
  const more = ctx.continues ? ` ${ctx.continues}` : '';
  if (below.length === last - source.row && below.every((c) => c.rc === source.rc)) {
    const rest = rangeAddress({ row: source.row + 1, col: ctx.run.col }, last - source.row, 1);
    return `Fill it down: copy ${mono(source.address)}, select ${mono(rest)}, and paste.${more}`;
  }
  return `Rewrite the other ${what} in ${mono(ctx.range)} the same way.${more}`;
}

const checkStep = (ctx: FixContext): string => `Select Check my fix. Every result in ${mono(ctx.range)} must still match your original sheet.`;

type PlanFields = Pick<FixPlan, 'title' | 'intro'> & Partial<Omit<FixPlan, 'steps'>> & { steps: (string | null)[] };

function plan(ctx: FixContext, { steps, ...fields }: PlanFields): FixPlan {
  return {
    findingId: ctx.finding.id,
    sheet: ctx.sheet.sheet,
    range: ctx.range,
    valuesMayChange: [],
    forbid: [],
    require: [],
    allFormulas: false,
    continues: ctx.continues,
    ...fields,
    steps: steps.filter((step): step is string => !!step),
  };
}

const matches = (rule: FixRule, formula: string): boolean => (rule.test ? rule.test(formula) : rule.pattern.test(maskFormula(normalizeFormula(formula))));

/** A rewrite counts only when it actually clears the rules it's meant to satisfy. */
const clears = (formula: string | null, forbid: FixRule[]): formula is string => formula !== null && forbid.every((r) => !matches(r, formula));

// ---------- Plans ----------

function planLookup(ctx: FixContext): FixPlan | null {
  const affected = ctx.formulas.filter((c) => c.facts.vlookupColumnNumber);
  const source = affected[0];
  // Approximate matches can give different results once rewritten; that's the vlookup-approximate finding.
  if (!source || source.facts.vlookupApproximate) return null;
  const fn = /\bVLOOKUP\s*\(/i.test(maskFormula(normalizeFormula(source.formula))) ? 'VLOOKUP' : 'HLOOKUP';
  const unit = fn === 'VLOOKUP' ? 'column' : 'row';
  const forbid: FixRule[] = [
    { pattern: /\b[VH]LOOKUP\s*\(/i, label: 'No VLOOKUP or HLOOKUP left', advice: `Replace each one with XLOOKUP, pointing at the key ${unit} and the return ${unit}.` },
  ];
  const formula = rewriteCalls(source.formula, LOOKUP_CALLS, (call, body) => xlookupFor(call, body));
  const exact = clears(formula, forbid) ? formula : null;
  const body = normalizeFormula(source.formula);
  const first = findCalls(maskFormula(body)).find((c) => LOOKUP_CALLS.has(c.name));
  const args = first?.args.map((s) => textOf(body, s)) ?? [];
  const require: FixRule[] =
    affected.length === ctx.formulas.length
      ? [
          {
            pattern: /\bXLOOKUP\s*\(|\bINDEX\s*\([\s\S]*\bX?MATCH\s*\(/i,
            label: 'Uses XLOOKUP or INDEX and MATCH',
            advice: exact ? `For example, ${source.address} becomes ${exact}.` : `Use XLOOKUP(lookup_value, lookup_array, return_array).`,
          },
        ]
      : [];

  return plan(ctx, {
    title: `Replace ${fn} with XLOOKUP`,
    intro: `${fn} counts ${unit}s by position, so inserting or deleting a ${unit} in the lookup range quietly returns the wrong field. XLOOKUP points at the return ${unit} itself.`,
    steps: [
      args.length >= 3 && /^\d+$/.test(args[2])
        ? `On the copy, select ${mono(source.address)}. Its ${fn} returns ${unit} ${args[2]} of ${mono(args[1])}, counted by position.`
        : `On the copy, select ${mono(source.address)}.`,
      exact
        ? `Replace the formula with ${mono(exact)}. It names the key ${unit} and the return ${unit} directly, so a new ${unit} can’t shift the result.`
        : `Rewrite each ${fn} as ${mono('XLOOKUP(lookup_value, lookup_array, return_array)')}: the ${unit} ${fn} searches, then the ${unit} it returns.`,
      fillStep(ctx, source, affected, `${fn} formulas`),
      checkStep(ctx),
    ],
    forbid,
    require,
  });
}

function planNotFound(ctx: FixContext): FixPlan | null {
  const affected = ctx.formulas.filter((c) => c.facts.iferrorLookup);
  const source = affected[0];
  if (!source) return null;
  // IFNA is a fine outcome for an IFERROR; an IFNA itself is fixed by moving its fallback into XLOOKUP.
  const wrapper = wrapsLookup(source.formula, false) ? 'IFERROR' : 'IFNA';
  const orIfna = wrapper === 'IFNA';
  const forbid: FixRule[] = [
    {
      pattern: orIfna ? /\bIF(?:ERROR|NA)\s*\(/i : /\bIFERROR\s*\(/i,
      test: (f) => wrapsLookup(f, orIfna),
      label: orIfna ? 'No IFERROR or IFNA around the lookup' : 'No IFERROR around the lookup',
      advice: orIfna ? 'Give XLOOKUP the fallback as its fourth argument instead.' : 'Use XLOOKUP’s if_not_found argument, or IFNA, so only a missing key gets the fallback.',
    },
  ];
  const formula = rewriteCalls(source.formula, WRAPPERS, iferrorRewrite);
  const exact = clears(formula, forbid) ? formula : null;
  const body = normalizeFormula(source.formula);
  const calls = findCalls(maskFormula(body));
  const wrap = calls.find((c) => c.name === wrapper && c.args.length === 2);
  const fallback = wrap ? textOf(body, wrap.args[1]) : '';
  const shows = fallback === '""' ? 'a blank' : fallback ? mono(fallback) : 'its fallback';
  const viaIfna = !!exact && /\bIFNA\s*\(/i.test(exact);

  return plan(ctx, {
    title: orIfna ? 'Move the fallback into XLOOKUP' : 'Use the fallback for missing keys only',
    intro: orIfna
      ? 'XLOOKUP’s if_not_found argument gives a missing key its fallback directly, so the formula reads as one lookup with no wrapper around it.'
      : 'IFERROR turns every error into the same fallback, so an error in the data or a broken range looks like a missing item. XLOOKUP’s if_not_found argument catches only keys that aren’t found.',
    steps: [
      orIfna
        ? `On the copy, select ${mono(source.address)}. IFNA wraps the lookup to show ${shows} when the key is missing.`
        : `On the copy, select ${mono(source.address)}. IFERROR shows ${shows} for any error, not only a missing key.`,
      !exact
        ? `Move the fallback into XLOOKUP’s fourth argument, ${mono('XLOOKUP(key, lookup_array, return_array, if_not_found)')}${orIfna ? '' : ', or wrap the lookup in IFNA instead of IFERROR'}.`
        : viaIfna
          ? `Replace the formula with ${mono(exact)}. IFNA catches only #N/A, the error a missing key returns.`
          : `Replace the formula with ${mono(exact)}. The last argument is used only when the key isn’t found; other errors still show.`,
      fillStep(ctx, source, affected, 'lookups'),
      orIfna
        ? null
        : 'A mistyped key or a number stored as text still gets the fallback, because it isn’t found. If a cell now shows an error, IFERROR was hiding it: an error in the key cell or the cell it returns, or a broken range. Check my fix points to it.',
      checkStep(ctx),
    ],
    forbid,
    fallback: orIfna ? undefined : literalValue(fallback),
  });
}

/** The value a formula literal shows: 0, "" (blank), "Missing" or TRUE. Undefined for anything else. */
function literalValue(text: string): Cell | undefined {
  if (/^-?\d+(?:\.\d+)?$/.test(text)) return Number(text);
  if (/^"(?:[^"]|"")*"$/.test(text)) return text.slice(1, -1).replace(/""/g, '"');
  if (/^(TRUE|FALSE)$/i.test(text)) return text.toUpperCase() === 'TRUE';
  return undefined;
}

function planIfs(ctx: FixContext): FixPlan | null {
  const affected = ctx.formulas.filter((c) => c.facts.ifDepth >= 3);
  const source = affected[0];
  if (!source) return null;
  const forbid: FixRule[] = [
    {
      pattern: /\bIF\s*\(/i,
      test: (f) => analyzeFormula(f).ifDepth >= 3,
      label: 'No IFs nested three or more deep',
      advice: 'List the conditions in order with IFS, or look the tiers up in a small table with XLOOKUP.',
    },
  ];
  const formula = rewriteCalls(source.formula, new Set(['IF']), ifsRewrite);
  const exact = clears(formula, forbid) ? formula : null;

  return plan(ctx, {
    title: 'Replace nested IFs with IFS',
    intro: 'IFs inside IFs are hard to read, and one misplaced parenthesis changes the result. IFS lists each condition with its result, in order. For tiers on one number, XLOOKUP on a small threshold table works too.',
    steps: [
      `On the copy, select ${mono(source.address)}. Its IFs nest ${source.facts.ifDepth} levels deep.`,
      exact
        ? `Replace the formula with ${mono(exact)}. IFS returns the result for the first condition that’s TRUE, and the final ${mono('TRUE')} catches everything else.`
        : `Rewrite the chain as ${mono('IFS(test1, result1, test2, result2, …, TRUE, otherwise)')}, or put the tiers in a small table and use XLOOKUP with match mode ${mono('-1')}.`,
      fillStep(ctx, source, affected, 'nested IFs'),
      checkStep(ctx),
    ],
    forbid,
  });
}

function planSumifs(ctx: FixContext): FixPlan | null {
  const affected = ctx.formulas.filter((c) => c.facts.sumproductConditions);
  const source = affected[0];
  if (!source) return null;
  const forbid: FixRule[] = [
    {
      pattern: /\bSUMPRODUCT\s*\(/i,
      test: (f) => analyzeFormula(f).sumproductConditions,
      label: 'No SUMPRODUCT used as a conditional sum',
      advice: 'Use SUMIFS or COUNTIFS, with one range-and-criteria pair per condition.',
    },
  ];
  const formula = rewriteCalls(source.formula, new Set(['SUMPRODUCT']), sumifsRewrite);
  const exact = clears(formula, forbid) ? formula : null;
  const counts = !!exact && /\bCOUNTIFS\(/i.test(exact) && !/\bSUMIFS\(/i.test(exact);

  return plan(ctx, {
    title: `Replace SUMPRODUCT with ${counts ? 'COUNTIFS' : 'SUMIFS'}`,
    intro: `${counts ? 'COUNTIFS' : 'SUMIFS'} takes one range-and-criteria pair per condition, so the formula reads as what it ${counts ? 'counts' : 'adds up'}. It’s also faster than multiplying comparisons on long lists.`,
    steps: [
      `On the copy, select ${mono(source.address)}. Its SUMPRODUCT multiplies comparisons together to ${counts ? 'count' : 'add up'} the matching rows.`,
      exact
        ? counts
          ? `Replace the formula with ${mono(exact)}. Each pair is a range and the criteria it must meet; COUNTIFS counts the rows that meet all of them.`
          : `Replace the formula with ${mono(exact)}. The first range is what gets added up; each pair after it is a range and the criteria it must meet.`
        : `Rewrite it as ${mono('SUMIFS(sum_range, criteria_range1, criteria1, …)')}, one pair per condition. Write a comparison as text joined to its value, like ${mono('">="&H1')}.`,
      fillStep(ctx, source, affected, 'SUMPRODUCT formulas'),
      checkStep(ctx),
    ],
    forbid,
  });
}

function planVolatile(ctx: FixContext): FixPlan | null {
  const affected = ctx.formulas.filter((c) => c.facts.volatile);
  const source = affected[0];
  if (!source) return null;
  const masked = maskFormula(normalizeFormula(source.formula));
  const fns = ['OFFSET', 'INDIRECT'].filter((fn) => new RegExp(String.raw`\b${fn}\s*\(`, 'i').test(masked));
  const named = fns.join(' and ') || 'OFFSET';
  const forbid: FixRule[] = [{ pattern: /\b(?:OFFSET|INDIRECT)\s*\(/i, label: 'No OFFSET or INDIRECT left', advice: 'Pick the cells with INDEX, XLOOKUP or a Table column instead.' }];
  const formula = rewriteCalls(source.formula, new Set(['OFFSET', 'INDIRECT']), indexRewrite);
  const exact = clears(formula, forbid) ? formula : null;

  return plan(ctx, {
    title: fns.length === 2 ? 'Replace OFFSET and INDIRECT' : `Replace ${named} with INDEX`,
    intro: `${named} ${fns.length === 2 ? 'build references' : 'builds its reference'} while Excel calculates, so moving or renaming things breaks ${fns.length === 2 ? 'them' : 'it'} without warning, and every edit recalculates ${fns.length === 2 ? 'them' : 'it'}. INDEX picks the same cells without either problem.`,
    steps: [
      `On the copy, select ${mono(source.address)}.`,
      exact
        ? `Replace the formula with ${mono(exact)}. It returns the same cells, and Excel recalculates it only when they change.`
        : `Rewrite the reference without ${named}: INDEX to pick a cell or the end of a range, XLOOKUP to find a row, or a Table column such as ${mono('Sales[Amount]')} for a list that grows.`,
      fillStep(ctx, source, affected, `${named} formulas`),
      checkStep(ctx),
    ],
    forbid,
  });
}

function planTypedOver(ctx: FixContext): FixPlan | null {
  const typed = ctx.cells.filter((c) => !hasFormula(c));
  const fill = majority(ctx.formulas, ctx.cells.length);
  if (!typed.length || !fill) return null;
  const model = nearest(fill, typed[0].row);
  const spill = downstream(ctx, runRefFacts(model.rc, ctx.run.col), typed.map((c) => c.row));
  const where = typed.map((c) => c.address);
  const values = typed.map((c) => (typeof c.value === 'number' ? formatNumber(c.value) : String(c.value)));

  return plan(ctx, {
    title: 'Restore the column’s formula',
    intro: `A number typed over a formula looks like a result but never updates, so ${columnName(ctx.run)} goes stale without any sign. Putting the formula back keeps every row live.`,
    steps: [
      typed.length === 1
        ? `On the copy, select ${mono(where[0])}. It holds the typed number ${mono(values[0])} where the rest of ${columnName(ctx.run)} has a formula.`
        : `On the copy, ${codeList(where, 4)} hold typed numbers where the rest of ${columnName(ctx.run)} has a formula.`,
      `If ${typed.length === 1 ? 'it was' : 'one was'} a deliberate override, note the number first. Overrides belong in a separate, labeled input column, not on top of a formula.`,
      `Copy ${mono(model.address)}, which has the column’s formula, and paste it into ${typed.length === 1 ? mono(where[0]) : typed.length <= 4 ? codeList(where, 4) : 'each typed cell'}. Excel adjusts the references for each row.`,
      `Select Check my fix. The restored ${plural(typed.length, 'cell')} can show new values${spill.length ? ', and so can the rows that build on them' : ''}; every other result must match your original sheet.`,
    ],
    valuesMayChange: [...new Set([...where, ...spill])],
    allFormulas: true,
    // The fix touches only the typed cells, so there's no new formula to carry down the column.
    continues: undefined,
    restore: {
      cells: where,
      model: model.address,
      key: relativeKey(model.formula, model.address),
      label: 'Typed values replaced with the column’s formula',
      advice: `Copy ${model.address} and paste it over each typed value.`,
    },
  });
}

/** Mirrors the scanner: cells that differ from their run on purpose (totals, seeds, deliberate lines). */
function excused(cell: RunRefFacts, fill: RunRefFacts, first: boolean, last: boolean): boolean {
  if (cell.ownRange || cell.summary) return true;
  if (cell.ownCell && !fill.ownCell) return true;
  if (first && fill.rowAbove) return true;
  if (last && (fill.rowBelow || cell.ownCell)) return true;
  return cell.hasRef && cell.shape !== fill.shape && !cell.shape.includes(fill.shape);
}

function planConsistent(ctx: FixContext): FixPlan | null {
  const fill = majority(ctx.formulas, ctx.cells.length);
  if (!fill) return null;
  const fillRc = fill[0].rc;
  const fillFacts = runRefFacts(fillRc, ctx.run.col);
  const flagged = new Set(ctx.finding.cells.filter((c) => c.sheet === ctx.sheet.sheet).map((c) => c.address));
  const first = ctx.run.top;
  const last = ctx.run.top + ctx.run.cells.length - 1;

  const odd: FormulaCell[] = [];
  for (let i = 0; i < ctx.formulas.length; ) {
    let j = i + 1;
    while (j < ctx.formulas.length && ctx.formulas[j].rc === ctx.formulas[i].rc && ctx.formulas[j].row === ctx.formulas[j - 1].row + 1) j++;
    const segment = ctx.formulas.slice(i, j);
    if (segment[0].rc !== fillRc) {
      const facts = runRefFacts(segment[0].rc, ctx.run.col);
      for (const c of segment) {
        const deliberate = segment.length >= MIN_SECTION || excused(facts, fillFacts, c.row === first, c.row === last);
        if (!deliberate || flagged.has(c.address)) odd.push(c);
      }
    }
    i = j;
  }
  if (!odd.length) return null;
  const model = nearest(fill, odd[0].row);
  const spill = downstream(ctx, fillFacts, odd.map((c) => c.row));
  const where = odd.map((c) => c.address);

  return plan(ctx, {
    title: 'Fill one formula through the column',
    intro: `Most of ${mono(ctx.range)} shares one formula, but ${odd.length === 1 ? 'one cell differs' : `${odd.length} cells differ`}. That usually means a fill was interrupted or a cell was edited by hand, so ${plural(odd.length, 'its', 'their')} results can be wrong without any sign.`,
    steps: [
      odd.length === 1
        ? `On the copy, compare ${mono(where[0])} with ${mono(model.address)}. ${mono(model.address)} has ${mono(model.formula)}, like the rest of the column, but ${mono(where[0])} has ${mono(odd[0].formula)}.`
        : `On the copy, compare ${codeList(where, 4)} with ${mono(model.address)}, which has the formula the rest of the column shares: ${mono(model.formula)}.`,
      `If ${odd.length === 1 ? 'the' : 'a'} difference was deliberate, note it first. An adjustment belongs in its own labeled cell, not inside one row’s formula.`,
      `Copy ${mono(model.address)} and paste it into ${odd.length <= 4 ? codeList(where, 4) : 'each odd cell'}. Excel adjusts the references for each row.`,
      `Select Check my fix. The repaired ${plural(odd.length, 'cell')} can show new values${spill.length ? ', and so can the rows that build on them' : ''}; every other result must match your original sheet.`,
    ],
    valuesMayChange: [...new Set([...where, ...spill])],
    allFormulas: true,
    continues: undefined,
    restore: {
      cells: where,
      model: model.address,
      key: relativeKey(model.formula, model.address),
      label: 'Odd cells now match the column’s formula',
      advice: `Copy ${model.address} and paste it over each cell that differs.`,
    },
  });
}

/** Where one typed number turns up on a sheet: the distinct formulas and how many cells they fill. */
interface Spread {
  formulas: Set<string>;
  cells: number;
}

function planInputCell(ctx: FixContext): FixPlan | null {
  // Which numbers: the ones the finding names, or that this sheet types in the way the scanner
  // flags: into three or more formulas, or five or more cells (single digits only across formulas).
  const spread = new Map<number, Spread>();
  const grid = ctx.sheet.formulas ?? [];
  const r1c1 = ctx.sheet.r1c1 ?? [];
  const seen = new Map<string, number[]>();
  grid.forEach((row, r) =>
    row?.forEach((v, c) => {
      if (!isFormula(v)) return;
      const alt = r1c1[r]?.[c];
      const key = isFormula(alt) ? alt : v;
      let numbers = seen.get(key);
      if (!numbers) seen.set(key, (numbers = constantsIn(v)));
      for (const n of numbers) {
        let s = spread.get(n);
        if (!s) spread.set(n, (s = { formulas: new Set(), cells: 0 }));
        s.formulas.add(key);
        s.cells++;
      }
    }),
  );
  const formulasWith = (n: number) => spread.get(n)?.formulas.size ?? 0;
  const cellsWith = (n: number) => spread.get(n)?.cells ?? 0;
  const repeated = (n: number) => isRepeatedConstant(n, formulasWith(n), cellsWith(n));
  // The scan says which numbers it flagged; older findings only name them in the why sentence.
  const named = (n: number) =>
    ctx.finding.numbers
      ? ctx.finding.numbers.some((x) => Math.abs(x - n) <= 1e-9 * Math.max(1, Math.abs(n)))
      : new RegExp(String.raw`(?<![\d.])${formatNumber(n).replace('.', '\\.')}(?![\d])`).test(ctx.finding.why);
  const candidates = [...new Set(ctx.formulas.flatMap((c) => c.facts.constants))];
  let targets = candidates.filter((n) => named(n) || repeated(n));
  if (!targets.length) targets = ctx.formulas.find((c) => c.facts.constants.length)?.facts.constants ?? [];
  targets = targets.sort((a, b) => Number(named(b)) - Number(named(a)) || formulasWith(b) - formulasWith(a) || cellsWith(b) - cellsWith(a)).slice(0, 3);
  const affected = ctx.formulas.filter((c) => c.facts.constants.some((n) => targets.some((t) => sameNumber(n, t))));
  const source = affected[0];
  if (!source || !targets.length) return null;

  // The input cells go past the right edge of the used range, with a blank column between.
  let right: number;
  let top: number;
  try {
    const used = parseRange(ctx.sheet.address || 'A1');
    right = used.end.col;
    top = used.start.row;
  } catch {
    return null;
  }
  const labelCol = right + 2;
  if (labelCol + 1 > 16384) return null;
  const inputs = targets.map((n, i) => ({ n, label: cellAddress({ row: top + i, col: labelCol }), cell: cellAddress({ row: top + i, col: labelCol + 1 }) }));
  const absolute = (a: string) => a.replace(/^([A-Z]+)(\d+)$/, '$$$1$$$2');
  let formula = source.formula;
  for (const input of inputs) formula = replaceConstant(formula, input.n, absolute(input.cell));
  const forbid: FixRule[] = inputs.map(({ n, cell }) => ({
    pattern: new RegExp(String.raw`(?<![\w.$])${formatNumber(n).replace('.', '\\.')}(?![\w.])`),
    test: (f: string) => constantsIn(f).some((x) => sameNumber(x, n)),
    label: `No ${formatNumber(n)} typed into the formulas`,
    advice: `Point at the input cell instead, like ${absolute(cell)}.`,
  }));
  const exact = clears(formula, forbid) ? formula : null;
  const numbers = inputs.map((i) => formatNumber(i.n));
  const block = rangeAddress(parseCell(inputs[0].label), inputs.length, 2);

  return plan(ctx, {
    title: numbers.length === 1 ? `Move ${numbers[0]} into an input cell` : 'Move the typed numbers into input cells',
    intro: `${numbers.length === 1 ? `${numbers[0]} is` : `${numbers.slice(0, -1).join(', ')} and ${numbers[numbers.length - 1]} are`} typed into the formulas, so changing ${numbers.length === 1 ? 'it' : 'one'} means finding every copy. In a labeled input cell, it’s visible and changes in one place.`,
    steps: [
      inputs.length === 1
        ? `On the copy, type ${mono(numbers[0])} in ${mono(inputs[0].cell)}, with a short label in ${mono(inputs[0].label)} that says what it is.`
        : `On the copy, type ${inputs.map((i) => `${mono(formatNumber(i.n))} in ${mono(i.cell)}`).join(inputs.length === 2 ? ' and ' : ', ')}, each with a short label to its left that says what it is.`,
      exact
        ? `In ${mono(source.address)}, replace the typed ${plural(inputs.length, 'number')} with ${plural(inputs.length, 'a reference', 'references')}: ${mono(exact)}. Select the input cell while you edit and press {absKey} for the dollar signs, so the reference stays put when you fill.`
        : `In ${mono(source.address)}, replace ${codeList(numbers)} with ${plural(inputs.length, 'a reference', 'references')} to the input ${plural(inputs.length, 'cell')}, like ${mono(absolute(inputs[0].cell))}. Press {absKey} for the dollar signs, so the reference stays put when you fill.`,
      fillStep(ctx, source, affected, `formulas that use ${numbers.length === 1 ? numbers[0] : 'these numbers'}`),
      `Select Check my fix. The results should match exactly: the ${plural(inputs.length, 'number hasn’t', 'numbers haven’t')} changed, only where ${plural(inputs.length, 'it lives', 'they live')}.`,
    ],
    forbid,
    copyBack: `When you’re ready, copy the input ${plural(inputs.length, 'cell')} and ${plural(inputs.length, 'its label', 'their labels')} (${mono(block)}) to the same place on ${mono(ctx.sheet.sheet)} first. Then select ${mono(ctx.range)} on the copy, copy, and paste onto the same range on ${mono(ctx.sheet.sheet)}.`,
  });
}

/** A Table name built from the sheet's, e.g. "OpsData". */
function tableNameFor(sheet: string): string {
  const stem = sheet.replace(/[^A-Za-z0-9_]/g, '');
  return /^[A-Za-z_]/.test(stem) ? `${stem}Data` : `Data${stem}`;
}

const structuredName = (header: string): string => header.replace(/(['#[\]])/g, "'$1");

const LONG_AREA = /(?<![A-Za-z0-9_.$])((?:'[^']*'|[A-Za-z_][\w.]*)!)?\$?[A-Za-z]{1,3}\$?\d+:\$?[A-Za-z]{1,3}\$?\d+(?![A-Za-z0-9_(!])/g;

/** Explicit A1 ranges of LONG_RANGE_ROWS or more in a normalized formula, with their positions. */
function longRanges(body: string): { start: number; end: number; area: Area }[] {
  const out: { start: number; end: number; area: Area }[] = [];
  for (const m of maskFormula(body).matchAll(LONG_AREA)) {
    const area = parseArea(body.slice(m.index, m.index + m[0].length));
    if (area?.rows && area.cols && area.rows[1].n - area.rows[0].n + 1 >= LONG_RANGE_ROWS) out.push({ start: m.index, end: m.index + m[0].length, area });
  }
  return out;
}

interface TableShape {
  name: string;
  /** Header row through the last data row, e.g. "A1:C40". */
  range: string;
  formula: string;
}

/**
 * The Table that replaces the long ranges in `formula`, and the formula rewritten to use it. Null
 * when the exact shape can't be worked out; 'other-sheet' or 'beyond' when there should be no plan.
 */
function tableFor(ctx: FixContext, formula: string): TableShape | null | 'other-sheet' | 'beyond' {
  const body = normalizeFormula(formula);
  const longs = longRanges(body);
  if (!longs.length) return null;
  // Converting another sheet's data would change a sheet the fix doesn't copy.
  if (longs.some((l) => l.area.prefix && prefixSheet(l.area.prefix).toLowerCase() !== ctx.sheet.sheet.toLowerCase())) return 'other-sheet';

  let origin: CellRef;
  try {
    origin = parseRange(ctx.sheet.address || 'A1').start;
  } catch {
    return null;
  }
  const at = (row: number, col: number): Cell => ctx.sheet.formulas[row - origin.row]?.[col - origin.col] ?? '';
  const isHeader = (row: number, col: number) => {
    const v = at(row, col);
    return typeof v === 'string' && v.trim() !== '' && !isFormula(v);
  };
  const firstRow = longs[0].area.rows![0].n;
  if (longs.some((l) => l.area.rows![0].n !== firstRow)) return null;
  const header = firstRow - 1;
  let left = Math.min(...longs.map((l) => l.area.cols![0].n));
  let right = Math.max(...longs.map((l) => l.area.cols![1].n));
  for (let c = left; c <= right; c++) if (header < origin.row || !isHeader(header, c)) return null;
  while (left - 1 >= origin.col && left - 1 !== ctx.run.col && isHeader(header, left - 1)) left--;
  while (right + 1 !== ctx.run.col && isHeader(header, right + 1)) right++;
  if (ctx.run.col >= left && ctx.run.col <= right) return null;

  const filled = (row: number) => {
    for (let c = left; c <= right; c++) if (!isBlank(at(row, c))) return true;
    return false;
  };
  // The Table runs from the header to the last row with data that the ranges reach. Blank rows
  // inside the data stay inside the Table, as they're inside the ranges today.
  const ends = longs.map((l) => l.area.rows![1].n);
  const end = Math.max(...ends);
  const bottom = origin.row + ctx.sheet.formulas.length - 1;
  let lastData = Math.min(end, bottom);
  while (lastData > header && !filled(lastData)) lastData--;
  if (lastData === header) return null;
  // Rows the ranges leave out are left out of the results today. A Table would take them in, so
  // results would change: data that carries straight on past the ranges, or past a shorter one.
  if ((end < bottom && filled(end + 1)) || lastData > Math.min(...ends)) return 'beyond';

  const names = new Set<string>();
  const headers: string[] = [];
  for (let c = left; c <= right; c++) {
    const h = String(at(header, c)).trim();
    if (names.has(h.toLowerCase())) return null;
    names.add(h.toLowerCase());
    headers[c] = h;
  }
  const name = tableNameFor(ctx.sheet.sheet);
  let out = body;
  for (const l of [...longs].reverse()) {
    const [c1, c2] = l.area.cols!;
    const ref = c1.n === c2.n ? `${name}[${structuredName(headers[c1.n])}]` : `${name}[[${structuredName(headers[c1.n])}]:[${structuredName(headers[c2.n])}]]`;
    out = out.slice(0, l.start) + ref + out.slice(l.end);
  }
  return {
    name,
    range: `${cellAddress({ row: header, col: left })}:${cellAddress({ row: lastData, col: right })}`,
    formula: `=${out}`,
  };
}

function planTable(ctx: FixContext): FixPlan | null {
  const affected = ctx.formulas.filter((c) => c.facts.longestRange >= LONG_RANGE_ROWS);
  const source = affected[0];
  if (!source) return null;
  const table = tableFor(ctx, source.formula);
  if (table === 'other-sheet' || table === 'beyond') return null;
  const forbid: FixRule[] = [
    {
      pattern: /\$?[A-Za-z]{1,3}\$?\d+:\$?[A-Za-z]{1,3}\$?\d+/,
      test: (f) => analyzeFormula(f).longestRange >= LONG_RANGE_ROWS,
      label: 'No fixed ranges over the long list',
      advice: 'Refer to the Table’s columns instead, so the formula grows with the data.',
    },
  ];
  const require: FixRule[] =
    affected.length === ctx.formulas.length
      ? [
          {
            pattern: /\[/,
            test: (f) => analyzeFormula(f).structured,
            label: 'Refers to the Table’s columns by name',
            advice: table ? `For example, ${source.address} becomes ${table.formula}.` : 'Write the column as TableName[Column], for example Sales[Amount].',
          },
        ]
      : [];
  const exact = table && clears(table.formula, forbid) && analyzeFormula(table.formula).structured ? table : null;
  const body = normalizeFormula(source.formula);
  const first = longRanges(body)[0];

  return plan(ctx, {
    title: 'Turn the data into a Table',
    intro: `${first ? `${mono(body.slice(first.start, first.end))} stops` : 'A fixed range stops'} at a fixed row, so rows added below it are left out of the result. A Table grows with its data, and its column names make the formula readable.`,
    steps: exact
      ? [
          `On the copy, select ${mono(exact.range)} and press {tableKey}. Keep My table has headers selected.`,
          `On the Table tab (Table Design on Windows), type ${mono(exact.name)} in the Table Name box and press {enter}.`,
          `In ${mono(source.address)}, replace the formula with ${mono(exact.formula)}. The column names stand in for the fixed ranges.`,
          fillStep(ctx, source, affected, 'formulas with fixed ranges'),
          checkStep(ctx),
        ]
      : [
          'On the copy, select the data the formula refers to, including its header row, and press {tableKey}.',
          `In ${mono(source.address)}, replace each fixed range with the Table’s column, like ${mono('Table1[Amount]')}.`,
          fillStep(ctx, source, affected, 'formulas with fixed ranges'),
          checkStep(ctx),
        ],
    forbid,
    require,
    copyBack: `Formulas that name a Table can’t be pasted back as they are: they’d point at the copy’s Table. To use this on ${mono(ctx.sheet.sheet)}, convert ${exact ? mono(exact.range) : 'the same data'} there to a Table too, then write the same formula with that Table’s name.`,
  });
}

interface Planner {
  /** Typed numbers between formulas belong to the block. */
  typed: boolean;
  plan(ctx: FixContext): FixPlan | null;
}

// vlookup-approximate and concatenated-keys have no guided fix: an exact-match rewrite can change
// results on purpose, and removing a helper column has nothing safe to compare.
const PLANNERS: Record<string, Planner> = {
  'vlookup-column-number': { typed: false, plan: planLookup },
  'iferror-lookup': { typed: false, plan: planNotFound },
  'nested-if': { typed: false, plan: planIfs },
  'sumproduct-conditions': { typed: false, plan: planSumifs },
  'typed-over-formula': { typed: true, plan: planTypedOver },
  'inconsistent-column': { typed: false, plan: planConsistent },
  'hardcoded-constants': { typed: false, plan: planInputCell },
  'volatile-references': { typed: false, plan: planVolatile },
  'fixed-long-ranges': { typed: false, plan: planTable },
};

/** Detector ids that have a guided fix. */
export const FIXABLE: ReadonlySet<string> = new Set<string>(Object.keys(PLANNERS));

/**
 * Builds a plan for the finding's first example cell on `sheet`, or null when no guided fix fits.
 * When the first example doesn't fit (an approximate VLOOKUP, say), later examples on the sheet are
 * tried in order, so a fixable column elsewhere still gets its plan.
 */
export function planFix(finding: Finding, sheet: SheetFormulas, options: FixOptions = {}): FixPlan | null {
  const planner = PLANNERS[finding.id];
  if (!planner) return null;
  for (const example of finding.cells) {
    if (example.sheet !== sheet.sheet) continue;
    let cell: CellRef;
    try {
      cell = parseCell(example.address);
    } catch {
      continue;
    }
    const run = findRun(sheet, cell, planner.typed, options.partial);
    const fix = run && planner.plan(contextFor(finding, sheet, run));
    if (fix) return fix;
  }
  return null;
}

/**
 * The cell just below the plan's block when the block ends on the last row that was scanned, or
 * null. The scan reads very large sheets in part, so the panel reads this cell: a formula there
 * means the column goes on, and planFix should run again with `partial`.
 */
export function scanEdgeBelow(plan: FixPlan, sheet: SheetFormulas): string | null {
  try {
    const block = parseRange(plan.range);
    const used = parseRange(sheet.address || 'A1');
    const lastScanned = used.start.row + (sheet.formulas?.length ?? 0) - 1;
    return block.end.row === lastScanned && block.end.row < 1048576 ? cellAddress({ row: block.end.row + 1, col: block.start.col }) : null;
  } catch {
    return null;
  }
}

// ---------- References to the copy itself ----------

/** The copy the check reads, so it can tell formulas that would break once the copy is deleted. */
export interface FixCopy {
  /** The copy's sheet name, e.g. "Fix-Orders". */
  name: string;
  /** The learner's sheet as scanned, for the Table names its formulas use. */
  original?: SheetFormulas;
}

/** What the copy's formulas in the block point at that the learner's sheet doesn't have. */
export interface CopyReferences {
  /** Cells that name the copy's own sheet, and the prefix as Excel wrote it, e.g. 'Fix-Orders'!. */
  sheet?: { cells: string[]; prefix: string };
  /** Tables Excel renamed when it copied the sheet (Items became Items2), with the cells that name each. */
  renamed: { name: string; original: string; cells: string[] }[];
  /**
   * Other Tables the copy's formulas name that no formula on the original names. One may be a Table
   * on the copy that the original's formulas never used; the panel asks the learner to check.
   */
  unknown: string[];
}

/** A sheet prefix in masked text; error values such as #REF! aren't sheets. */
const SHEET_REF = /(?<![A-Za-z0-9_.\]#])('[^']*'|[A-Za-z_][\w.]*)!/g;
/** A Table name in masked text, followed by its "[". */
const TABLE_REF = /(?<![A-Za-z0-9_.\\\]])([A-Za-z_\\][\w.]*)\[/g;

const tableNames = (code: string): string[] => [...code.matchAll(TABLE_REF)].map((m) => m[1]);
/** Excel numbers the name of a Table it copies (Items becomes Items2), so the stem leaves the number off. */
const stem = (name: string): string => name.replace(/\d+$/, '').toLowerCase();

/** Table names the sheet's formulas use, lower-cased to as written. A scanned sheet never changes, so it's read once. */
const tablesSeen = new WeakMap<SheetFormulas, Map<string, string>>();
function tablesNamedOn(sheet: SheetFormulas): Map<string, string> {
  let names = tablesSeen.get(sheet);
  if (names) return names;
  names = new Map();
  for (const row of sheet.formulas ?? []) {
    for (const v of row ?? []) if (isFormula(v)) for (const t of tableNames(maskFormula(normalizeFormula(v)))) names.set(t.toLowerCase(), t);
  }
  tablesSeen.set(sheet, names);
  return names;
}

/**
 * Finds formulas in the copy's block that point at the copy itself. Excel rewrites a sheet's
 * references to its own name when it copies the sheet ('Fix-Orders'!B2), and renames each Table on
 * it. Pasted back onto the original, such formulas still point at the copy, and show #REF! once
 * the copy is deleted.
 */
export function copyReferences(plan: FixPlan, after: RangeRead, copy: FixCopy): CopyReferences {
  const origin = parseRange(plan.range).start;
  const copyName = copy.name.toLowerCase();
  const known = copy.original ? tablesNamedOn(copy.original) : new Map<string, string>();

  const sheetCells: string[] = [];
  let prefix = '';
  const named = new Map<string, { name: string; cells: string[] }>();
  after.formulas.forEach((row, r) =>
    row.forEach((f, c) => {
      if (!isFormula(f)) return;
      const address = cellAddress({ row: origin.row + r, col: origin.col + c });
      const body = normalizeFormula(f);
      const code = maskFormula(body);
      for (const m of code.matchAll(SHEET_REF)) {
        const text = body.slice(m.index, m.index + m[0].length);
        if (prefixSheet(text).toLowerCase() !== copyName) continue;
        sheetCells.push(address);
        prefix ||= text;
        break;
      }
      for (const name of new Set(tableNames(code))) {
        if (known.has(name.toLowerCase())) continue;
        let entry = named.get(name.toLowerCase());
        if (!entry) named.set(name.toLowerCase(), (entry = { name, cells: [] }));
        entry.cells.push(address);
      }
    }),
  );

  const renamed: CopyReferences['renamed'] = [];
  const unknown: string[] = [];
  for (const { name, cells } of named.values()) {
    const from = /\d$/.test(name) ? [...known.values()].filter((k) => stem(k) === stem(name)) : [];
    const original = from.find((k) => k.toLowerCase() === stem(name)) ?? (from.length === 1 ? from[0] : undefined);
    if (original) renamed.push({ name, original, cells });
    // The Table fix makes a Table on the copy on purpose; its own copy-back step covers it.
    else if (plan.findingId !== 'fixed-long-ranges') unknown.push(name);
  }
  return { sheet: sheetCells.length ? { cells: sheetCells, prefix } : undefined, renamed, unknown };
}

/**
 * The formula with prefixes naming any of `sheets` dropped (on the sheet itself they change nothing)
 * and renamed Tables given their original names, so a formula on the copy compares with the original's.
 */
function comparable(formula: string, sheets: string[], tables: ReadonlyMap<string, string>): string {
  const body = normalizeFormula(formula);
  const code = maskFormula(body);
  const own = new Set(sheets.map((s) => s.toLowerCase()));
  const edits: { start: number; end: number; text: string }[] = [];
  for (const m of code.matchAll(SHEET_REF)) {
    const end = m.index + m[0].length;
    if (own.has(prefixSheet(body.slice(m.index, end)).toLowerCase())) edits.push({ start: m.index, end, text: '' });
  }
  for (const m of code.matchAll(TABLE_REF)) {
    const original = tables.get(m[1].toLowerCase());
    if (original) edits.push({ start: m.index, end: m.index + m[1].length, text: original });
  }
  edits.sort((a, b) => a.start - b.start);
  let out = '';
  let last = 0;
  for (const e of edits) {
    out += body.slice(last, e.start) + e.text;
    last = e.end;
  }
  return `=${out}${body.slice(last)}`;
}

/** The check item for formulas that point at the copy itself, or null when there are none. */
function copyRefsItem(plan: FixPlan, refs: CopyReferences): CheckItem | null {
  const label = 'Formulas don’t depend on the copy';
  const breaks = (cells: string[]) => `so ${cells.length === 1 ? 'it' : 'they'}’d show #REF! once you delete the copy`;
  const replace = (find: string, by: string) => `Select ${plan.range} on the copy, choose Home › Find & Select › Replace, and replace ${find} with ${by}.`;
  if (refs.sheet) {
    const { cells, prefix } = refs.sheet;
    return {
      id: 'fix-copy-refs',
      label,
      status: 'fail',
      detail: `${cellsPhrase(cells)} ${cells.length === 1 ? 'names' : 'name'} the copy’s own sheet, ${breaks(cells)}. ${replace(prefix, 'nothing')}`,
      focus: cells[0],
    };
  }
  const [table] = refs.renamed;
  if (!table) return null;
  const { cells, name, original } = table;
  return {
    id: 'fix-copy-refs',
    label,
    status: 'fail',
    detail: `${cellsPhrase(cells)} ${cells.length === 1 ? 'refers' : 'refer'} to ${name}, the copy’s own Table, ${breaks(cells)}. Change ${name} to ${original}, your sheet’s Table. ${replace(`${name}[`, `${original}[`)}`,
    focus: cells[0],
  };
}

// ---------- Grading ----------

/** A value as a learner would say it: blank, #N/A, 412.5 or “East”. */
function show(value: Cell): string {
  if (isBlank(value)) return 'blank';
  if (isErrorValue(value)) return value;
  return describeValue(value);
}

/** Before and after, with more digits when the usual rounding would make them look the same. */
function showPair(was: Cell, now: Cell): [string, string] {
  const pair: [string, string] = [show(was), show(now)];
  if (pair[0] === pair[1] && typeof was === 'number' && typeof now === 'number') return [formatNumber(was), formatNumber(now)];
  return pair;
}

/** "D9", "D9 and D14", or "D9 and 4 other cells". */
function cellsPhrase(cells: string[]): string {
  if (cells.length === 1) return cells[0];
  if (cells.length === 2) return `${cells[0]} and ${cells[1]}`;
  return `${cells[0]} and ${cells.length - 1} other cells`;
}

/**
 * Compares the original (before) with the copy (after). Passes when results match and the rules
 * hold. With `copy`, it also fails formulas that point at the copy itself, and reads formulas the
 * way they'd be on the original, so a Table Excel renamed on the copy doesn't count as a change.
 */
export function gradeFix(plan: FixPlan, before: RangeRead, after: RangeRead, copy?: FixCopy): CheckItem[] {
  const label = 'Results match the original';
  const { rows, cols } = rangeSize(plan.range);
  const fits = (g: Grid) => g.length === rows && g.every((row) => row.length === cols);
  if (!fits(before.values) || !fits(after.values) || !fits(after.formulas)) {
    return [
      {
        id: 'fix-values',
        label,
        status: 'fail',
        detail: `The coach read different sizes from the original and the copy, where both should cover ${plan.range}. Make a fresh copy and check again.`,
      },
    ];
  }

  const origin = parseRange(plan.range).start;
  const at = (r: number, c: number) => cellAddress({ row: origin.row + r, col: origin.col + c });
  const formulaAt = (g: Grid, address: string): Cell => {
    const { row, col } = parseCell(address);
    return g[row - origin.row]?.[col - origin.col] ?? '';
  };
  const items: CheckItem[] = [];
  const refs = copy && copyReferences(plan, after, copy);
  const renamed = new Map(refs?.renamed.map((t) => [t.name.toLowerCase(), t.original]));
  const copyKey = (f: string, a: string) => relativeKey(comparable(f, copy ? [copy.name, plan.sheet] : [plan.sheet], renamed), a);
  const originalKey = (f: string, a: string) => relativeKey(comparable(f, [plan.sheet], new Map()), a);

  // Values: every cell except the ones the fix is allowed to change, compared like answers.
  const mayChange = new Set(plan.valuesMayChange.map((a) => cellAddress(parseCell(a))));
  const settled = (g: Grid): Grid => g.map((row, r) => row.map((v, c) => (mayChange.has(at(r, c)) ? '' : v)));
  const mismatches = compareGrids(settled(after.values), settled(before.values));
  if (mismatches.length) {
    const { row, col } = mismatches[0];
    const address = at(row, col);
    const was = before.values[row][col];
    const now = after.values[row][col];
    const [a, b] = showPair(was, now);
    let detail = `${address} was ${a}, now ${b}.`;
    if (mismatches.length > 1) detail = `${mismatches.length} cells changed. ${detail}`;
    if (isErrorValue(now) && !isErrorValue(was)) detail += ` ${explainError(now)}`;
    // An IFERROR fix lets errors through that the fallback used to cover. They're real problems in
    // the data, and fixing them on the copy alone would only make it differ from the original.
    const fallback = plan.fallback;
    const hidden = (m: Mismatch) => fallback !== undefined && cellMatches(before.values[m.row][m.col], fallback) && isErrorValue(m.actual) && m.actual !== '#N/A';
    if (isErrorValue(now) && mismatches.every(hidden)) {
      const fresh = 'on your original sheet, then start over with a fresh copy.';
      detail =
        mismatches.length === 1
          ? `${address} now shows ${now}, which IFERROR was hiding. ${explainError(now)} Fix that ${fresh}`
          : `${address} and ${mismatches.length - 1} other ${plural(mismatches.length - 1, 'cell')} now show errors IFERROR was hiding, starting with ${now} in ${address}. Fix them ${fresh}`;
    }
    items.push({ id: 'fix-values', label, status: 'fail', detail, focus: address });
  } else {
    items.push({ id: 'fix-values', label, status: 'pass' });
  }

  const cells: { address: string; formula: Cell }[] = [];
  after.formulas.forEach((row, r) => row.forEach((formula, c) => cells.push({ address: at(r, c), formula })));
  const formulas = cells.filter((c): c is { address: string; formula: string } => isFormula(c.formula));

  if (plan.allFormulas) {
    const missing = cells.filter((c) => !isFormula(c.formula));
    const formulaLabel = `Every cell in ${plan.range} has a formula`;
    if (missing.length) {
      const first = missing[0];
      const what = isBlank(first.formula) ? 'is empty' : 'has a typed value';
      const detail = missing.length === 1 ? `${first.address} ${what}.` : `${missing.length} cells are empty or have typed values, starting at ${first.address}.`;
      items.push({ id: 'fix-formulas', label: formulaLabel, status: 'fail', detail, focus: first.address });
    } else {
      items.push({ id: 'fix-formulas', label: formulaLabel, status: 'pass' });
    }
  }

  if (plan.restore) {
    const { cells: repaired, model, key, label: restoreLabel, advice } = plan.restore;
    // The column's formula as the copy has it, so names Excel changed on the copy change it too.
    const modelFormula = formulaAt(after.formulas, model);
    const modelKey = isFormula(modelFormula) ? copyKey(modelFormula, model) : undefined;
    const wrong = repaired.filter((a) => {
      const f = formulaAt(after.formulas, a);
      return !isFormula(f) || (relativeKey(f, a) !== key && copyKey(f, a) !== modelKey);
    });
    items.push(
      wrong.length
        ? {
            id: 'fix-restore',
            label: restoreLabel,
            status: 'fail',
            detail: `${cellsPhrase(wrong)} ${wrong.length === 1 ? 'doesn’t' : 'don’t'} hold the column’s formula yet. ${advice}`,
            focus: wrong[0],
          }
        : { id: 'fix-restore', label: restoreLabel, status: 'pass' },
    );

    // Rows that build on the repaired cells can show new values, but their formulas are the
    // learner's own, deliberate sections included, and must stay as they were.
    const own = new Set(repaired.map((a) => cellAddress(parseCell(a))));
    const builders = plan.valuesMayChange.filter((a) => !own.has(cellAddress(parseCell(a))));
    if (builders.length) {
      const changed = builders.filter((a) => {
        const was = formulaAt(before.formulas, a);
        const now = formulaAt(after.formulas, a);
        return isFormula(was) && (!isFormula(now) || copyKey(now, a) !== originalKey(was, a));
      });
      const builderLabel = 'Rows that build on the repaired cells keep their formulas';
      const which = plan.findingId === 'typed-over-formula' ? 'typed cells' : 'cells that differed';
      if (changed.length) {
        const first = changed[0];
        const detail =
          changed.length === 1
            ? `${first} no longer has its formula from your sheet, ${formulaAt(before.formulas, first)}. Only the ${which} needed the column’s formula, so put it back.`
            : `${cellsPhrase(changed)} no longer have their formulas from your sheet. Only the ${which} needed the column’s formula, so copy the others back from your sheet.`;
        items.push({ id: 'fix-builders', label: builderLabel, status: 'fail', detail, focus: first });
      } else {
        items.push({ id: 'fix-builders', label: builderLabel, status: 'pass' });
      }
    }
  }

  plan.forbid.forEach((rule, i) => {
    const hits = formulas.filter((c) => matches(rule, c.formula)).map((c) => c.address);
    items.push(
      hits.length
        ? { id: `fix-forbid-${i}`, label: rule.label, status: 'fail', detail: `Found in ${cellsPhrase(hits)}. ${rule.advice}`, focus: hits[0] }
        : { id: `fix-forbid-${i}`, label: rule.label, status: 'pass' },
    );
  });

  plan.require.forEach((rule, i) => {
    const misses = formulas.filter((c) => !matches(rule, c.formula)).map((c) => c.address);
    const id = `fix-require-${i}`;
    if (!formulas.length) items.push({ id, label: rule.label, status: 'fail', detail: `${plan.range} has no formulas yet. ${rule.advice}` });
    else if (misses.length) items.push({ id, label: rule.label, status: 'fail', detail: `Not yet in ${cellsPhrase(misses)}. ${rule.advice}`, focus: misses[0] });
    else items.push({ id, label: rule.label, status: 'pass' });
  });

  // Listed only when it fails: a copy that points at itself is rare, and it's a fix, not a goal.
  const selfRefs = refs && copyRefsItem(plan, refs);
  if (selfRefs) items.push(selfRefs);

  return items;
}

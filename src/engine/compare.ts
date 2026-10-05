import type { Cell, CellMatcher, ExpectedCell, ExpectedGrid, Grid } from './types';

const ERROR_RE = /^#(NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A|SPILL!|CALC!|GETTING_DATA|FIELD!|BLOCKED!|CONNECT!|BUSY!|UNKNOWN!|PYTHON!|TEXT!)$/;

export function isMatcher(cell: ExpectedCell): cell is CellMatcher {
  return typeof cell === 'object' && cell !== null && 'match' in cell;
}

export function isErrorValue(value: Cell): value is string {
  return typeof value === 'string' && ERROR_RE.test(value);
}

export function isBlank(value: Cell | undefined): boolean {
  return value === null || value === undefined || value === '';
}

/** Excel stores doubles; allow tiny float noise and nothing more. */
export function sameNumber(a: number, b: number): boolean {
  return Math.abs(a - b) <= 1e-7 * Math.max(1, Math.abs(b));
}

export function cellMatches(actual: Cell, expected: ExpectedCell): boolean {
  if (isMatcher(expected)) return typeof actual === 'string' && expected.match.test(actual);
  if (isBlank(expected)) return isBlank(actual);
  if (typeof expected === 'number') return typeof actual === 'number' && sameNumber(actual, expected);
  if (typeof expected === 'boolean') return actual === expected;
  return typeof actual === 'string' && actual === expected;
}

export interface Mismatch {
  row: number;
  col: number;
  actual: Cell;
  expected: ExpectedCell;
}

/**
 * Compares `actual` against `expected`. `actual` may be larger than `expected`
 * (for spills we read one extra row and column); any extra cell must be blank.
 */
export function compareGrids(actual: Grid, expected: ExpectedGrid): Mismatch[] {
  const out: Mismatch[] = [];
  const rows = Math.max(actual.length, expected.length);
  for (let r = 0; r < rows; r++) {
    const cols = Math.max(actual[r]?.length ?? 0, expected[r]?.length ?? 0);
    for (let c = 0; c < cols; c++) {
      const a = actual[r]?.[c] ?? '';
      const inExpected = r < expected.length && c < (expected[r]?.length ?? 0);
      const e: ExpectedCell = inExpected ? expected[r][c] : '';
      if (!cellMatches(a, e)) out.push({ row: r, col: c, actual: a, expected: e });
    }
  }
  return out;
}

const numberFormat = new Intl.NumberFormat('en-US', { maximumFractionDigits: 4 });

export function describeValue(value: ExpectedCell | undefined): string {
  if (value === undefined || isBlank(value as Cell)) return 'a blank cell';
  if (isMatcher(value)) return value.describe;
  if (typeof value === 'number') return numberFormat.format(value);
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
  return `“${value}”`;
}

/** Plain-language explanation for an Excel error value. */
export function explainError(value: string): string {
  switch (value) {
    case '#NAME?':
      return 'Excel doesn’t recognize a name in the formula. Check spelling of functions, Table names and named ranges.';
    case '#SPILL!':
      return 'The result can’t spill because something is in the way. Clear the cells it needs.';
    case '#N/A':
      return 'A lookup didn’t find a match.';
    case '#VALUE!':
      return 'A value has the wrong type, often text where a number is expected.';
    case '#REF!':
      return 'The formula points at a cell that no longer exists.';
    case '#DIV/0!':
      return 'The formula divides by zero or by an empty cell.';
    case '#CALC!':
      return 'Excel couldn’t calculate a result. Common causes: a FILTER with no matches, a whole list where one value belongs, or a LAMBDA that’s never called.';
    default:
      return 'Excel returned an error.';
  }
}

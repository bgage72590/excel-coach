import { cellAddress, parseCell, parseRange } from './address';
import { cellMatches, describeValue, explainError, isErrorValue } from './compare';
import { isFormula } from './formula';
import type { AnswerRead } from './grade';
import type { AnswerArea, Cell, CellMark, CheckItem, ExpectedCell, ExpectedGrid } from './types';

/** One read of the bug-hunt range: the current data first, then each variant in order. */
export interface BugHuntRun {
  /** '' for the current data; the variant's label otherwise ("the tax rate changes"). */
  label: string;
  explain?: string;
  read: AnswerRead;
  expected: ExpectedGrid;
}

export interface BugHuntGrade {
  items: CheckItem[];
  /** Cells to paint. Must never reveal where an unfixed bug is: only regressions (red) and, on a full pass, every cell (green). */
  marks: CellMark[];
}

/** The first run in which a cell stops matching. */
interface Miss {
  run: BugHuntRun;
  actual: Cell;
  expected: ExpectedCell;
}

const plural = (n: number, one: string, many: string) => (n === 1 ? one : many);

/** "K5 was right and now shows 1,240; expected 1,310 when the handling fee changes." */
function regressionSentence(address: string, miss: Miss): string {
  const when = miss.run.label ? ` when ${miss.run.label}` : '';
  const shown = isErrorValue(miss.actual) ? miss.actual : describeValue(miss.actual);
  const base = `${address} was right and now shows ${shown}; expected ${describeValue(miss.expected)}${when}.`;
  return isErrorValue(miss.actual) ? `${base} ${explainError(miss.actual)}` : base;
}

/** How many mistakes are left and how many of them hide until the data changes. Never where they are. */
function remainingDetail(left: number, hidden: number): string {
  const head = left === 1 ? '1 mistake is still in the report.' : `${left} mistakes are still in the report.`;
  let tail: string;
  if (left === 1) tail = hidden ? 'It shows only when the data changes.' : 'It shows on the current data.';
  else if (hidden === 0) tail = 'All of them show on the current data.';
  else if (hidden === left) tail = 'All of them show only when the data changes.';
  else tail = `${hidden} of them ${plural(hidden, 'shows', 'show')} only when the data changes.`;
  const tip = hidden ? ' Change an input and watch which results don’t follow it. The coach puts the original inputs back when you check.' : '';
  return `${head} ${tail}${tip}`;
}

/**
 * Grades a bug hunt. `runs[0]` is the current data. A bug is fixed when every one of its cells
 * matches in every run. Cells that belong to no bug must match in every run (no regressions), and
 * every cell in the range that the coach wrote as a formula must still hold a formula.
 *
 * A cell without a formula inside an unfixed bug is either the planted mistake itself (a pasted
 * value) or the learner's typed-over attempt; either way the bug stays unfixed, and the formulas
 * item counts it without naming it, so no check ever points at a mistake that is still there.
 */
export function gradeBugHunt(area: Extract<AnswerArea, { kind: 'bugHunt' }>, runs: BugHuntRun[]): BugHuntGrade {
  const range = parseRange(area.range);
  const rows = range.end.row - range.start.row + 1;
  const cols = range.end.col - range.start.col + 1;
  const at = (r: number, c: number) => cellAddress({ row: range.start.row + r, col: range.start.col + c });

  const owner = new Map<string, number>();
  area.bugs.forEach((bug, i) => bug.cells.forEach((cell) => owner.set(cellAddress(parseCell(cell)), i)));

  const order: string[] = [];
  const noFormula = new Set<string>();
  const misses = new Map<string, Miss>();
  const missesNow = new Set<string>();
  const missesLater = new Set<string>();
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      const address = at(r, c);
      order.push(address);
      if (runs.length && !isFormula(runs[0].read.formulas[r]?.[c])) noFormula.add(address);
      for (const [i, run] of runs.entries()) {
        const actual = run.read.values[r]?.[c] ?? '';
        const expected = run.expected[r]?.[c] ?? '';
        if (cellMatches(actual, expected)) continue;
        if (!misses.has(address)) misses.set(address, { run, actual, expected });
        (i === 0 ? missesNow : missesLater).add(address);
      }
    }
  }

  const items: CheckItem[] = [];

  // Every report cell still has a formula. Only cells outside the mistakes are named.
  const typed = order.filter((a) => noFormula.has(a));
  const nameable = typed.filter((a) => !owner.has(a));
  const formulasLabel = 'Every report cell still has a formula';
  if (!typed.length) {
    items.push({ id: 'formulas', label: formulasLabel, status: 'pass' });
  } else {
    const fix = typed.length === 1 ? ' It needs a formula that keeps up when the data changes.' : ' Each one needs a formula that keeps up when the data changes.';
    const where = nameable.length
      ? typed.length === 1
        ? `${nameable[0]} is empty or has a typed value.`
        : `${typed.length} cells are empty or have typed values, including ${nameable[0]}.`
      : typed.length === 1
        ? 'One cell in the report is empty or has a typed value.'
        : `${typed.length} cells in the report are empty or have typed values.`;
    items.push({ id: 'formulas', label: formulasLabel, status: 'fail', detail: `${where}${fix}`, ...(nameable.length ? { focus: nameable[0] } : {}) });
  }

  // Cells that belong to no mistake must keep matching.
  const regressions = order.filter((a) => !owner.has(a) && misses.has(a));
  if (regressions.length) {
    const first = regressions[0];
    const sentence = regressionSentence(first, misses.get(first)!);
    // Every variant writes the coach's inputs back first, so cells that are wrong only on the sheet
    // as it stands most likely read an input the learner changed to test the report. The host
    // restores the inputs once the check is done.
    const inputChanged = runs.length > 1 && regressions.every((a) => !missesLater.has(a));
    const tip = inputChanged ? ' If you changed an input to test the report, the coach has put it back. Check again.' : '';
    items.push({
      id: 'regression',
      label: 'The rest of the report still matches',
      status: 'fail',
      detail: `${regressions.length === 1 ? sentence : `${regressions.length} cells that were right no longer match. ${sentence}`}${tip}`,
      focus: first,
    });
  }

  const inRange = (cell: string) => {
    const ref = parseCell(cell);
    return ref.row >= range.start.row && ref.row <= range.end.row && ref.col >= range.start.col && ref.col <= range.end.col;
  };
  const fixed = area.bugs.map((bug) =>
    bug.cells.every((cell) => {
      const address = cellAddress(parseCell(cell));
      return !inRange(address) || (!misses.has(address) && !noFormula.has(address));
    }),
  );
  area.bugs.forEach((bug, i) => {
    if (fixed[i]) items.push({ id: `bug-${bug.id}`, label: `Fixed: ${bug.label}`, status: 'pass' });
  });

  const left = fixed.filter((f) => !f).length;
  if (left) {
    // A mistake hides when all of its cells match on the current data.
    const hidden = area.bugs.filter((bug, i) => !fixed[i] && bug.cells.every((cell) => !missesNow.has(cellAddress(parseCell(cell))))).length;
    items.push({
      id: 'bugs-left',
      label: `${area.bugs.length - left} of ${area.bugs.length} mistakes fixed`,
      status: 'fail',
      detail: remainingDetail(left, hidden),
    });
  }

  const marks: CellMark[] = regressions.length ? regressions.map((address) => ({ address, ok: false })) : [];
  const allGood = !typed.length && !regressions.length && !left;
  return { items, marks: allGood ? order.map((address) => ({ address, ok: true })) : marks };
}

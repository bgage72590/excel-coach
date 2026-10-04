import { offsetCell } from './address';
import { compareGrids, describeValue, explainError, isErrorValue, type Mismatch } from './compare';
import { disallowedNumbers, forbiddenStrings, isFormula } from './formula';
import type { AnswerArea, Cell, CellCheck, CellMark, CheckItem, ExpectedGrid, Grid, Rules } from './types';

/** What the host read back from the answer area. */
export interface AnswerRead {
  /** Top-left-anchored address of the block that was read, without sheet name. */
  address: string;
  formulas: Grid;
  values: Grid;
  r1c1: Grid;
}

interface FormulaCell {
  address: string;
  formula: string;
  r1c1: string;
  row: number;
  col: number;
}

function formulaCells(read: AnswerRead, rows: number, cols: number): FormulaCell[] {
  const out: FormulaCell[] = [];
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      const f = read.formulas[r]?.[c];
      if (isFormula(f)) out.push({ address: offsetCell(read.address, r, c), formula: f, r1c1: String(read.r1c1[r]?.[c] ?? f), row: r, col: c });
    }
  }
  return out;
}

/** Checks that formulas exist where they should, and that they're the same formula throughout. */
export function gradeStructure(area: AnswerArea, read: AnswerRead, expectedRows: number, expectedCols: number): { items: CheckItem[]; formulas: string[] } {
  const items: CheckItem[] = [];
  const cells = formulaCells(read, expectedRows, expectedCols);

  if (area.kind === 'spill') {
    const anchorFormula = read.formulas[0]?.[0];
    if (!isFormula(anchorFormula)) {
      items.push({
        id: 'formula',
        label: `One formula in ${area.anchor} returns the whole result`,
        status: 'fail',
        detail: `${area.anchor} doesn’t have a formula yet.`,
        focus: area.anchor,
      });
      return { items, formulas: [] };
    }
    const extra = cells.filter((c) => c.address !== offsetCell(read.address, 0, 0));
    items.push(
      extra.length
        ? {
            id: 'formula',
            label: `One formula in ${area.anchor} returns the whole result`,
            status: 'fail',
            detail: `${extra[0].address} has its own formula. Delete it and let ${area.anchor} spill into it.`,
            focus: extra[0].address,
          }
        : { id: 'formula', label: `One formula in ${area.anchor} returns the whole result`, status: 'pass' },
    );
    return { items, formulas: [anchorFormula] };
  }

  // Data tables: Office.js reports their cells as values, never =TABLE(...). So a visible formula
  // means the grid was built by hand; values are confirmed live by the variants that follow.
  if (area.kind === 'cells' && area.liveValues) {
    const label = 'Built with Data › What-If Analysis › Data Table';
    items.push(
      cells.length
        ? { id: 'formula', label, status: 'fail', detail: `${cells[0].address} has its own formula. Clear the grid, select the whole table including the corner cell, and use Data › What-If Analysis › Data Table.`, focus: cells[0].address }
        : { id: 'formula', label, status: 'pass' },
    );
    return { items, formulas: [] };
  }

  // cells / tableColumn
  const missing: string[] = [];
  for (let r = 0; r < expectedRows; r++) {
    for (let c = 0; c < expectedCols; c++) {
      if (isFormula(read.formulas[r]?.[c])) continue;
      const spilledFromLeft =
        area.kind === 'cells' && area.spillOk && Array.from({ length: c }, (_, k) => read.formulas[r]?.[k]).some(isFormula);
      if (!spilledFromLeft) missing.push(offsetCell(read.address, r, c));
    }
  }
  items.push(
    missing.length
      ? {
          id: 'formula',
          label: 'Every answer cell has a formula',
          status: 'fail',
          detail:
            missing.length === 1
              ? `${missing[0]} is empty or has a typed value.`
              : `${missing.length} cells are empty or have typed values, starting at ${missing[0]}.`,
          focus: missing[0],
        }
      : { id: 'formula', label: 'Every answer cell has a formula', status: 'pass' },
  );

  const consistency = area.kind === 'cells' ? area.consistency : area.kind === 'tableColumn' ? 'all' : 'none';
  if (consistency !== 'none' && cells.length > 1) {
    const groups = new Map<string, FormulaCell[]>();
    for (const cell of cells) {
      const key = consistency === 'columns' ? String(cell.col) : consistency === 'rows' ? String(cell.row) : 'all';
      groups.set(key, [...(groups.get(key) ?? []), cell]);
    }
    let odd: FormulaCell | undefined;
    let oddCommon: string | undefined;
    for (const group of groups.values()) {
      const counts = new Map<string, number>();
      for (const c of group) counts.set(c.r1c1, (counts.get(c.r1c1) ?? 0) + 1);
      const common = [...counts.entries()].sort((a, b) => b[1] - a[1])[0]?.[0];
      const found = group.find((c) => c.r1c1 !== common);
      if (!odd && found) {
        odd = found;
        oddCommon = common;
      }
    }
    const label =
      consistency === 'columns'
        ? 'One formula filled down each column'
        : consistency === 'rows'
          ? 'One formula filled across each row'
          : expectedCols > 1 && expectedRows > 1
            ? 'One formula filled across and down'
            : 'One formula filled through the range';
    items.push(
      odd
        ? {
            id: 'consistent',
            label,
            status: 'fail',
            detail: tableColumnsSlid(odd.r1c1, oddCommon)
              ? `${odd.address} names different Table columns than the cells around it. Dragging the fill handle right slides Table column names along; fill with Home › Fill › Right instead, or write each column as Table[[Column]:[Column]].`
              : `${odd.address} doesn’t match the formula in the cells around it. Write it once, anchor it with $, and fill.`,
            focus: odd.address,
          }
        : { id: 'consistent', label, status: 'pass' },
    );
  }

  return { items, formulas: [...new Set(cells.map((c) => c.formula))] };
}

/** Applies required patterns and literal rules to the learner's formulas. */
export function gradeRules(rules: Rules | undefined, formulas: string[]): CheckItem[] {
  if (!rules || formulas.length === 0) return [];
  const items: CheckItem[] = [];

  for (const [i, req] of (rules.require ?? []).entries()) {
    const ok = formulas.every((f) => req.pattern.test(f));
    items.push({ id: `require-${i}`, label: req.label, status: ok ? 'pass' : 'fail', detail: ok ? undefined : req.advice });
  }

  if (rules.forbidText) {
    const found = formulas.flatMap((f) => forbiddenStrings(f, rules.forbidText!.values));
    items.push(
      found.length
        ? { id: 'no-typed-text', label: 'Criteria come from cells, not typed text', status: 'fail', detail: `The formula types “${found[0]}”. ${rules.forbidText.advice}` }
        : { id: 'no-typed-text', label: 'Criteria come from cells, not typed text', status: 'pass' },
    );
  }

  if (rules.allowNumbers) {
    const found = formulas.flatMap((f) => disallowedNumbers(f, rules.allowNumbers!));
    items.push(
      found.length
        ? {
            id: 'no-typed-numbers',
            label: 'No typed-in numbers',
            status: 'fail',
            detail: `The formula contains ${found[0]}. Point to the cell that holds it, so the result updates when the data changes.`,
          }
        : { id: 'no-typed-numbers', label: 'No typed-in numbers', status: 'pass' },
    );
  }
  return items;
}

export interface ValueGrade {
  item: CheckItem;
  mismatches: Mismatch[];
  marks: CellMark[];
}

function mismatchSentence(read: AnswerRead, m: Mismatch): { address: string; text: string } {
  const address = offsetCell(read.address, m.row, m.col);
  if (isErrorValue(m.actual)) return { address, text: `${address} shows ${m.actual}. ${explainError(m.actual)}` };
  return { address, text: `${address} shows ${describeValue(m.actual)}; expected ${describeValue(m.expected)}.` };
}

/** True when two R1C1 formulas differ only inside structured-reference brackets: Table column names that slid during a fill. */
export function tableColumnsSlid(a: string, b: string | undefined): boolean {
  if (b === undefined || a === b || !a.includes('[') || !b.includes('[')) return false;
  const shape = (f: string) => f.replace(/\[[^[\]]*\]/g, '[]');
  return shape(a) === shape(b);
}

/**
 * Spots the classic fill-down mistake: a range without $ that slides as the formula fills
 * (F2:F49 becomes F3:F50). Returns an explanation, or undefined. With the answer's shape, a
 * one-row range in a column filled down (each row summing its own row, like B4:D4) isn't flagged,
 * and neither is a one-column range in a row filled across.
 */
export function slidingRangeHint(formulas: string[], answerCells: number, shape?: { rows: number; cols: number }): string | undefined {
  if (answerCells < 2) return undefined;
  for (const f of formulas) {
    const code = f.replace(/"(?:[^"]|"")*"/g, '');
    for (const m of code.matchAll(/(?<![$A-Za-z])([A-Za-z]{1,3})(\d+):([A-Za-z]{1,3})(\d+)(?![\d(])/g)) {
      const oneRow = m[2] === m[4];
      const oneCol = m[1].toUpperCase() === m[3].toUpperCase();
      if (shape && ((oneRow && shape.cols === 1) || (oneCol && shape.rows === 1))) continue;
      const shifted = `${m[1]}${Number(m[2]) + 1}:${m[3]}${Number(m[4]) + 1}`;
      return `Your range ${m[0]} moves as the formula fills (one row down it becomes ${shifted}). Lock it with $, or use the Table’s columns instead.`;
    }
  }
  return undefined;
}

/** Compares values against the expected grid on the current data. */
export function gradeValues(read: AnswerRead, expected: ExpectedGrid, formulas: string[] = []): ValueGrade {
  const mismatches = compareGrids(read.values, expected);
  const marks: CellMark[] = [];
  const bad = new Set(mismatches.map((m) => `${m.row}:${m.col}`));
  expected.forEach((row, r) => row.forEach((_, c) => marks.push({ address: offsetCell(read.address, r, c), ok: !bad.has(`${r}:${c}`) })));
  for (const m of mismatches) {
    if (m.row >= expected.length || m.col >= (expected[0]?.length ?? 0)) marks.push({ address: offsetCell(read.address, m.row, m.col), ok: false });
  }

  if (mismatches.length === 0) return { item: { id: 'values', label: 'Correct on the current data', status: 'pass' }, mismatches, marks };

  const first = mismatchSentence(read, mismatches[0]);
  const outside = mismatches[0].row >= expected.length || mismatches[0].col >= (expected[0]?.length ?? 0);
  const base = outside
    ? `The result is bigger than expected: ${first.address} should be empty.`
    : mismatches.length === 1
      ? first.text
      : `${mismatches.length} cells don’t match. ${first.text}`;
  const why = slidingRangeHint(formulas, expected.length * (expected[0]?.length ?? 1), { rows: expected.length, cols: expected[0]?.length ?? 1 });
  const detail = why ? `${base} ${why}` : base;
  return { item: { id: 'values', label: 'Correct on the current data', status: 'fail', detail, focus: first.address }, mismatches, marks };
}

/** Result of one perturbation trial. */
export function gradeVariant(
  index: number,
  label: string,
  explain: string | undefined,
  read: AnswerRead,
  expected: ExpectedGrid,
): { item: CheckItem; badCells: string[] } {
  const mismatches = compareGrids(read.values, expected);
  const id = `variant-${index}`;
  const fullLabel = `Still correct when ${label}`;
  if (mismatches.length === 0) return { item: { id, label: fullLabel, status: 'pass' }, badCells: [] };
  const m = mismatches[0];
  const address = offsetCell(read.address, m.row, m.col);
  const shown = isErrorValue(m.actual) ? m.actual : describeValue(m.actual);
  const detail = `${address} showed ${shown} instead of ${describeValue(m.expected)}.${explain ? ` ${explain}` : ''}`;
  return {
    item: { id, label: fullLabel, status: 'fail', detail, focus: address },
    badCells: mismatches.map((x) => offsetCell(read.address, x.row, x.col)),
  };
}

// ---------- Power Query output tables ----------

/**
 * Compares a query's output rows with the expected rows. With order 'any', rows are matched as a
 * multiset. Returns a single check item.
 */
export function gradeQueryRows(actual: Grid, expected: ExpectedGrid, order: 'any' | 'asis'): CheckItem {
  const label = 'The output has the right rows';
  if (actual.length !== expected.length) {
    return {
      id: 'values',
      label,
      status: 'fail',
      detail: `The output has ${actual.length} row${actual.length === 1 ? '' : 's'}; it should have ${expected.length}. Check for rows you didn’t remove or didn’t keep.`,
    };
  }
  if (order === 'asis') {
    const mismatches = compareGrids(actual, expected);
    if (!mismatches.length) return { id: 'values', label, status: 'pass' };
    const m = mismatches[0];
    return { id: 'values', label, status: 'fail', detail: `Row ${m.row + 1} shows ${describeValue(m.actual)} where ${describeValue(m.expected)} was expected.` };
  }
  // Order-insensitive: each expected row must be matched by one actual row.
  const pool = actual.map((r) => ({ row: r, used: false }));
  for (const exp of expected) {
    const hit = pool.find((p) => !p.used && compareGrids([p.row], [exp]).length === 0);
    if (!hit) {
      const shown = exp.map((c) => describeValue(c)).join(', ');
      return { id: 'values', label, status: 'fail', detail: `A row is missing or different: ${shown}.` };
    }
    hit.used = true;
  }
  return { id: 'values', label, status: 'pass' };
}


// ---------- single-cell checks (Goal Seek, Solver, constraints) ----------

const squash = (f: string) => f.replace(/\s+/g, '').replace(/_xlfn\.|_xlpm\./gi, '').toUpperCase();

/** Grades one cell against its CellCheck. `formula` is what Excel reports in .formulas. */
export function gradeCellCheck(check: CellCheck, value: Cell, formula: Cell): CheckItem {
  const id = `cell-${check.cell}`;
  const fail = (detail: string): CheckItem => ({ id, label: check.label, status: 'fail', detail: check.advice ? `${detail} ${check.advice}` : detail, focus: check.cell });
  const hasFormula = isFormula(formula);

  if (check.holds === 'number' && (hasFormula || typeof value !== 'number')) {
    return fail(hasFormula ? `${check.cell} has a formula; it should hold a typed number.` : `${check.cell} should hold a number.`);
  }
  if (check.holds === 'formula' && !hasFormula) return fail(`${check.cell} should contain a formula.`);
  if (check.formula !== undefined && (!hasFormula || squash(String(formula)) !== squash(check.formula))) {
    return fail(`${check.cell} should still contain ${check.formula}. Undo any change to it.`);
  }
  const needsNumber = check.value !== undefined || check.min !== undefined || check.max !== undefined || check.integer;
  if (needsNumber && typeof value !== 'number') {
    return fail(isErrorValue(value) ? `${check.cell} shows ${value}. ${explainError(value)}` : `${check.cell} should show a number.`);
  }
  const n = value as number;
  const tol = check.tolerance ?? 0.005;
  if (check.value !== undefined && Math.abs(n - check.value) > tol) return fail(`${check.cell} shows ${describeValue(n)}; it should be ${describeValue(check.value)}.`);
  if (check.min !== undefined && n < check.min - 1e-9) return fail(`${check.cell} shows ${describeValue(n)}; it must be at least ${describeValue(check.min)}.`);
  if (check.max !== undefined && n > check.max + 1e-9) return fail(`${check.cell} shows ${describeValue(n)}; it must be at most ${describeValue(check.max)}.`);
  if (check.integer && Math.abs(n - Math.round(n)) > 1e-6) return fail(`${check.cell} shows ${describeValue(n)}; it must be a whole number.`);
  return { id, label: check.label, status: 'pass' };
}

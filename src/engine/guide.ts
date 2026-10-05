import { offsetCell, parseCell, parseRange } from './address';
import { cellMatches, compareGrids, describeValue, explainError, isBlank, isErrorValue } from './compare';
import { isFormula } from './formula';
import type { AnswerRead } from './grade';
import type { AnswerArea, Cell, ExpectedCell, ExpectedGrid, FormulaPart } from './types';

/** What the coach sees for a walkthrough step. */
export interface StepProbe {
  done: boolean;
  /** Something is there but not right yet, e.g. "I2 shows #CALC!, …". */
  note?: string;
  /** The cell the note is about. */
  focus?: string;
  /** The learner's formula in that cell, to set beside the walkthrough's. */
  formula?: string;
}

export const WAITING: StepProbe = { done: false };

/** The whole formula a walkthrough step spells out. */
export function partsFormula(parts: FormulaPart[]): string {
  return parts.map((p) => p.text).join('');
}

/** True when `cell` lies inside `range`, which may list several areas ("F6,F9:F12"). */
export function cellInRange(cell: string, range: string): boolean {
  const c = parseCell(cell);
  return range.split(',').some((area) => {
    const r = parseRange(area.trim());
    return c.row >= r.start.row && c.row <= r.end.row && c.col >= r.start.col && c.col <= r.end.col;
  });
}

function wrong(address: string, actual: Cell, expected: ExpectedCell, formula?: string): StepProbe {
  const note = isErrorValue(actual)
    ? `${address} shows ${actual}. ${explainError(actual)}`
    : isBlank(actual)
      ? `${address} is blank, but it should show ${describeValue(expected)}.`
      : `${address} shows ${describeValue(actual)}, but it should show ${describeValue(expected)}.`;
  return { done: false, note, focus: address, formula };
}

const typedValue = (address: string): StepProbe => ({
  done: false,
  note: `${address} holds a typed value. Start with = so Excel treats it as a formula.`,
  focus: address,
});

/**
 * Judges an 'answer' step from the answer area as the host read it (the same read the checker
 * uses). `cells` narrows a 'cells' area to part of it, e.g. the first cell before filling down.
 * Blank cells mean "not yet" and get no note; a formula showing the wrong value gets one.
 */
export function probeAnswer(area: AnswerArea, expected: ExpectedGrid, read: AnswerRead, cells?: string): StepProbe {
  if (area.kind === 'cells') {
    const origin = parseRange(area.range).start;
    const target = parseRange(cells ?? area.range);
    let pending = false;
    for (let row = target.start.row; row <= target.end.row; row++) {
      for (let col = target.start.col; col <= target.end.col; col++) {
        const r = row - origin.row;
        const c = col - origin.col;
        const address = offsetCell(area.range, r, c);
        const f = read.formulas[r]?.[c] ?? '';
        const v = read.values[r]?.[c] ?? '';
        const e = expected[r]?.[c] ?? '';
        if (area.liveValues || (area.spillOk && !isFormula(f) && !isBlank(v))) {
          // A what-if data table reports plain values; a spill from the left fills cells without formulas.
          if (isBlank(v)) pending = true;
          else if (!cellMatches(v, e)) return wrong(address, v, e);
          continue;
        }
        if (!isFormula(f)) {
          if (isBlank(f)) pending = true;
          else return typedValue(address);
          continue;
        }
        if (!cellMatches(v, e)) return wrong(address, v, e, f);
      }
    }
    return pending ? WAITING : { done: true };
  }

  if (area.kind === 'spill') {
    const f = read.formulas[0]?.[0] ?? '';
    if (!isFormula(f)) return isBlank(f) ? WAITING : typedValue(area.anchor);
    const mismatches = compareGrids(read.values, expected);
    if (!mismatches.length) return { done: true };
    const m = mismatches[0];
    const address = offsetCell(read.address, m.row, m.col);
    const outside = m.row >= expected.length || m.col >= (expected[0]?.length ?? 0);
    if (outside) return { done: false, note: `The result is bigger than expected: ${address} should be empty.`, focus: address, formula: f };
    return wrong(address, m.actual, m.expected, f);
  }

  if (area.kind === 'tableColumn') {
    const first = read.formulas[0]?.[0] ?? '';
    if (!read.values.length || isBlank(first)) return WAITING;
    if (!isFormula(first)) return typedValue(offsetCell(read.address, 0, 0));
    for (let r = 0; r < read.values.length; r++) {
      const v = read.values[r]?.[0] ?? '';
      const e = expected[r]?.[0] ?? '';
      if (!cellMatches(v, e)) return wrong(offsetCell(read.address, r, 0), v, e, String(read.formulas[r]?.[0] ?? first));
    }
    return { done: true };
  }

  return WAITING;
}

/** A SheetSpot taken apart: a Table reference, or A1 ranges (possibly several). */
export type ParsedSpot =
  | { kind: 'table'; table: string; part: 'data' | 'all' | 'headers' }
  | { kind: 'column'; table: string; column: string }
  | { kind: 'range'; address: string; areas: number };

export function parseSpot(spot: string): ParsedSpot {
  const m = /^([A-Za-z_\\][\w.\\]*)\[(.+)\]$/.exec(spot.trim());
  if (m) {
    const inner = m[2].replace(/^\[(.*)\]$/, '$1');
    const special = /^#(data|all|headers)$/i.exec(inner);
    if (special) return { kind: 'table', table: m[1], part: special[1].toLowerCase() as 'data' | 'all' | 'headers' };
    return { kind: 'column', table: m[1], column: inner.replace(/'(.)/g, '$1') };
  }
  const address = spot.replace(/\s+/g, '');
  return { kind: 'range', address, areas: address.split(',').length };
}

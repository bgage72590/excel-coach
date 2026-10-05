import type { FormulaPart, GuideStep, SheetPointer, SheetSpot } from '../engine/types';

// Building blocks for walkthroughs. A walkthrough is a few small steps: look at the data, see what
// the answer is made of, type the formula, fill it, check. Every step says exactly what to click or
// type, and the coach ticks it off by itself when it sees the result in Excel.

/** A formula part with what it means and, optionally, where it is on the sheet. */
export function part(text: string, means: string, at?: SheetSpot): FormulaPart {
  return at ? { text, means, at } : { text, means };
}

/** Punctuation, or anything that needs no explanation: ", " and ")". */
export function raw(text: string): FormulaPart {
  return { text };
}

/** "Click `I2` and type this formula, then press Return." Ticks off when the cell shows the right value. */
export function typeStep(o: {
  cell: string;
  formula: FormulaPart[];
  why?: string;
  show?: SheetPointer[];
  /** Judge the whole answer area instead of just `cell` (spills, Table columns, single-cell answers). */
  whole?: boolean;
  /** Override the instruction. */
  do?: string;
}): GuideStep {
  return {
    do: o.do ?? `Click \`${o.cell}\` and type this formula, then press {enter}.`,
    formula: o.formula,
    why: o.why,
    show: o.show,
    done: o.whole ? { kind: 'answer' } : { kind: 'answer', cells: o.cell },
  };
}

/** Copies the first cell's formula over the rest of the answer area with Fill Down/Right. */
export function fillStep(o: { from: string; range: string; direction: 'down' | 'right' | 'both'; why?: string }): GuideStep {
  const how =
    o.direction === 'down'
      ? `Select \`${o.range}\` (click \`${o.from}\`, then Shift-click the last cell) and press {fillDown}.`
      : o.direction === 'right'
        ? `Select \`${o.range}\` (click \`${o.from}\`, then Shift-click the last cell) and press {fillRight}.`
        : `Select \`${o.range}\` (click \`${o.from}\`, then Shift-click the bottom-right cell). Press {fillDown}, then {fillRight}.`;
  return { do: how, why: o.why, done: { kind: 'answer' } };
}

/** The last step of every walkthrough. */
export function checkStep(why = 'The coach changes the data a few times behind the scenes to make sure your work keeps giving the right answer, then puts it back.'): GuideStep {
  return { do: 'Click **Check** below.', why, done: { kind: 'check' } };
}

/**
 * Cells in one column for the given 1-based row numbers, with runs merged: ("F", [6, 7, 8, 12])
 * gives "F6:F8,F12". For selecting the rows a formula will use.
 */
export function cellList(column: string, rows: number[]): string {
  const sorted = [...new Set(rows)].sort((a, b) => a - b);
  const out: string[] = [];
  for (let i = 0; i < sorted.length; ) {
    let j = i;
    while (j + 1 < sorted.length && sorted[j + 1] === sorted[j] + 1) j++;
    out.push(i === j ? `${column}${sorted[i]}` : `${column}${sorted[i]}:${column}${sorted[j]}`);
    i = j + 1;
  }
  return out.join(',');
}

/** Sheet row numbers (1-based) of the data rows that pass `keep`, for data whose header is in `headerRow`. */
export function rowsWhere<T>(rows: readonly T[], keep: (row: T) => boolean, headerRow = 1): number[] {
  return rows.flatMap((r, i) => (keep(r) ? [headerRow + 1 + i] : []));
}

const usd = new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD' });

/** $1,234.56, the way Excel shows a currency cell. */
export function money(n: number): string {
  return usd.format(n);
}

/** Formats a date serial the way the sheet shows month headers: "Mar 2026". */
export function monthName(serial: number): string {
  const d = new Date(Date.UTC(1899, 11, 30) + serial * 86_400_000);
  return d.toLocaleString('en-US', { month: 'short', year: 'numeric', timeZone: 'UTC' });
}

/** Formats a date serial as the sheet's yyyy-mm-dd. */
export function isoDate(serial: number): string {
  return new Date(Date.UTC(1899, 11, 30) + serial * 86_400_000).toISOString().slice(0, 10);
}

/** A tip that recurs: picking Table columns as you type. */
export const TABLE_TYPING_TIP =
  'Typing tip: after you type a Table’s name and `[`, Excel lists its columns. Pick one with the arrow keys and press Tab, then type `]`. Spaces don’t matter, and capitals don’t either.';

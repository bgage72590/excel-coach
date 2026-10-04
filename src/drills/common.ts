import type { Block, ColumnSpec, Exercise, Grid, Inspection, SheetCheck } from '../engine/types';
import type { DrillItem } from './types';

/**
 * Drill items are one-action exercises graded by inspections alone: no answer cells, no variants.
 * Sheets are plain (no autofit, frozen header or header styling) so every bit of formatting the
 * checks look for is the learner's, and number formats are left at General unless the item is
 * about changing one.
 */

interface DrillSpec<D> extends Omit<Exercise<D>, 'module' | 'minutes' | 'layout' | 'expected' | 'inputs' | 'variants' | 'inspections'> {
  blocks(d: D): Block[];
  inspections(d: D): Inspection[];
  /** The fastest route per platform, shown once the item passes. */
  shortcut: DrillItem['shortcut'];
}

/** Builds a drill item: a plain sheet, graded by `inspections`, with a shortcut tip. */
export function drillItem<D>({ blocks, shortcut, ...spec }: DrillSpec<D>): DrillItem {
  const exercise: Exercise<D> = {
    ...spec,
    module: 'drills',
    minutes: 1,
    layout: (d) => ({ blocks: blocks(d), answer: { kind: 'objects' }, plain: true }),
    expected: () => [],
    inputs: () => [],
    variants: [],
  };
  return { exercise, shortcut };
}

/** Every drill item's id is its exercise id: what progress keys per-item bests by. */
export const drillItemId = (item: DrillItem): string => item.exercise.id;

/** A 'sheet' inspection. */
export const sheetCheck = (check: SheetCheck, label: string, advice?: string): Inspection => ({ kind: 'sheet', check, label, advice });

/** Header-only columns: drill data has no number formats unless the item is about one. */
export const headers = (...names: string[]): ColumnSpec[] => names.map((header) => ({ header }));

/** The last sheet row of a data block with a header in row 1. */
export const lastRow = (rows: readonly unknown[]) => rows.length + 1;

// ---------- number formats, as Office.js reports them ----------

/*
 * Each pattern skips the parts of a format code that aren't codes: [tags] such as [Red] or the
 * locale tag in [$-409], "quoted text", \escapes, and _x / *x padding. A "$" right after "[" is a
 * locale tag, not a currency symbol.
 */

/**
 * Currency with two decimals: "$#,##0.00" (Format Cells), "$#,##0.00_);($#,##0.00)" (Ctrl+Shift+$
 * on Windows), "$#,##0.00_);[Red]($#,##0.00)" (⌃⇧$ on a Mac), "[$$-409]#,##0.00", and Accounting,
 * _("$"* #,##0.00_);… from the ribbon's $ button. Whole-dollar and percent formats don't count.
 */
export const CURRENCY_FORMAT = /^(?!.*%)(?=.*0\.00)(?=.*(?<!\[)\$)/;

/** A percentage: "0%", "0.0%". */
export const PERCENT_FORMAT = /^(?:"[^"]*"|\\.|[^"\\%])*%/;

/** A date: a d or y code. Accepts d-mmm-yy, m/d/yyyy, yyyy-mm-dd and [$-409]mmmm d, yyyy;@. */
export const DATE_FORMAT = /^(?:\[[^\]]*\]|"[^"]*"|\\.|_.|\*.|[^"[\\_*dy])*[dy]/i;

/**
 * A time without a date: an h or s code (or elapsed [h], [mm], [ss]) and no d or y code. Accepts
 * h:mm AM/PM, [$-F400]h:mm:ss AM/PM (Windows' Time) and [h]:mm; rejects m/d/yyyy h:mm.
 */
export const TIME_FORMAT = /^(?!(?:\[[^\]]*\]|"[^"]*"|\\.|_.|\*.|[^"[\\_*dy])*[dy])(?:\[[^\]]*\]|"[^"]*"|\\.|_.|\*.|[^"[\\_*hs])*(?:[hs]|\[[hms]+\])/i;

/**
 * A number with a thousands separator and no currency symbol or percent: "#,##0.00" (Ctrl+Shift+!),
 * "#,##0", and Comma Style, _(* #,##0.00_);…
 */
export const THOUSANDS_FORMAT = /^(?!.*%)(?!.*(?<!\[)\$)(?=.*#,##[0#])/;

/** Negatives in red: [Red] in a section after the first, e.g. "0.00;[Red]0.00" or "$#,##0.00_);[Red]($#,##0.00)". */
export const RED_NEGATIVE_FORMAT = /;[^;]*\[Red\]/i;

export const GENERAL_FORMAT = /^General$/i;

// ---------- values the checks compare ----------

/** Remove Duplicates compares text ignoring case. */
function rowKey(row: Grid[number]): string {
  return JSON.stringify(row.map((v) => (typeof v === 'string' ? v.toLowerCase() : v)));
}

/**
 * What Data › Remove Duplicates leaves in the range, with every column checked: the first copy of
 * each row stays, later copies are deleted, the rows below move up, and the cells left over at the
 * bottom of the range are blank.
 */
export function removeDuplicateRows(rows: Grid): Grid {
  const seen = new Set<string>();
  const kept = rows.filter((row) => {
    const key = rowKey(row);
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
  const width = rows[0]?.length ?? 0;
  return [...kept, ...Array.from({ length: rows.length - kept.length }, () => Array.from({ length: width }, () => ''))];
}

// ---------- times ----------

/** A drill time as m:ss.t, truncated to tenths like a stopwatch: 4,280 ms is "0:04.2". */
export function formatDrillTime(ms: number): string {
  const tenths = Math.floor(Math.max(0, ms) / 100);
  const m = Math.floor(tenths / 600);
  const s = Math.floor((tenths % 600) / 10);
  return `${m}:${String(s).padStart(2, '0')}.${tenths % 10}`;
}

/** A gap between two times: "8.4 s" under a minute, "1:02.3" from a minute up. */
export function formatDrillGap(ms: number): string {
  const abs = Math.abs(ms);
  return abs < 60_000 ? `${(Math.floor(abs / 100) / 10).toFixed(1)} s` : formatDrillTime(abs);
}

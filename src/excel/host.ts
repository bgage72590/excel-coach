import type { RangeRead } from '../engine/fix';
import type { SheetFormulas } from '../engine/scan';
import type { Exercise, CheckReport, InputWrite } from '../engine/types';

/** The bridge between the panel and a workbook. ExcelHost talks to Excel; MockHost fakes it for design work. */
export interface CoachHost {
  /** True when connected to a real workbook. */
  readonly live: boolean;
  /** Adds a fresh practice sheet for this rep, replacing earlier coach sheets. */
  setup(ex: Exercise<any>, data: unknown, sheet: string): Promise<void>;
  /** Checks the learner's work. Rewrites inputs temporarily for each variant, then restores them. */
  check(ex: Exercise<any>, data: unknown, sheet: string, seed: number): Promise<CheckReport>;
  sheetExists(sheet: string): Promise<boolean>;
  select(sheet: string, address: string): Promise<void>;
  /** Reads formulas from every sheet except the coach's own, for the My Work scan. Local only. */
  readWorkbook(): Promise<{ name: string; sheets: SheetFormulas[]; truncated: boolean }>;
  /** Writes input regions on a practice sheet and leaves them (a mission step that adds new rows). */
  writeInputs(sheet: string, writes: InputWrite[]): Promise<void>;
  /**
   * Copies one of the learner's sheets for a guided fix and returns the copy's name (FIX_PREFIX +
   * the original name). An earlier copy of the same sheet is replaced. The original is never written.
   */
  copySheetForFix(sheet: string): Promise<string>;
  /** Reads values and formulas from any sheet. */
  readRange(sheet: string, address: string): Promise<RangeRead>;
}

export const SHEET_PREFIX = 'Coach-';
/** Copies made by "Fix it in My Work". Kept when practice sheets are replaced; skipped by the scan. */
export const FIX_PREFIX = 'Fix-';

export function sheetNameFor(exerciseId: string): string {
  return `${SHEET_PREFIX}${exerciseId}`.slice(0, 31);
}

/** The copy's name for a fix of `sheet`: Excel caps sheet names at 31 characters. */
export function fixSheetNameFor(sheet: string): string {
  return `${FIX_PREFIX}${sheet}`.slice(0, 31);
}

import type { Grid } from './types';

/**
 * "My Work" scanner. Pure functions over the formulas Excel reports for each sheet; nothing here
 * talks to Excel or leaves the machine. The host reads used ranges and passes them in.
 */

export interface SheetFormulas {
  sheet: string;
  /** Address of the used range that `formulas` starts at, e.g. "A1:K240" (no sheet prefix). */
  address: string;
  /** Excel's .formulas for the range: formula strings, or the value for non-formula cells. */
  formulas: Grid;
  /** Excel's .formulasR1C1 for the same range. */
  r1c1: Grid;
}

export interface FindingCell {
  sheet: string;
  address: string;
}

export interface Finding {
  /** Detector id, e.g. "vlookup". */
  id: string;
  severity: 'high' | 'medium' | 'low';
  /** Short headline, e.g. "VLOOKUP with a hard-coded column number". */
  title: string;
  /** One or two sentences: why it's a risk or a time sink. */
  why: string;
  /** One sentence: what to do instead. */
  fix: string;
  /** Exercise that teaches the fix, if there is one. */
  exerciseId?: string;
  /** Total occurrences. */
  count: number;
  /** Up to 8 example cells. */
  cells: FindingCell[];
  /** For hardcoded-constants: the numbers that were flagged, most widespread first. */
  numbers?: number[];
}

export interface ScanOptions {
  /** Ignore findings on sheets whose name starts with this (the coach's practice sheets). */
  skipPrefix?: string;
}

export { scanWorkbook } from './scanDetectors';

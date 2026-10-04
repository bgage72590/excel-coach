import type { CellsBlock, ColumnSpec, DataBlock, Exercise, Grid, InputWrite } from '../engine/types';

export const FMT = {
  currency: '$#,##0.00',
  rate: '$0.00',
  int: '#,##0',
  plain: '0',
  dec1: '0.0',
  date: 'yyyy-mm-dd',
  month: 'mmm yyyy',
  pct: '0.0%',
} as const;

/** Identity helper that keeps each exercise's data type inferred. */
export function defineExercise<D>(ex: Exercise<D>): Exercise<D> {
  return ex;
}

export function dataBlock(table: string, at: string, columns: ColumnSpec[], rows: Grid, asTable = true): DataBlock {
  return { kind: 'data', at, table, columns, rows, asTable };
}

export function cells(at: string, values: Grid, role: CellsBlock['role'], format?: string | (string | undefined)[]): CellsBlock {
  return Array.isArray(format) ? { kind: 'cells', at, values, role, formats: format } : { kind: 'cells', at, values, role, format };
}

export function tableWrite(table: string, columns: ColumnSpec[], rows: Grid): InputWrite {
  return { kind: 'table', table, columns: columns.map((c) => c.header), rows };
}

export function rangeWrite(address: string, values: Grid): InputWrite {
  return { kind: 'range', address, values };
}

/** Wraps a flat list as a single-column grid. */
export function column<T extends string | number>(values: readonly T[]): Grid {
  return values.map((v) => [v]);
}

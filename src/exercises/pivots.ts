import { ITEMS, WAREHOUSES } from '../engine/data';
import type { Rng } from '../engine/rng';
import type { ColumnSpec } from '../engine/types';
import { FMT, dataBlock, defineExercise } from './common';

interface Outbound {
  month: string;
  warehouse: string;
  sku: string;
  units: number;
}

const OUTBOUND_COLS: ColumnSpec[] = [{ header: 'Month' }, { header: 'Warehouse' }, { header: 'SKU' }, { header: 'Units', format: FMT.int }];

function outbound(rng: Rng): Outbound[] {
  return Array.from({ length: 100 }, () => ({
    month: rng.pick(['2026-07', '2026-08', '2026-09']),
    warehouse: rng.pick(WAREHOUSES),
    sku: `SKU-${100 + ITEMS.indexOf(rng.pick(ITEMS)) * 7}`,
    units: rng.int(5, 600),
  }));
}

export const pivotShare = defineExercise<{ rows: Outbound[] }>({
  id: 'pivot-share',
  module: 'pivots',
  title: 'Each warehouse’s volume by month',
  replaces: 'A SUMIFS grid plus a second grid of percentages',
  minutes: 5,
  task: () =>
    'Build a PivotTable from the `Outbound` Table with Warehouse as rows, Month as columns and the sum of Units as values. Then show the values as % of row total, so each warehouse’s months add up to 100%.',
  concept: {
    summary:
      'A PivotTable summarizes a table by dragging fields into Rows, Columns and Values. Show Values As changes what each number means, such as its share of the row, without any formulas.',
    syntax: 'Insert › PivotTable, then drag fields: Rows = Warehouse, Columns = Month, Values = Units',
    example: 'Right-click any value › Show Values As › % of Row Total.',
    tip: 'The PivotTable can go on a new sheet; the coach finds it anywhere in the workbook.',
  },
  hints: [
    'Click inside the Outbound Table, then Insert › PivotTable › OK.',
    'In the field list, drag Warehouse to Rows, Month to Columns and Units to Values.',
    'Right-click a number in the PivotTable › Show Values As › % of Row Total.',
  ],
  solution: () => 'Rows: Warehouse · Columns: Month · Values: Sum of Units · Show Values As: % of Row Total',
  make: (rng) => ({ rows: outbound(rng) }),
  layout: (d) => ({
    blocks: [dataBlock('Outbound', 'A1', OUTBOUND_COLS, d.rows.map((r) => [r.month, r.warehouse, r.sku, r.units]))],
    answer: { kind: 'pivot' },
  }),
  expected: () => [],
  inputs: () => [],
  variants: [],
  inspections: () => [
    {
      kind: 'pivot',
      rows: 'Warehouse',
      columns: 'Month',
      valuesField: 'Units',
      summarizeBy: 'Sum',
      showAs: 'PercentOfRowTotal',
      label: 'PivotTable layout',
    },
  ],
});

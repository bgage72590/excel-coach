import { CARRIERS, ITEMS, WAREHOUSES, serial, skuCode, sum } from '../engine/data';
import { round, type Rng } from '../engine/rng';
import type { ColumnSpec } from '../engine/types';
import { FMT, cells, dataBlock, defineExercise, tableWrite } from './common';

// ---------- Turn a range into a Table ----------

interface Shipment {
  id: string;
  date: number;
  carrier: string;
  warehouse: string;
  kg: number;
  freight: number;
}

const SHIP_COLS: ColumnSpec[] = [
  { header: 'Ship ID' },
  { header: 'Date', format: FMT.date },
  { header: 'Carrier' },
  { header: 'Warehouse' },
  { header: 'Weight kg', format: FMT.int },
  { header: 'Freight', format: FMT.currency },
];

function shipment(rng: Rng, n: number): Shipment {
  const kg = rng.int(40, 2400);
  return {
    id: `SH-${10200 + n}`,
    date: rng.int(serial(2026, 7, 1), serial(2026, 9, 30)),
    carrier: rng.pick(CARRIERS),
    warehouse: rng.pick(WAREHOUSES),
    kg,
    freight: round(45 + kg * rng.float(0.18, 0.42, 3), 2),
  };
}

const shipGrid = (rows: Shipment[]) => rows.map((s) => [s.id, s.date, s.carrier, s.warehouse, s.kg, s.freight]);

export const tablesConvert = defineExercise<{ rows: Shipment[] }>({
  id: 'tables-convert',
  module: 'tables',
  title: 'Turn a range into a Table',
  replaces: 'Totals that silently miss the rows you add next week',
  minutes: 4,
  task: () =>
    'Turn the shipment data in `A1:F41` into a Table named `Shipments`. Then in `I2`, total the Freight column by its Table name instead of a cell range.',
  concept: {
    summary:
      'A Table is a range Excel knows the shape of. It grows when you add rows, and formulas can refer to columns by name, so a total keeps counting new data without you editing it.',
    syntax: '=SUM(Shipments[Freight])',
    example: 'Shipments[Freight] means “every value in the Freight column of the Shipments Table”, however many rows it has.',
    tip: 'Rename a Table on the Table tab (Table Design on Windows) in the Table Name box.',
  },
  hints: [
    'Click any cell in the data and press {tableKey}. Leave “My table has headers” checked.',
    'With a cell in the Table selected, open the Table tab (Table Design on Windows) and set Table Name to Shipments.',
    'In I2 type =SUM( and then click the top edge of the Freight header, or type Shipments[Freight] directly.',
  ],
  solution: () => '=SUM(Shipments[Freight])',
  make: (rng) => ({ rows: Array.from({ length: 40 }, (_, i) => shipment(rng, i)).sort((a, b) => a.date - b.date) }),
  layout: (d) => ({
    blocks: [dataBlock('Shipments', 'A1', SHIP_COLS, shipGrid(d.rows), false), cells('H2', [['Total freight']], 'label')],
    answer: { kind: 'cells', range: 'I2', format: FMT.currency, consistency: 'none' },
  }),
  expected: (d) => [[sum(d.rows.map((r) => r.freight))]],
  inputs: (d) => [tableWrite('Shipments', SHIP_COLS, shipGrid(d.rows))],
  variants: [
    {
      label: 'five more shipments are added to the Table',
      explain: 'A fixed range like F2:F41 stops at the old last row. A Table column grows with the data.',
      apply: (d, rng) => ({ rows: [...d.rows, ...Array.from({ length: 5 }, (_, i) => shipment(rng, 40 + i))] }),
    },
    {
      label: 'the freight amounts change',
      apply: (d, rng) => ({ rows: d.rows.map((r) => ({ ...r, freight: round(r.freight * rng.float(0.5, 1.8, 3), 2) })) }),
    },
  ],
  rules: {
    require: [
      {
        pattern: /Shipments\[\s*Freight\s*\]/i,
        label: 'Totals the Table column by name',
        advice: 'Refer to the column as Shipments[Freight] instead of a cell range like F2:F41.',
      },
    ],
    allowNumbers: [],
  },
  inspections: (d) => [
    { kind: 'tableExists', table: 'Shipments', at: 'A1', rows: d.rows.length + 1, label: 'A Table named Shipments holds all the shipment data' },
  ],
});

// ---------- Add a calculated column ----------

interface OrderLine {
  order: string;
  sku: string;
  item: string;
  qty: number;
  cost: number;
}

const ORDER_COLS: ColumnSpec[] = [
  { header: 'Order' },
  { header: 'SKU' },
  { header: 'Item' },
  { header: 'Qty', format: FMT.int },
  { header: 'Unit cost', format: FMT.currency },
];

function orderLine(rng: Rng, n: number): OrderLine {
  const idx = rng.int(0, ITEMS.length - 1);
  const item = ITEMS[idx];
  return { order: `PO-${4100 + Math.floor(n / 3)}`, sku: skuCode(100 + idx * 7), item: item.item, qty: rng.int(1, 120), cost: item.cost };
}

const orderGrid = (rows: OrderLine[]) => rows.map((o) => [o.order, o.sku, o.item, o.qty, o.cost]);

export const tablesCalcColumn = defineExercise<{ rows: OrderLine[] }>({
  id: 'tables-calc-column',
  module: 'tables',
  title: 'Add a calculated column',
  replaces: 'Copying a formula down by hand and missing the new rows',
  minutes: 3,
  task: () =>
    'The `Orders` Table lists purchase order lines. Add a column named `Line total` in `F` that multiplies Qty by Unit cost for every row.',
  concept: {
    summary:
      'Type a formula once in a new Table column and Excel fills the whole column, including rows added later. Row references such as [@Qty] read “the Qty value on this row”.',
    syntax: '=[@Qty]*[@[Unit cost]]',
    example: 'Column names with spaces get an extra pair of brackets: [@[Unit cost]].',
  },
  hints: [
    'Type Line total in F1 and press {enter}. The Table grows to include the new column.',
    'In F2 type = and click D2, type *, then click E2. Excel writes [@Qty] and [@[Unit cost]] for you.',
    'Press {enter} once. The Table fills the formula down by itself.',
  ],
  solution: () => '=[@Qty]*[@[Unit cost]]',
  make: (rng) => ({ rows: Array.from({ length: 30 }, (_, i) => orderLine(rng, i)) }),
  layout: (d) => ({
    blocks: [dataBlock('Orders', 'A1', ORDER_COLS, orderGrid(d.rows))],
    answer: { kind: 'tableColumn', table: 'Orders', column: 'Line total', format: FMT.currency },
  }),
  expected: (d) => d.rows.map((r) => [r.qty * r.cost]),
  inputs: (d) => [tableWrite('Orders', ORDER_COLS, orderGrid(d.rows))],
  variants: [
    {
      label: 'quantities change',
      apply: (d, rng) => ({ rows: d.rows.map((r) => ({ ...r, qty: rng.int(1, 150) })) }),
    },
    {
      label: 'three new order lines are added',
      explain: 'A calculated column fills new rows automatically. Formulas typed cell by cell don’t.',
      apply: (d, rng) => ({ rows: [...d.rows, ...Array.from({ length: 3 }, (_, i) => orderLine(rng, 30 + i))] }),
    },
  ],
  rules: {
    require: [
      {
        pattern: /\[@/,
        label: 'Uses row references like [@Qty]',
        advice: 'Click the cells on the same row while you type, and Excel writes [@Qty] and [@[Unit cost]].',
      },
    ],
    allowNumbers: [],
  },
  inspections: () => [{ kind: 'tableColumn', table: 'Orders', column: 'Line total', label: 'The Orders Table has a Line total column' }],
});

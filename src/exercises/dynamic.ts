import { CARRIERS, EXTRA_VENDORS, VENDORS, excelTextCompare, serial, sum } from '../engine/data';
import { round, type Rng } from '../engine/rng';
import type { ColumnSpec, ExpectedGrid } from '../engine/types';
import { FMT, cells, dataBlock, defineExercise, rangeWrite, tableWrite } from './common';
import { SPEND_COLS, spendGrid, spendLine } from './sumifs';

// ---------- Late shipments for one carrier ----------

interface Delivery {
  id: string;
  carrier: string;
  promised: number;
  delivered: number;
  late: number;
}

const DELIVERY_COLS: ColumnSpec[] = [
  { header: 'Ship ID' },
  { header: 'Carrier' },
  { header: 'Promised', format: FMT.date },
  { header: 'Delivered', format: FMT.date },
  { header: 'Days late', format: FMT.plain },
];

function deliver(rng: Rng, d: Omit<Delivery, 'delivered' | 'late'>, forceLate = false): Delivery {
  const late = forceLate ? rng.int(1, 6) : rng.chance(0.3) ? rng.int(1, 6) : -rng.int(0, 3);
  return { ...d, delivered: d.promised + late, late };
}

function deliveries(rng: Rng): Delivery[] {
  const rows = Array.from({ length: 30 }, (_, i) =>
    deliver(rng, { id: `SH-${31000 + i}`, carrier: rng.pick(CARRIERS), promised: rng.int(serial(2026, 8, 1), serial(2026, 9, 25)) }),
  );
  return ensureLatePerCarrier(rng, rows);
}

/** Every carrier gets at least two late deliveries, so a FILTER never comes back empty. */
function ensureLatePerCarrier(rng: Rng, rows: Delivery[]): Delivery[] {
  const out = rows.slice();
  for (const carrier of CARRIERS) {
    let late = out.filter((r) => r.carrier === carrier && r.late > 0).length;
    while (late < 2) {
      const idx = rng.int(0, out.length - 1);
      if (out[idx].late > 0) continue;
      out[idx] = deliver(rng, { id: out[idx].id, carrier, promised: out[idx].promised }, true);
      late++;
    }
  }
  return out;
}

const deliveryGrid = (rows: Delivery[]) => rows.map((r) => [r.id, r.carrier, r.promised, r.delivered, r.late]);

interface LateData {
  rows: Delivery[];
  carrier: string;
}

export const filterLate = defineExercise<LateData>({
  id: 'filter-late',
  module: 'dynamic',
  title: 'Late shipments for one carrier',
  replaces: 'Filtering, copying and pasting the same report every week',
  minutes: 5,
  task: () =>
    'With one formula in `G4`, list every delivery for the carrier named in `H1` that arrived late (Days late above 0). Return all five columns.',
  concept: {
    summary:
      'FILTER returns every row that meets a condition, and the result spills into as many cells as it needs. For two conditions, multiply them: both must be true.',
    syntax: '=FILTER(array, (condition1)*(condition2), [if_empty])',
    example: '=FILTER(Deliveries, (Deliveries[Carrier]=H1)*(Deliveries[Days late]>0))',
  },
  hints: [
    'Return the whole Table: the first argument of FILTER is Deliveries.',
    'Each condition is a comparison in parentheses, such as (Deliveries[Carrier]=H1).',
    'Multiply the two conditions: (Deliveries[Carrier]=H1)*(Deliveries[Days late]>0).',
  ],
  solution: () => '=FILTER(Deliveries,(Deliveries[Carrier]=H1)*(Deliveries[Days late]>0))',
  make: (rng) => ({ rows: deliveries(rng), carrier: rng.pick(CARRIERS) }),
  layout: (d) => ({
    blocks: [
      dataBlock('Deliveries', 'A1', DELIVERY_COLS, deliveryGrid(d.rows)),
      cells('G1', [['Carrier']], 'label'),
      cells('H1', [[d.carrier]], 'input'),
      cells('G3', [DELIVERY_COLS.map((c) => c.header)], 'header'),
    ],
    answer: { kind: 'spill', anchor: 'G4', formats: DELIVERY_COLS.map((c) => c.format) },
  }),
  expected: (d) => deliveryGrid(d.rows.filter((r) => r.carrier === d.carrier && r.late > 0)),
  inputs: (d) => [tableWrite('Deliveries', DELIVERY_COLS, deliveryGrid(d.rows)), rangeWrite('H1', [[d.carrier]])],
  variants: [
    {
      label: 'a different carrier is chosen in H1',
      explain: 'Compare against H1 instead of typing the carrier’s name.',
      apply: (d, rng) => ({ ...d, carrier: rng.pick(CARRIERS.filter((c) => c !== d.carrier)) }),
    },
    {
      label: 'delivery dates change',
      apply: (d, rng) => ({ ...d, rows: ensureLatePerCarrier(rng, d.rows.map((r) => deliver(rng, r))) }),
    },
  ],
  rules: {
    require: [{ pattern: /FILTER\(/i, label: 'Uses FILTER', advice: 'FILTER(array, include) returns every matching row in one go.' }],
    forbidText: { values: [...CARRIERS], advice: 'Compare against H1 instead.' },
    allowNumbers: [0, 1],
  },
});

// ---------- A sorted list of vendors ----------

interface SpendData {
  rows: ReturnType<typeof spendLine>[];
}

const sortedUnique = (values: string[]) => [...new Set(values)].sort(excelTextCompare);

function spendData(rng: Rng, count: number): SpendData {
  const vendors = rng.sample(VENDORS, 5);
  return { rows: Array.from({ length: count }, () => spendLine(rng, vendors)) };
}

function withNewVendor(d: SpendData, rng: Rng): SpendData {
  const vendor = rng.pick(EXTRA_VENDORS);
  return { rows: [...d.rows, ...Array.from({ length: 3 }, () => spendLine(rng, [vendor]))] };
}

export const uniqueVendors = defineExercise<SpendData>({
  id: 'unique-vendors',
  module: 'dynamic',
  title: 'A sorted list of vendors',
  replaces: 'Copy, paste, Remove Duplicates, sort, and doing it again next month',
  minutes: 3,
  task: () => 'With one formula in `F2`, list each vendor in the `Spend` Table once, sorted A to Z.',
  concept: {
    summary: 'UNIQUE returns each distinct value once. Wrap it in SORT and you get a clean list that updates itself when the data changes.',
    syntax: '=SORT(UNIQUE(column))',
    example: '=SORT(UNIQUE(Spend[Vendor]))',
  },
  hints: ['UNIQUE(Spend[Vendor]) gives each vendor once.', 'Wrap it in SORT to put the list in A to Z order.'],
  solution: () => '=SORT(UNIQUE(Spend[Vendor]))',
  make: (rng) => spendData(rng, 60),
  layout: (d) => ({
    blocks: [dataBlock('Spend', 'A1', SPEND_COLS, spendGrid(d.rows)), cells('F1', [['Vendors']], 'header')],
    answer: { kind: 'spill', anchor: 'F2' },
  }),
  expected: (d) => sortedUnique(d.rows.map((r) => r.vendor)).map((v) => [v]),
  inputs: (d) => [tableWrite('Spend', SPEND_COLS, spendGrid(d.rows))],
  variants: [
    {
      label: 'a new vendor appears in the data',
      explain: 'Point UNIQUE at the whole Table column so new names show up.',
      apply: withNewVendor,
    },
    { label: 'the data is re-sorted', apply: (d, rng) => ({ rows: rng.shuffle(d.rows) }) },
  ],
  rules: {
    require: [
      { pattern: /UNIQUE\(/i, label: 'Uses UNIQUE', advice: 'UNIQUE(Spend[Vendor]) returns each vendor once.' },
      { pattern: /SORT(BY)?\(/i, label: 'Sorts the list A to Z', advice: 'Wrap the list in SORT( … ).' },
    ],
  },
});

// ---------- A summary that grows by itself ----------

export const spillReference = defineExercise<SpendData>({
  id: 'spill-reference',
  module: 'dynamic',
  title: 'A summary that grows by itself',
  replaces: 'Summary tables you have to extend whenever a new vendor shows up',
  minutes: 4,
  task: () =>
    '`F2` already lists every vendor with one spilling formula. With one formula in `G2`, total each vendor’s spend so the totals grow and shrink with that list.',
  concept: {
    summary:
      'Add # after a cell that spills to mean “the whole spilled result”. F2# grows when F2’s list grows, so formulas built on it never need extending.',
    syntax: '=SUMIFS(Spend[Amount], Spend[Vendor], F2#)',
    example: 'Giving SUMIFS a whole list of criteria returns a whole list of totals, one per vendor.',
  },
  hints: ['A normal SUMIFS adds Spend[Amount] where Spend[Vendor] matches a vendor.', 'Use F2# as the criteria to mean the entire vendor list.'],
  solution: () => '=SUMIFS(Spend[Amount],Spend[Vendor],F2#)',
  make: (rng) => spendData(rng, 60),
  layout: (d) => ({
    blocks: [
      dataBlock('Spend', 'A1', SPEND_COLS, spendGrid(d.rows)),
      cells('F1', [['Vendor', 'Spend']], 'header'),
      cells('F2', [['=SORT(UNIQUE(Spend[Vendor]))']], 'formula'),
    ],
    answer: { kind: 'spill', anchor: 'G2', format: FMT.currency },
  }),
  expected: (d) => sortedUnique(d.rows.map((r) => r.vendor)).map((v) => [sum(d.rows.filter((r) => r.vendor === v).map((r) => r.amount))]),
  inputs: (d) => [tableWrite('Spend', SPEND_COLS, spendGrid(d.rows))],
  variants: [
    { label: 'a new vendor appears', explain: 'Use F2# so the totals follow the list as it grows.', apply: withNewVendor },
    { label: 'the amounts change', apply: (d, rng) => ({ rows: d.rows.map((r) => ({ ...r, amount: round(r.amount * rng.float(0.5, 1.8, 3), 2) })) }) },
  ],
  rules: {
    require: [
      {
        pattern: /F2#|ANCHORARRAY\(\s*F2\s*\)/i,
        label: 'Refers to the whole vendor list with F2#',
        advice: 'Use F2# to mean “everything F2 spills”, so the totals grow with the list.',
      },
    ],
  },
});

// ---------- Spend by category in one formula ----------

const TOTAL = { match: /total/i, describe: 'a Total label' };

export const groupbyCategory = defineExercise<SpendData>({
  id: 'groupby-category',
  module: 'dynamic',
  title: 'Spend by category in one formula',
  replaces: 'Building a PivotTable for a quick two-column summary',
  minutes: 3,
  m365: true,
  task: () => 'With one formula in `F2`, summarize total spend by category, sorted A to Z with a grand total row at the bottom.',
  concept: {
    summary:
      'GROUPBY builds a summary like a small PivotTable, but as a formula that recalculates instantly. By default it sorts the groups and adds a Total row.',
    syntax: '=GROUPBY(row_fields, values, function)',
    example: '=GROUPBY(Spend[Category], Spend[Amount], SUM)',
    tip: 'Pass the function by name, without parentheses: SUM, AVERAGE, COUNT.',
  },
  hints: ['The groups come from Spend[Category]; the numbers come from Spend[Amount].', 'The third argument is the function to apply: SUM.'],
  solution: () => '=GROUPBY(Spend[Category],Spend[Amount],SUM)',
  make: (rng) => spendData(rng, 70),
  layout: (d) => ({
    blocks: [dataBlock('Spend', 'A1', SPEND_COLS, spendGrid(d.rows)), cells('F1', [['Category', 'Spend']], 'header')],
    answer: { kind: 'spill', anchor: 'F2', formats: [undefined, FMT.currency] },
  }),
  expected: (d): ExpectedGrid => {
    const groups = sortedUnique(d.rows.map((r) => r.category));
    const rows: ExpectedGrid = groups.map((g) => [g, sum(d.rows.filter((r) => r.category === g).map((r) => r.amount))]);
    rows.push([TOTAL, sum(d.rows.map((r) => r.amount))]);
    return rows;
  },
  inputs: (d) => [tableWrite('Spend', SPEND_COLS, spendGrid(d.rows))],
  variants: [
    { label: 'the amounts change', apply: (d, rng) => ({ rows: d.rows.map((r) => ({ ...r, amount: round(r.amount * rng.float(0.5, 1.8, 3), 2) })) }) },
    {
      label: 'a new category appears',
      explain: 'Group on the whole Table column so new categories get their own row.',
      apply: (d, rng) => ({
        rows: [
          ...d.rows,
          ...Array.from({ length: 3 }, () => ({ ...spendLine(rng, [d.rows[0].vendor]), category: 'Office' })),
        ],
      }),
    },
  ],
  rules: {
    require: [{ pattern: /GROUPBY\(/i, label: 'Uses GROUPBY', advice: 'GROUPBY(Spend[Category], Spend[Amount], SUM) builds the whole summary.' }],
  },
});

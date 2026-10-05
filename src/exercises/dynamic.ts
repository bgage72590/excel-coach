import { CARRIERS, EXTRA_VENDORS, VENDORS, excelTextCompare, serial, sum } from '../engine/data';
import { round, type Rng } from '../engine/rng';
import type { ColumnSpec, ExpectedGrid } from '../engine/types';
import { FMT, cells, dataBlock, defineExercise, rangeWrite, tableWrite } from './common';
import { TABLE_TYPING_TIP, cellList, checkStep, money, part, raw, rowsWhere, typeStep } from './guides';
import { SPEND_COLS, spendGrid, spendLine } from './sumifs';

/** Whole rows from column `from` to column `to` for the given sheet rows, runs merged: "A2:E2,A5:E7". */
function rowSpans(from: string, to: string, rows: number[]): string {
  return cellList(from, rows)
    .split(',')
    .map((area) => {
      const [a, b = a] = area.split(':');
      return `${a}:${to}${b.slice(from.length)}`;
    })
    .join(',');
}

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
  guide: (d) => {
    const keep = (r: Delivery) => r.carrier === d.carrier && r.late > 0;
    const late = d.rows.filter(keep);
    const lastRow = 3 + Math.max(late.length, 1);
    const carrierRows = d.rows.filter((r) => r.carrier === d.carrier).length;
    return [
      {
        do: `Meet the data. The blue block in columns \`A\` to \`E\` is a Table named **Deliveries**. \`H1\` names the carrier to report on: ${d.carrier}.`,
        why: 'In a formula, `Deliveries` on its own means the whole Table: every row and all five columns, without the header row. `Deliveries[Carrier]` means one column. Tap the buttons to see each part.',
        show: [
          { label: 'The whole Table', at: 'Deliveries[#Data]', note: 'That’s `Deliveries`: all five columns. FILTER copies whole rows from here.' },
          { label: 'Carrier column', at: 'Deliveries[Carrier]', note: `That’s \`Deliveries[Carrier]\` (column \`B\`). ${carrierRows} of the ${d.rows.length} rows are ${d.carrier}.` },
          { label: 'Days late column', at: 'Deliveries[Days late]', note: 'That’s `Deliveries[Days late]` (column `E`). Above 0 means the delivery arrived late; 0 or below means on time or early.' },
        ],
      },
      {
        do: `See which rows the list should hold: ${d.carrier} deliveries with Days late above 0.`,
        why: 'This is what you’d get by filtering Carrier and Days late, then copying the rows out. FILTER does both filters in one formula and rebuilds the list whenever the data changes.',
        show: late.length
          ? [
              {
                label: `Select ${d.carrier}’s late rows`,
                at: rowSpans('A', 'E', rowsWhere(d.rows, keep)),
                note: `These ${late.length} rows are ${d.carrier}’s late deliveries. FILTER copies them, all five columns and in this order, into \`G4\` and the cells below. \`G4\` should show ${late[0].id}.`,
              },
            ]
          : undefined,
      },
      typeStep({
        cell: 'G4',
        whole: true,
        formula: [
          part('=FILTER(', 'Returns every row that passes a test.'),
          part('Deliveries', 'What to return: the whole Deliveries Table, all five columns.', 'Deliveries[#Data]'),
          raw(', '),
          raw('('),
          part('Deliveries[Carrier]', 'Test 1 looks in the Carrier column…', 'Deliveries[Carrier]'),
          part('=H1', `…for the carrier named in \`H1\` (${d.carrier}). Every row gets TRUE or FALSE.`, 'H1'),
          raw(')'),
          part('*', 'Multiplies the two tests. TRUE counts as 1 and FALSE as 0, so a row scores 1 only when both tests are TRUE. Read `*` as “and”.'),
          raw('('),
          part('Deliveries[Days late]', 'Test 2 looks in the Days late column…', 'Deliveries[Days late]'),
          part('>0', '…for numbers above 0: the late deliveries.'),
          raw(')'),
          raw(')'),
        ],
        why: `${TABLE_TYPING_TIP} One formula fills the whole list: the rows spill, which means Excel writes them into \`G4\` and as many cells below and to the right as it needs (\`G4:K${lastRow}\` today). Keep those cells empty: a #SPILL! error in \`G4\` means something is in the way.`,
      }),
      {
        do: 'Click `I4`, a cell inside the list that you didn’t type in.',
        why: 'The formula bar shows `G4`’s formula in grey: `I4` has no formula of its own. Every cell of the list belongs to `G4`’s spill, so to change the list, edit `G4`.',
        done: { kind: 'select', range: `H4:K${lastRow}` },
      },
      checkStep('The coach picks a different carrier in `H1` and changes the delivery dates behind the scenes, then puts everything back. Your formula reads `H1` and the whole Table, so the list rebuilds itself each time.'),
    ];
  },
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
  guide: (d) => {
    const vendors = sortedUnique(d.rows.map((r) => r.vendor));
    // UNIQUE keeps each name where it first appears, top to bottom.
    const firstRow = new Map<string, number>();
    d.rows.forEach((r, i) => {
      if (!firstRow.has(r.vendor)) firstRow.set(r.vendor, i + 2);
    });
    return [
      {
        do: 'Meet the data. The blue block in columns `A` to `D` is a Table named **Spend**.',
        why: `Its Vendor column has ${d.rows.length} rows but only ${vendors.length} different vendors, because each one repeats. In a formula, \`Spend[Vendor]\` means that whole column, and it grows when rows are added.`,
        show: [{ label: 'Vendor column', at: 'Spend[Vendor]', note: 'That’s `Spend[Vendor]` (column `B`): the names to list once each.' }],
      },
      {
        do: `See what the list should be: ${vendors.length} names, A to Z, starting in \`F2\`.`,
        why: `UNIQUE keeps the first row where each name appears and drops the repeats. SORT then puts those names in A to Z order: ${vendors.join(', ')}.`,
        show: [
          {
            label: 'Select each vendor’s first row',
            at: cellList('B', [...firstRow.values()]),
            note: `UNIQUE keeps these ${firstRow.size} names, top to bottom: ${[...firstRow.keys()].join(', ')}. SORT reorders them A to Z.`,
          },
        ],
      },
      typeStep({
        cell: 'F2',
        whole: true,
        formula: [
          part('=SORT(', 'Puts whatever is inside its brackets in A to Z order.'),
          part('UNIQUE(', 'Keeps each value once, dropping the repeats.'),
          part('Spend[Vendor]', 'Which values: the whole Vendor column.', 'Spend[Vendor]'),
          part('))', 'Closes UNIQUE, then SORT. Every opening bracket needs a closing one.'),
        ],
        why: `${TABLE_TYPING_TIP} One formula makes the whole list: the names spill, which means Excel writes them into \`F2\` and the cells below it (\`F2:F${vendors.length + 1}\` today). Keep those cells empty: a #SPILL! error in \`F2\` means something is in the way.`,
      }),
      checkStep('The coach adds rows for a new vendor and shuffles the data behind the scenes, then puts everything back. Your list picks up the new name and stays A to Z.'),
    ];
  },
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
  guide: (d) => {
    const vendors = sortedUnique(d.rows.map((r) => r.vendor));
    const last = vendors.length + 1;
    const list = `F2:F${last}`;
    const first = vendors[0];
    const keep = (r: SpendData['rows'][number]) => r.vendor === first;
    const total = sum(d.rows.filter(keep).map((r) => r.amount));
    return [
      {
        do: 'Meet the layout. The Table on the left is named **Spend**, and `F2` holds one formula that lists every vendor.',
        why: `That formula, \`=SORT(UNIQUE(Spend[Vendor]))\`, spills: Excel writes its results into \`F2\` and as many cells below as the list needs. Today that’s ${vendors.length} names in \`${list}\`.`,
        show: [
          { label: 'The vendor list', at: list, note: `${vendors.length} vendors. Only \`F2\` holds a formula; the cells below it show its spilled results.` },
          { label: 'Amount column', at: 'Spend[Amount]', note: 'That’s `Spend[Amount]` (column `D`): the numbers you’ll add up.' },
          { label: 'Vendor column', at: 'Spend[Vendor]', note: 'That’s `Spend[Vendor]` (column `B`): where SUMIFS looks for each vendor.' },
        ],
      },
      {
        do: 'Click `F3` and look at the formula bar.',
        why: 'It shows `F2`’s formula in grey: `F3` has no formula of its own. It belongs to `F2`’s spill, the block of cells that one formula fills. To mean that whole block in another formula, write `F2#`. The `#` says “everything `F2` spills”, so `F2#` grows and shrinks with the list.',
        done: { kind: 'select', range: `F3:F${Math.max(3, last)}` },
      },
      {
        do: `See what \`G2\` should show: the total for ${first}, the first name in the list.`,
        show: d.rows.some(keep)
          ? [
              {
                label: `Select ${first}’s amounts`,
                at: cellList('D', rowsWhere(d.rows, keep)),
                note: `The status bar’s **Sum** is ${money(total)}. That’s the number beside ${first} in \`G2\`. The other vendors’ totals fill in below it.`,
              },
            ]
          : undefined,
      },
      typeStep({
        cell: 'G2',
        whole: true,
        formula: [
          part('=SUMIFS(', 'Adds up the rows that match a condition.'),
          part('Spend[Amount]', 'What to add up: the Amount column.', 'Spend[Amount]'),
          raw(', '),
          part('Spend[Vendor]', 'Where to look: the Vendor column.', 'Spend[Vendor]'),
          raw(', '),
          part('F2#', `What to look for: every vendor in \`F2\`’s list at once. Give SUMIFS a list of ${vendors.length} vendors and it gives back ${vendors.length} totals, one per vendor.`, list),
          raw(')'),
        ],
        why: `Type the \`#\` straight after \`F2\`. The totals spill down \`G2:G${last}\`, each one beside its vendor, so keep those cells empty. A #SPILL! error means something is in the way.`,
      }),
      checkStep('The coach adds a new vendor behind the scenes, then puts everything back. `F2`’s list grows by one name, `F2#` grows with it, and a new total appears beside it with no editing.'),
    ];
  },
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
  guide: (d) => {
    const groups = sortedUnique(d.rows.map((r) => r.category));
    const first = groups[0];
    const keep = (r: SpendData['rows'][number]) => r.category === first;
    const total = sum(d.rows.filter(keep).map((r) => r.amount));
    const grand = sum(d.rows.map((r) => r.amount));
    return [
      {
        do: 'Meet the data. The blue block in columns `A` to `D` is a Table named **Spend**. Each row is one purchase, tagged with a Category.',
        why: 'You want one row per category with its total spend: the summary a PivotTable would give, written as a formula instead. In a formula, `Spend[Category]` means the whole Category column.',
        show: [
          { label: 'Category column', at: 'Spend[Category]', note: `That’s \`Spend[Category]\` (column \`C\`): ${groups.length} different categories, repeated down the rows.` },
          { label: 'Amount column', at: 'Spend[Amount]', note: 'That’s `Spend[Amount]` (column `D`): the numbers to total.' },
        ],
      },
      {
        do: `See what the summary should hold: ${groups.length} categories A to Z, each with its total, then a Total row.`,
        why: `${first} comes first, so its total goes in \`G2\`. The Total row at the bottom adds up every purchase: ${money(grand)}.`,
        show: d.rows.some(keep)
          ? [
              {
                label: `Select the ${first} amounts`,
                at: cellList('D', rowsWhere(d.rows, keep)),
                note: `The status bar’s **Sum** is ${money(total)}. That’s the number beside ${first} in the first row of the summary.`,
              },
            ]
          : undefined,
      },
      typeStep({
        cell: 'F2',
        whole: true,
        formula: [
          part('=GROUPBY(', 'Groups rows that share a value and sums up each group, like a small PivotTable.'),
          part('Spend[Category]', 'What to group by: each category becomes one row of the summary.', 'Spend[Category]'),
          raw(', '),
          part('Spend[Amount]', 'What to add up for each group: the Amount column.', 'Spend[Amount]'),
          raw(', '),
          part('SUM', 'How to combine each group’s amounts: add them. Type the name alone, with no bracket after it. If you pick SUM from Excel’s list and it adds `(`, delete the `(`.'),
          raw(')'),
        ],
        why: `The summary spills, which means Excel writes it into \`F2\` and the cells below and to the right: two columns wide, from \`F2\` down to \`G${groups.length + 2}\`. Keep that area empty; a #SPILL! error means something is in the way. GROUPBY sorts the categories A to Z and adds the Total row by itself.`,
      }),
      checkStep('The coach changes the amounts and adds purchases in a new category behind the scenes, then puts everything back. The new category gets its own row, with no editing.'),
    ];
  },
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

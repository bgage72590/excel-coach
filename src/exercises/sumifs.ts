import { CATEGORIES, ITEMS, PRODUCTS, REGIONS, REPS, VENDORS, WAREHOUSES, eomonth, serial, skuCode, sum } from '../engine/data';
import { round, type Rng } from '../engine/rng';
import type { ColumnSpec } from '../engine/types';
import { FMT, cells, column, dataBlock, defineExercise, rangeWrite, tableWrite } from './common';
import { TABLE_TYPING_TIP, cellList, checkStep, fillStep, money, monthName, part, raw, rowsWhere, typeStep } from './guides';

// ---------- Inventory value by warehouse ----------

interface StockLine {
  sku: string;
  item: string;
  warehouse: string;
  onHand: number;
  cost: number;
  value: number;
}

const INV_COLS: ColumnSpec[] = [
  { header: 'SKU' },
  { header: 'Item' },
  { header: 'Warehouse' },
  { header: 'On hand', format: FMT.int },
  { header: 'Unit cost', format: FMT.currency },
  { header: 'Value', format: FMT.currency },
];

function stockLine(rng: Rng): StockLine {
  const idx = rng.int(0, ITEMS.length - 1);
  const onHand = rng.int(0, 900);
  const cost = ITEMS[idx].cost;
  return { sku: skuCode(100 + idx * 7), item: ITEMS[idx].item, warehouse: rng.pick(WAREHOUSES), onHand, cost, value: round(onHand * cost, 2) };
}

const invGrid = (rows: StockLine[]) => rows.map((s) => [s.sku, s.item, s.warehouse, s.onHand, s.cost, s.value]);

interface InvData {
  rows: StockLine[];
  order: string[];
}

export const sumifsWarehouse = defineExercise<InvData>({
  id: 'sumifs-warehouse',
  module: 'sumifs',
  title: 'Inventory value by warehouse',
  replaces: 'Filtering by warehouse and reading the total off the status bar',
  minutes: 4,
  task: () =>
    'In `I2:I5`, total the inventory Value for each warehouse listed in `H2:H5`. Write one formula in `I2` and fill it down.',
  concept: {
    summary:
      'SUMIFS adds up one column, keeping only the rows that meet every condition you give it. Point each condition at a cell so the same formula works for every row of your summary.',
    syntax: '=SUMIFS(sum_column, criteria_column1, criteria1, [criteria_column2, criteria2], …)',
    example: '=SUMIFS(Inventory[Value], Inventory[Warehouse], H2)',
  },
  hints: [
    'SUMIFS takes the column to add first, then pairs of (column to test, value to match).',
    'The column to add is Inventory[Value]. The column to test is Inventory[Warehouse].',
    'Match against H2, not a typed warehouse name, then fill the formula down to I5.',
  ],
  solution: () => '=SUMIFS(Inventory[Value],Inventory[Warehouse],H2)',
  guide: (d) => {
    const first = d.order[0];
    const rows = rowsWhere(d.rows, (r) => r.warehouse === first);
    const total = sum(d.rows.filter((r) => r.warehouse === first).map((r) => r.value));
    return [
      {
        do: 'Meet the data. The blue block in columns `A` to `F` is a Table named **Inventory**.',
        why: 'A Table has a name, and so does each of its columns. In a formula, `Inventory[Value]` means the whole Value column, and it grows when rows are added. Tap the buttons to see each column.',
        show: [
          { label: 'Value column', at: 'Inventory[Value]', note: 'That’s `Inventory[Value]` (column `F`): the numbers you’ll add up.' },
          { label: 'Warehouse column', at: 'Inventory[Warehouse]', note: 'That’s `Inventory[Warehouse]` (column `C`): SUMIFS reads it on every row to decide whether to add that row.' },
        ],
      },
      {
        do: `See what \`I2\` should add up: every ${first} row’s Value.`,
        why: `This is what you’d get by filtering Warehouse to ${first} and reading the status bar. SUMIFS does the same thing in one formula, and it stays up to date.`,
        show: [
          {
            label: `Select the ${first} values`,
            at: cellList('F', rows),
            note: `Look at **Sum** in the status bar at the bottom of the Excel window: ${money(total)}. That’s the number your formula in \`I2\` will show.`,
          },
        ],
      },
      typeStep({
        cell: 'I2',
        formula: [
          part('=SUMIFS(', 'Adds up only the rows that match the condition.'),
          part('Inventory[Value]', 'What to add up: the Value column.', 'Inventory[Value]'),
          raw(', '),
          part('Inventory[Warehouse]', 'Where to look: the Warehouse column.', 'Inventory[Warehouse]'),
          raw(', '),
          part('H2', `What to look for: the warehouse named in \`H2\` (${first}). Point at the cell instead of typing the name, so the same formula works on every row.`, 'H2'),
          raw(')'),
        ],
        why: TABLE_TYPING_TIP,
      }),
      fillStep({
        from: 'I2',
        range: 'I2:I5',
        direction: 'down',
        why: 'Fill Down copies `I2`’s formula into the cells below. Excel moves `H2` along to `H3`, `H4` and `H5`, so each row totals its own warehouse. The Table columns stay put.',
      }),
      checkStep(),
    ];
  },
  make: (rng) => ({ rows: Array.from({ length: 48 }, () => stockLine(rng)), order: rng.shuffle(WAREHOUSES) }),
  layout: (d) => ({
    blocks: [
      dataBlock('Inventory', 'A1', INV_COLS, invGrid(d.rows)),
      cells('H1', [['Warehouse', 'Inventory value']], 'header'),
      cells('H2', column(d.order), 'input'),
    ],
    answer: { kind: 'cells', range: 'I2:I5', format: FMT.currency, consistency: 'all' },
  }),
  expected: (d) => d.order.map((w) => [sum(d.rows.filter((r) => r.warehouse === w).map((r) => r.value))]),
  inputs: (d) => [tableWrite('Inventory', INV_COLS, invGrid(d.rows)), rangeWrite('H2:H5', column(d.order))],
  variants: [
    {
      label: 'stock levels change',
      apply: (d, rng) => ({
        ...d,
        rows: d.rows.map((r) => {
          const onHand = rng.int(0, 900);
          return { ...r, onHand, value: round(onHand * r.cost, 2) };
        }),
      }),
    },
    {
      label: 'the warehouse list in column H is reordered',
      explain: 'Point the criteria at H2 instead of typing a warehouse name.',
      apply: (d, rng) => {
        let order = rng.shuffle(d.order);
        while (order.join() === d.order.join()) order = rng.shuffle(d.order);
        return { ...d, order };
      },
    },
    {
      label: 'four new stock lines are added',
      explain: 'Use whole Table columns like Inventory[Value] so new rows are counted.',
      apply: (d, rng) => ({ ...d, rows: [...d.rows, ...Array.from({ length: 4 }, () => stockLine(rng))] }),
    },
  ],
  rules: {
    forbidText: { values: [...WAREHOUSES], advice: 'Point to the warehouse name in column H instead.' },
    allowNumbers: [],
  },
});

// ---------- Spend by vendor and month ----------

interface SpendLine {
  date: number;
  vendor: string;
  category: string;
  amount: number;
}

export const SPEND_COLS: ColumnSpec[] = [
  { header: 'Date', format: FMT.date },
  { header: 'Vendor' },
  { header: 'Category' },
  { header: 'Amount', format: FMT.currency },
];

export function spendLine(rng: Rng, vendors: readonly string[], from = serial(2026, 1, 1), to = serial(2026, 4, 30)): SpendLine {
  return { date: rng.int(from, to), vendor: rng.pick(vendors), category: rng.pick(CATEGORIES), amount: rng.float(40, 4800, 2) };
}

export const spendGrid = (rows: SpendLine[]) => rows.map((s) => [s.date, s.vendor, s.category, s.amount]);

const MONTHS = [serial(2026, 1, 1), serial(2026, 2, 1), serial(2026, 3, 1), serial(2026, 4, 1)];

interface GridData {
  rows: SpendLine[];
  vendors: string[];
  months: number[];
}

export const sumifsGrid = defineExercise<GridData>({
  id: 'sumifs-grid',
  module: 'sumifs',
  title: 'Spend by vendor and month',
  replaces: 'A summary grid where every cell has a slightly different hand-edited formula',
  minutes: 7,
  task: () =>
    'Fill `G2:J7` with each vendor’s spend (column `F`) in each month (row `1`; each header is the first day of the month). Write one formula in `G2` that works when filled both right and down.',
  concept: {
    summary:
      'A $ in a reference locks part of it when you fill. $F2 keeps the column fixed (vendors stay in F) while the row moves; G$1 keeps the row fixed (months stay in row 1) while the column moves.',
    syntax: '=SUMIFS(Spend[Amount], Spend[Vendor], $F2, Spend[Date], ">="&G$1, Spend[Date], "<="&EOMONTH(G$1,0))',
    example: '">="&G$1 builds the condition “on or after the date in row 1”. EOMONTH(G$1,0) is the last day of that month.',
    tip: 'While editing a reference, press {absKey} to cycle through $F$2, F$2, $F2 and F2.',
  },
  hints: [
    'Start with the vendor condition: Spend[Vendor] matched against $F2. The $ keeps it in column F when you fill right.',
    'For the month, use two date conditions: on or after G$1, and on or before EOMONTH(G$1,0).',
    'Join an operator and a cell with &: ">="&G$1. Press {absKey} while the cursor is on a reference to add the $. To fill right, select the row and press {fillRight} (Home › Fill › Right) rather than dragging the fill handle, which slides Table column names along.',
  ],
  solution: () => '=SUMIFS(Spend[Amount],Spend[Vendor],$F2,Spend[Date],">="&G$1,Spend[Date],"<="&EOMONTH(G$1,0))',
  guide: (d) => {
    const vendor = d.vendors[0];
    const month = d.months[0];
    const keep = (r: SpendLine) => r.vendor === vendor && r.date >= month && r.date <= eomonth(month);
    const total = sum(d.rows.filter(keep).map((r) => r.amount));
    return [
      {
        do: 'Meet the layout. Vendors run down column `F`, months run across row `1`, and the Table on the left is named **Spend**.',
        why: 'Each month header is really a date: the first day of that month, shown as “Jan 2026”. Every cell in the grid is one vendor in one month.',
        show: [
          { label: 'Vendors', at: 'F2:F7', note: 'Each row of the grid is one of these vendors.' },
          { label: 'Months', at: 'G1:J1', note: 'Each column of the grid is one of these months.' },
          { label: 'Amount column', at: 'Spend[Amount]', note: '`Spend[Amount]`: the numbers you’ll add up.' },
        ],
      },
      {
        do: `See what \`G2\` should add up: ${vendor}’s amounts dated in ${monthName(month)}.`,
        show: d.rows.some(keep)
          ? [
              {
                label: `Select ${vendor} in ${monthName(month)}`,
                at: cellList('D', rowsWhere(d.rows, keep)),
                note: `The status bar’s **Sum** is ${money(total)}. That’s the number for \`G2\`.`,
              },
            ]
          : undefined,
        why: 'That takes three conditions: the vendor matches, the date is on or after the 1st of the month, and the date is on or before the month’s last day.',
      },
      typeStep({
        cell: 'G2',
        formula: [
          part('=SUMIFS(', 'Adds up the rows that match every condition.'),
          part('Spend[Amount]', 'What to add up.', 'Spend[Amount]'),
          raw(', '),
          part('Spend[Vendor]', 'Condition 1 looks in the Vendor column…', 'Spend[Vendor]'),
          raw(', '),
          part('$F2', '…for the vendor in column `F`. The `$` before F keeps it in column F when you fill right.', 'F2'),
          raw(', '),
          part('Spend[Date]', 'Condition 2 looks in the Date column…', 'Spend[Date]'),
          raw(', '),
          part('">="&G$1', '…for dates on or after the month in row `1`. `&` joins the text ">=" to the date. The `$` before 1 keeps it in row 1 when you fill down.', 'G1'),
          raw(', '),
          part('Spend[Date]', 'Condition 3 looks in the Date column again…', 'Spend[Date]'),
          raw(', '),
          part('"<="&EOMONTH(G$1,0)', '…for dates on or before the last day of that month. `EOMONTH(G$1,0)` is the month’s last day.'),
          raw(')'),
        ],
        why: 'Adding the `$`: type it, or click inside a reference while typing and press {absKey} until it reads `$F2` or `G$1`.',
      }),
      fillStep({
        from: 'G2',
        range: 'G2:J7',
        direction: 'both',
        why: 'The `$` signs do the work here: every cell keeps reading its vendor from column `F` and its month from row `1`. Use the Fill commands rather than dragging the corner handle across, which would slide the Table column names along with it.',
      }),
      checkStep(),
    ];
  },
  make: (rng) => ({ rows: Array.from({ length: 90 }, () => spendLine(rng, VENDORS)), vendors: rng.shuffle(VENDORS), months: MONTHS }),
  layout: (d) => ({
    blocks: [
      dataBlock('Spend', 'A1', SPEND_COLS, spendGrid(d.rows)),
      cells('F1', [['Vendor']], 'header'),
      cells('G1', [d.months], 'input', FMT.month),
      cells('F2', column(d.vendors), 'input'),
    ],
    answer: { kind: 'cells', range: 'G2:J7', format: FMT.currency, consistency: 'all' },
  }),
  expected: (d) =>
    d.vendors.map((v) => d.months.map((m) => sum(d.rows.filter((r) => r.vendor === v && r.date >= m && r.date <= eomonth(m)).map((r) => r.amount)))),
  inputs: (d) => [tableWrite('Spend', SPEND_COLS, spendGrid(d.rows)), rangeWrite('G1:J1', [d.months]), rangeWrite('F2:F7', column(d.vendors))],
  variants: [
    {
      label: 'the amounts change',
      apply: (d, rng) => ({ ...d, rows: d.rows.map((r) => ({ ...r, amount: round(r.amount * rng.float(0.4, 1.9, 3), 2) })) }),
    },
    {
      label: 'the vendor list in column F is reordered',
      explain: 'Lock the column with $F2 so every cell reads its vendor from column F.',
      apply: (d, rng) => ({ ...d, vendors: rng.shuffle(d.vendors) }),
    },
    {
      label: 'the months in row 1 are reordered',
      explain: 'Lock the row with G$1 so every cell reads its month from row 1.',
      apply: (d, rng) => {
        let months = rng.shuffle(d.months);
        while (months.join() === d.months.join()) months = rng.shuffle(d.months);
        return { ...d, months };
      },
    },
  ],
  rules: {
    forbidText: { values: [...VENDORS], advice: 'Point to the vendor in column F instead.' },
    allowNumbers: [0, 1],
  },
});

// ---------- This month's revenue by rep ----------

interface SaleLine {
  date: number;
  region: string;
  rep: string;
  product: string;
  units: number;
  revenue: number;
}

export const SALES_COLS: ColumnSpec[] = [
  { header: 'Date', format: FMT.date },
  { header: 'Region' },
  { header: 'Rep' },
  { header: 'Product' },
  { header: 'Units', format: FMT.int },
  { header: 'Revenue', format: FMT.currency },
];

const PRICE: Record<string, number> = { Basic: 20, Plus: 35, Pro: 60, Max: 95 };

export function saleLine(rng: Rng, reps: readonly string[] = REPS): SaleLine {
  const product = rng.pick(PRODUCTS);
  const units = rng.int(1, 40);
  return {
    date: rng.int(serial(2026, 1, 1), serial(2026, 6, 30)),
    region: rng.pick(REGIONS),
    rep: rng.pick(reps),
    product,
    units,
    revenue: round(units * PRICE[product] * rng.float(0.85, 1, 2), 2),
  };
}

export const salesGrid = (rows: SaleLine[]) => rows.map((s) => [s.date, s.region, s.rep, s.product, s.units, s.revenue]);

const SALE_MONTHS = [1, 2, 3, 4, 5, 6].map((m) => serial(2026, m, 1));
const SORTED_REPS = [...REPS].sort();

interface MonthData {
  rows: SaleLine[];
  month: number;
}

export const sumifsMonth = defineExercise<MonthData>({
  id: 'sumifs-month',
  module: 'sumifs',
  title: 'This month’s revenue by rep',
  replaces: 'Retyping date ranges every month-end',
  minutes: 5,
  task: () =>
    'In `I4:I11`, show each rep’s revenue for the month in `I1`. When someone changes `I1`, every total should update by itself.',
  concept: {
    summary:
      'Date windows are two conditions on the same column: on or after the first day, and on or before the last day. EOMONTH gives you the last day of any month.',
    syntax: '=SUMIFS(Sales[Revenue], Sales[Rep], H4, Sales[Date], ">="&$I$1, Sales[Date], "<="&EOMONTH($I$1,0))',
    example: 'EOMONTH($I$1, 0) is the last day of I1’s month; EOMONTH($I$1, -1) would be the last day of the month before.',
  },
  hints: [
    'You need three conditions: the rep, a start date and an end date. The start and end both test Sales[Date].',
    'Build each date condition by joining an operator to a cell: ">="&$I$1.',
    'The end of the month is EOMONTH($I$1,0). Lock I1 with $ so it doesn’t move when you fill down.',
  ],
  solution: () => '=SUMIFS(Sales[Revenue],Sales[Rep],H4,Sales[Date],">="&$I$1,Sales[Date],"<="&EOMONTH($I$1,0))',
  guide: (d) => {
    const rep = SORTED_REPS[0];
    const keep = (r: SaleLine) => r.rep === rep && r.date >= d.month && r.date <= eomonth(d.month);
    const total = sum(d.rows.filter(keep).map((r) => r.revenue));
    return [
      {
        do: 'Meet the layout. `I1` holds the month, the reps run down `H4:H11`, and the Table on the left is named **Sales**.',
        why: `\`I1\` is really a date: the first day of ${monthName(d.month)}. Your formulas will read it, so changing \`I1\` updates every total.`,
        show: [
          { label: 'The month', at: 'I1' },
          { label: 'Revenue column', at: 'Sales[Revenue]', note: '`Sales[Revenue]`: the numbers you’ll add up.' },
          { label: 'Date column', at: 'Sales[Date]', note: '`Sales[Date]`: tested twice, once for the start of the month and once for the end.' },
        ],
      },
      {
        do: `See what \`I4\` should add up: ${rep}’s revenue dated in ${monthName(d.month)}.`,
        show: d.rows.some(keep)
          ? [
              {
                label: `Select ${rep} in ${monthName(d.month)}`,
                at: cellList('F', rowsWhere(d.rows, keep)),
                note: `The status bar’s **Sum** is ${money(total)}. That’s the number for \`I4\`.`,
              },
            ]
          : undefined,
      },
      typeStep({
        cell: 'I4',
        formula: [
          part('=SUMIFS(', 'Adds up the rows that match every condition.'),
          part('Sales[Revenue]', 'What to add up.', 'Sales[Revenue]'),
          raw(', '),
          part('Sales[Rep]', 'Condition 1 looks in the Rep column…', 'Sales[Rep]'),
          raw(', '),
          part('H4', '…for the rep named in `H4`. No `$`, so it moves to `H5`, `H6`… as you fill down.', 'H4'),
          raw(', '),
          part('Sales[Date]', 'Condition 2 looks in the Date column…', 'Sales[Date]'),
          raw(', '),
          part('">="&$I$1', '…for dates on or after the first of the month in `I1`. `$I$1` is locked both ways, so every row reads `I1`.', 'I1'),
          raw(', '),
          part('Sales[Date]', 'Condition 3 looks in the Date column again…', 'Sales[Date]'),
          raw(', '),
          part('"<="&EOMONTH($I$1,0)', '…for dates on or before the month’s last day. `EOMONTH($I$1,0)` works it out from `I1`.'),
          raw(')'),
        ],
        why: 'Adding the `$`: click inside `I1` while typing and press {absKey} until it reads `$I$1`.',
      }),
      fillStep({ from: 'I4', range: 'I4:I11', direction: 'down', why: '`H4` moves down to each rep in turn; `$I$1` stays on the month.' }),
      checkStep(),
    ];
  },
  make: (rng) => ({ rows: Array.from({ length: 120 }, () => saleLine(rng)), month: rng.pick(SALE_MONTHS.slice(1, 5)) }),
  layout: (d) => ({
    blocks: [
      dataBlock('Sales', 'A1', SALES_COLS, salesGrid(d.rows)),
      cells('H1', [['Month']], 'label'),
      cells('I1', [[d.month]], 'input', FMT.month),
      cells('H3', [['Rep', 'Revenue']], 'header'),
      cells('H4', column(SORTED_REPS), 'input'),
    ],
    answer: { kind: 'cells', range: 'I4:I11', format: FMT.currency, consistency: 'all' },
  }),
  expected: (d) =>
    SORTED_REPS.map((rep) => [sum(d.rows.filter((r) => r.rep === rep && r.date >= d.month && r.date <= eomonth(d.month)).map((r) => r.revenue))]),
  inputs: (d) => [tableWrite('Sales', SALES_COLS, salesGrid(d.rows)), rangeWrite('I1', [[d.month]])],
  variants: [
    {
      label: 'the month in I1 changes',
      explain: 'Build the date window from I1 and EOMONTH instead of typing dates.',
      apply: (d, rng) => ({ ...d, month: rng.pick(SALE_MONTHS.filter((m) => m !== d.month)) }),
    },
    {
      label: 'revenue changes',
      apply: (d, rng) => ({ ...d, rows: d.rows.map((r) => ({ ...r, revenue: round(r.revenue * rng.float(0.5, 1.7, 3), 2) })) }),
    },
    {
      label: 'new sales are added',
      explain: 'Use whole Table columns so new rows are counted.',
      apply: (d, rng) => ({ ...d, rows: [...d.rows, ...Array.from({ length: 12 }, () => ({ ...saleLine(rng), date: rng.int(d.month, eomonth(d.month)) }))] }),
    },
  ],
  rules: {
    forbidText: { values: [...REPS], advice: 'Point to the rep’s name in column H instead.' },
    allowNumbers: [0, 1],
  },
});

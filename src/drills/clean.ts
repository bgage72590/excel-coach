import { DEPARTMENTS, FIRST_NAMES, ITEMS, LAST_NAMES, REGIONS, REPS, WAREHOUSES, excelTextCompare } from '../engine/data';
import { round, type Rng } from '../engine/rng';
import type { Grid } from '../engine/types';
import { cells, dataBlock } from '../exercises/common';
import { drillItem, headers, lastRow, removeDuplicateRows, sheetCheck } from './common';
import type { Drill } from './types';

// ---------- sorting: orders with unique amounts, so every sort has one right answer ----------

export interface SaleRow {
  order: string;
  rep: string;
  region: string;
  units: number;
  amount: number;
}

export interface SalesData {
  rows: SaleRow[];
}

export const SALE_COLS = headers('Order', 'Rep', 'Region', 'Units', 'Amount');

function sales(rng: Rng): SalesData {
  const count = rng.int(9, 12);
  const first = rng.int(20, 70) * 100;
  const used = new Set<number>();
  const rows = Array.from({ length: count }, (_, i) => {
    let amount = rng.float(120, 4800);
    // No ties: a tie would let Excel keep either row first.
    while (used.has(amount)) amount = round(amount + 0.01, 2);
    used.add(amount);
    return { order: `SO-${first + i}`, rep: rng.pick(REPS), region: rng.pick(REGIONS), units: rng.int(1, 40), amount };
  });
  return { rows };
}

export const saleGrid = (rows: SaleRow[]): Grid => rows.map((r) => [r.order, r.rep, r.region, r.units, r.amount]);

export const byAmountDesc = (rows: SaleRow[]): SaleRow[] => [...rows].sort((a, b) => b.amount - a.amount);

/** Region A to Z (Excel ignores case), then Amount largest first. */
export const byRegionThenAmount = (rows: SaleRow[]): SaleRow[] => [...rows].sort((a, b) => excelTextCompare(a.region, b.region) || b.amount - a.amount);

/** Data that isn't already in the target order, so the check can't pass before the learner acts. */
function unsorted(d: SalesData, sort: (rows: SaleRow[]) => SaleRow[]): SalesData {
  const target = sort(d.rows);
  return target.every((r, i) => r === d.rows[i]) ? { rows: [...d.rows].reverse() } : d;
}

const salesBlocks = (d: SalesData) => [dataBlock('Sales', 'A1', SALE_COLS, saleGrid(d.rows), false)];

const sortLargestFirst = drillItem<SalesData>({
  id: 'drill-sort-desc',
  title: 'Sort largest first',
  replaces: 'Scanning a column by eye for the biggest orders',
  task: () => 'Sort the rows by Amount, largest first.',
  concept: {
    summary: 'Sorting on one column reorders whole rows, so each order’s details stay together.',
    syntax: 'Data › Sort Z to A',
  },
  hints: ['Click one cell in the Amount column. Don’t select the whole column.', 'Select Data › Sort Z to A, which sorts numbers largest to smallest.'],
  solution: () => 'Click a cell in Amount › Data › Sort Z to A',
  make: (rng) => unsorted(sales(rng), byAmountDesc),
  blocks: salesBlocks,
  inspections: (d) => [
    sheetCheck(
      { kind: 'values', range: `A2:E${lastRow(d.rows)}`, expected: saleGrid(byAmountDesc(d.rows)), describe: 'the order rows' },
      'Rows run from the largest Amount down',
      'Click one cell in Amount before you sort, so whole rows move together.',
    ),
  ],
  shortcut: { mac: 'Data › Sort Z to A', windows: 'Alt, A, S, D' },
});

const sortTwoKeys = drillItem<SalesData>({
  id: 'drill-sort-two-keys',
  title: 'Sort by two columns',
  replaces: 'Sorting twice and hoping the first order survives',
  task: () => 'Sort the rows by Region from A to Z, then by Amount, largest first, within each region.',
  concept: {
    summary: 'The Sort box sorts by several columns in turn. Each level orders the rows that tie on the level above it.',
    syntax: 'Data › Sort › Add Level',
  },
  hints: ['Click a cell in the data, then select Data › Sort.', 'Sort by Region, A to Z. Add a level (the + button on a Mac), then sort by Amount, Largest to Smallest.'],
  solution: () => 'Data › Sort › Region, A to Z › Add Level › Amount, Largest to Smallest',
  make: (rng) => unsorted(sales(rng), byRegionThenAmount),
  blocks: salesBlocks,
  inspections: (d) => [
    sheetCheck(
      { kind: 'values', range: `A2:E${lastRow(d.rows)}`, expected: saleGrid(byRegionThenAmount(d.rows)), describe: 'the order rows' },
      'Rows sorted by Region, then by Amount',
    ),
  ],
  shortcut: { mac: 'Data › Sort', windows: 'Alt, A, S, S' },
});

// ---------- duplicates: an export that repeated some rows ----------

const CUSTOMERS = ['Northgate Hardware', 'Lakeside Builders', 'Pinecrest Facilities', 'Harbor Logistics', 'Cedar Ridge Clinic', 'Maple Street Market'] as const;

export interface LineRow {
  order: string;
  customer: string;
  item: string;
  qty: number;
  amount: number;
}

export interface LinesData {
  rows: LineRow[];
}

export const LINE_COLS = headers('Order', 'Customer', 'Item', 'Qty', 'Amount');

/** Unique orders, then copies of two or three of them dropped in somewhere below the original. */
function withRepeats(rng: Rng): LinesData {
  const count = rng.int(7, 9);
  const first = rng.int(30, 80) * 100;
  const rows: LineRow[] = Array.from({ length: count }, (_, i) => {
    const it = rng.pick(ITEMS);
    const qty = rng.int(2, 48);
    return { order: `SO-${first + i}`, customer: rng.pick(CUSTOMERS), item: it.item, qty, amount: round(qty * it.cost * rng.float(1.3, 1.7, 3), 2) };
  });
  for (const original of rng.sample(rows.slice(), rng.int(2, 3))) {
    const at = rng.int(rows.indexOf(original) + 1, rows.length);
    rows.splice(at, 0, { ...original });
  }
  return { rows };
}

export const lineGrid = (rows: LineRow[]): Grid => rows.map((r) => [r.order, r.customer, r.item, r.qty, r.amount]);

const removeDuplicates = drillItem<LinesData>({
  id: 'drill-remove-duplicates',
  title: 'Remove duplicate rows',
  replaces: 'Hunting for repeated rows by eye and deleting them one at a time',
  task: () => 'The export repeated some orders. Remove the duplicate rows.',
  concept: {
    summary: 'Remove Duplicates deletes each row that repeats an earlier one, keeps the first copy, and moves the rows below it up.',
    syntax: 'Data › Remove Duplicates',
  },
  hints: ['Click any cell in the data.', 'Select Data › Remove Duplicates, leave every column selected, then select OK.'],
  solution: () => 'Click a cell in the data › Data › Remove Duplicates › OK',
  make: withRepeats,
  blocks: (d) => [dataBlock('Lines', 'A1', LINE_COLS, lineGrid(d.rows), false)],
  inspections: (d) => [
    sheetCheck(
      { kind: 'values', range: `A2:E${lastRow(d.rows)}`, expected: removeDuplicateRows(lineGrid(d.rows)), describe: 'the order rows' },
      'Each order appears once',
      'Leave every column selected in Remove Duplicates, so only exact repeats go.',
    ),
  ],
  shortcut: { mac: 'Data › Remove Duplicates', windows: 'Alt, A, M' },
});

// ---------- formulas: fill down, then freeze as values ----------

export interface PriceRow {
  order: string;
  item: string;
  qty: number;
  price: number;
}

export interface PricesData {
  rows: PriceRow[];
}

export const PRICE_COLS = headers('Order', 'Item', 'Qty', 'Price');

function priceLines(rng: Rng): PricesData {
  const count = rng.int(8, 12);
  const first = rng.int(30, 80) * 100;
  return {
    rows: Array.from({ length: count }, (_, i) => {
      const it = rng.pick(ITEMS);
      return { order: `SO-${first + i}`, item: it.item, qty: rng.int(2, 48), price: round(it.cost * rng.float(1.3, 1.8, 3), 2) };
    }),
  };
}

export const priceGrid = (rows: PriceRow[]): Grid => rows.map((r) => [r.order, r.item, r.qty, r.price]);

/** What =C2*D2 shows on each row. */
export const lineTotals = (rows: PriceRow[]): Grid => rows.map((r) => [r.qty * r.price]);

const totalFormula = (row: number) => `=C${row}*D${row}`;

const fillDown = drillItem<PricesData>({
  id: 'drill-fill-down',
  title: 'Fill a formula down',
  replaces: 'Retyping the same formula on every row',
  task: (d) => `Copy the formula in \`E2\` down to \`E${lastRow(d.rows)}\`.`,
  concept: {
    summary: 'Filling down copies a formula to the cells below. Its references shift row by row, so each row works out its own total.',
    syntax: 'Select the formula and the cells below it › Home › Fill › Down',
  },
  hints: ['Select `E2` and the empty cells below it, down to the last row of data.', 'Select Home › Fill › Down. Or double-click the fill handle, the small square at the corner of `E2`.'],
  solution: (d) => `Select E2:E${lastRow(d.rows)} › Home › Fill › Down`,
  make: priceLines,
  blocks: (d) => [dataBlock('Lines', 'A1', PRICE_COLS, priceGrid(d.rows), false), cells('E1', [['Total']], 'input'), cells('E2', [[totalFormula(2)]], 'formula')],
  inspections: (d) => [
    sheetCheck({ kind: 'formulas', range: `E2:E${lastRow(d.rows)}`, formulas: true }, 'Every Total cell has a formula'),
    sheetCheck({ kind: 'values', range: `E2:E${lastRow(d.rows)}`, expected: lineTotals(d.rows), describe: 'the totals' }, 'Each total is its own row’s Qty × Price'),
  ],
  shortcut: { mac: '⌘D, or double-click the fill handle', windows: 'Ctrl+D, or double-click the fill handle' },
});

const pasteValues = drillItem<PricesData>({
  id: 'drill-paste-values',
  title: 'Replace formulas with values',
  replaces: 'Retyping totals so they stop changing when prices do',
  task: (d) => `Replace the formulas in \`E2:E${lastRow(d.rows)}\` with the values they show.`,
  concept: {
    summary: 'Paste Special › Values swaps formulas for the results they show. Use it to fix results in place before you share or reuse them.',
    syntax: 'Copy › Paste Special › Values',
  },
  hints: ['Select the totals and copy them.', 'Leave the selection where it is, open Paste Special, choose Values, then select OK.'],
  solution: (d) => `Copy E2:E${lastRow(d.rows)} › Paste Special › Values`,
  make: priceLines,
  blocks: (d) => [
    dataBlock('Lines', 'A1', PRICE_COLS, priceGrid(d.rows), false),
    cells('E1', [['Total']], 'input'),
    cells('E2', d.rows.map((_, i) => [totalFormula(i + 2)]), 'formula'),
  ],
  inspections: (d) => [
    sheetCheck({ kind: 'formulas', range: `E2:E${lastRow(d.rows)}`, formulas: false }, 'The totals are values, not formulas'),
    sheetCheck({ kind: 'values', range: `E2:E${lastRow(d.rows)}`, expected: lineTotals(d.rows), describe: 'the totals' }, 'The totals still show the same amounts'),
  ],
  shortcut: { mac: '⌘C, then ⌃⌘V › Values', windows: 'Ctrl+C, then Ctrl+Alt+V, V, Enter' },
});

// ---------- splitting: full names into first and last ----------

export interface StaffRow {
  first: string;
  last: string;
  department: string;
}

export interface StaffData {
  rows: StaffRow[];
}

export const STAFF_COLS = headers('First name', 'Last name', 'Department');

/** Distinct one-word first and last names, so each splits at its only space. */
function staff(rng: Rng): StaffData {
  const count = rng.int(8, 12);
  const lasts = rng.sample(LAST_NAMES, count);
  return { rows: lasts.map((last) => ({ first: rng.pick(FIRST_NAMES), last, department: rng.pick(DEPARTMENTS) })) };
}

/** Full names in A, with B left empty for the split. */
export const staffGrid = (rows: StaffRow[]): Grid => rows.map((r) => [`${r.first} ${r.last}`, '', r.department]);

const splitNames = drillItem<StaffData>({
  id: 'drill-split-names',
  title: 'Split names into two columns',
  replaces: 'Retyping last names into their own column',
  task: (d) => `Split the full names in \`A2:A${lastRow(d.rows)}\` so the last names move to column \`B\`.`,
  concept: {
    summary: 'Text to Columns splits each cell at a delimiter, such as a space or a comma, and puts the parts in the columns to the right.',
    syntax: 'Data › Text to Columns › Delimited › Space',
  },
  hints: ['Select the full names, without the header, then select Data › Text to Columns.', 'Select Delimited, then Next. Select the Space check box, then Finish.'],
  solution: (d) => `Select A2:A${lastRow(d.rows)} › Data › Text to Columns › Delimited › Space › Finish`,
  make: staff,
  blocks: (d) => [dataBlock('Staff', 'A1', STAFF_COLS, staffGrid(d.rows), false)],
  inspections: (d) => [
    sheetCheck(
      { kind: 'values', range: `A2:B${lastRow(d.rows)}`, expected: d.rows.map((r) => [r.first, r.last]), describe: 'the names' },
      'First names in column A, last names in column B',
    ),
  ],
  shortcut: { mac: 'Data › Text to Columns', windows: 'Alt, A, E' },
});

// ---------- replacing: a warehouse that moved ----------

export const OLD_WAREHOUSE = 'Columbus';
export const NEW_WAREHOUSE = 'Cincinnati';

export interface TransferRow {
  transfer: string;
  warehouse: string;
  item: string;
  units: number;
}

export interface TransfersData {
  rows: TransferRow[];
}

export const TRANSFER_COLS = headers('Transfer', 'Warehouse', 'Item', 'Units');

function transfers(rng: Rng): TransfersData {
  const count = rng.int(8, 12);
  const first = rng.int(10, 60) * 100;
  const others = WAREHOUSES.filter((w) => w !== OLD_WAREHOUSE);
  const moved = new Set(rng.sample(Array.from({ length: count }, (_, i) => i), rng.int(2, 4)));
  return {
    rows: Array.from({ length: count }, (_, i) => ({
      transfer: `TR-${first + i}`,
      warehouse: moved.has(i) ? OLD_WAREHOUSE : rng.pick(others),
      item: rng.pick(ITEMS).item,
      units: rng.int(5, 400),
    })),
  };
}

export const transferGrid = (rows: TransferRow[]): Grid => rows.map((r) => [r.transfer, r.warehouse, r.item, r.units]);

const findReplace = drillItem<TransfersData>({
  id: 'drill-find-replace',
  title: 'Find and replace',
  replaces: 'Editing the same word in cell after cell',
  task: () => `The ${OLD_WAREHOUSE} warehouse moved to ${NEW_WAREHOUSE}. Replace every “${OLD_WAREHOUSE}” with “${NEW_WAREHOUSE}”.`,
  concept: {
    summary: 'Replace All changes every match in one step, which is faster and safer than editing cells one at a time.',
    syntax: 'Home › Find & Select › Replace',
  },
  hints: ['Select Home › Find & Select › Replace.', `Find “${OLD_WAREHOUSE}”, replace it with “${NEW_WAREHOUSE}”, then select Replace All.`],
  solution: () => `Home › Find & Select › Replace › ${OLD_WAREHOUSE} › ${NEW_WAREHOUSE} › Replace All`,
  make: transfers,
  blocks: (d) => [dataBlock('Transfers', 'A1', TRANSFER_COLS, transferGrid(d.rows), false)],
  inspections: (d) => [
    sheetCheck(
      {
        kind: 'values',
        range: `A2:D${lastRow(d.rows)}`,
        expected: transferGrid(d.rows.map((r) => ({ ...r, warehouse: r.warehouse === OLD_WAREHOUSE ? NEW_WAREHOUSE : r.warehouse }))),
        describe: 'the transfers',
      },
      `Every ${OLD_WAREHOUSE} now reads ${NEW_WAREHOUSE}`,
    ),
  ],
  shortcut: { mac: '⌃H', windows: 'Ctrl+H' },
});

export const sortAndClean: Drill = {
  id: 'sort-clean',
  title: 'Sort and clean',
  blurb: 'Put rows in order and tidy up what an export leaves behind.',
  items: [sortLargestFirst, removeDuplicates, fillDown, pasteValues, sortTwoKeys, splitNames, findReplace],
  parSeconds: 120,
};

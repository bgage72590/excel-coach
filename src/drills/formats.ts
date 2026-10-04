import { CARRIERS, CATEGORIES, FIRST_NAMES, GL_ACCOUNTS, ITEMS, LAST_NAMES, WAREHOUSES, serial } from '../engine/data';
import { round, type Rng } from '../engine/rng';
import type { ColumnSpec } from '../engine/types';
import { dataBlock } from '../exercises/common';
import { DATE_FORMAT, GENERAL_FORMAT, PERCENT_FORMAT, RED_NEGATIVE_FORMAT, THOUSANDS_FORMAT, TIME_FORMAT, drillItem, headers, lastRow, sheetCheck } from './common';
import type { Drill } from './types';

/**
 * Each item is a short export with one column to fix. The other columns stay at General, so the
 * sheet looks like it came straight out of another system.
 */

/** Each item's data: the rows under its header. */
interface Rows<R> {
  rows: R[];
}

// ---------- percentages ----------

export interface DiscountRow {
  item: string;
  price: number;
  discount: number;
}

const DISCOUNT_COLS = headers('Item', 'List price', 'Discount');

function discounts(rng: Rng): Rows<DiscountRow> {
  return {
    rows: rng.sample(ITEMS, rng.int(8, 10)).map((it) => ({ item: it.item, price: round(it.cost * rng.float(1.3, 1.8, 3), 2), discount: rng.int(10, 60) / 200 })),
  };
}

const showPercent = drillItem<Rows<DiscountRow>>({
  id: 'drill-percent-format',
  title: 'Show percentages',
  replaces: 'Multiplying rates by 100 in a helper column',
  task: (d) => `Show the discounts in \`C2:C${lastRow(d.rows)}\` as percentages.`,
  concept: {
    summary: 'The percentage format shows a value times 100 with a percent sign, so 0.15 shows as 15%. The value itself doesn’t change.',
    syntax: 'Home › Percent Style',
  },
  hints: ['Select the discounts in column C.', 'Select Percent Style (%) on the Home tab.'],
  solution: (d) => `Select C2:C${lastRow(d.rows)} › Home › Percent Style`,
  make: discounts,
  blocks: (d) => [dataBlock('Discounts', 'A1', DISCOUNT_COLS, d.rows.map((r) => [r.item, r.price, r.discount]), false)],
  inspections: (d) => [
    sheetCheck({ kind: 'numberFormat', range: `C2:C${lastRow(d.rows)}`, matches: PERCENT_FORMAT, describe: 'a percentage' }, 'Discounts show as percentages'),
  ],
  shortcut: { mac: '⌃⇧%', windows: 'Ctrl+Shift+%' },
});

// ---------- dates ----------

export interface ShipRow {
  order: string;
  shipped: number;
  carrier: string;
}

const SHIP_COLS = headers('Order', 'Ship date', 'Carrier');

function shipments(rng: Rng): Rows<ShipRow> {
  const first = rng.int(20, 90) * 100;
  return {
    rows: Array.from({ length: rng.int(8, 10) }, (_, i) => ({
      order: `SO-${first + i}`,
      shipped: rng.int(serial(2026, 1, 5), serial(2026, 12, 18)),
      carrier: rng.pick(CARRIERS),
    })),
  };
}

const showDates = drillItem<Rows<ShipRow>>({
  id: 'drill-date-format',
  title: 'Show dates',
  replaces: 'Retyping dates that arrived as numbers',
  task: (d) => `The ship dates in \`B2:B${lastRow(d.rows)}\` came through as numbers. Format them as dates.`,
  concept: {
    summary: 'Excel stores a date as a count of days since 1900. A date format shows that count as a calendar date.',
    syntax: 'Home › Number Format › Short Date',
  },
  hints: ['Select the ship dates in column B.', 'Choose Short Date in the Number Format box on the Home tab.'],
  solution: (d) => `Select B2:B${lastRow(d.rows)} › Home › Number Format › Short Date`,
  make: shipments,
  blocks: (d) => [dataBlock('Shipments', 'A1', SHIP_COLS, d.rows.map((r) => [r.order, r.shipped, r.carrier]), false)],
  inspections: (d) => [sheetCheck({ kind: 'numberFormat', range: `B2:B${lastRow(d.rows)}`, matches: DATE_FORMAT, describe: 'a date format' }, 'Ship dates show as dates')],
  shortcut: { mac: '⌃⇧#', windows: 'Ctrl+Shift+#' },
});

// ---------- times ----------

/** Shifts and the hours their start times fall between. */
const SHIFTS = [
  { name: 'Early', from: 5, to: 7 },
  { name: 'Day', from: 8, to: 10 },
  { name: 'Late', from: 14, to: 16 },
] as const;

export interface ShiftRow {
  employee: string;
  shift: string;
  /** Fraction of a day, on the quarter hour: 0.3125 is 7:30 AM. */
  start: number;
}

const SHIFT_COLS = headers('Employee', 'Shift', 'Start');

function shifts(rng: Rng): Rows<ShiftRow> {
  return {
    rows: rng.sample(LAST_NAMES, rng.int(8, 10)).map((last) => {
      const shift = rng.pick(SHIFTS);
      return { employee: `${rng.pick(FIRST_NAMES)} ${last}`, shift: shift.name, start: rng.int(shift.from * 4, shift.to * 4) / 96 };
    }),
  };
}

const showTimes = drillItem<Rows<ShiftRow>>({
  id: 'drill-time-format',
  title: 'Show times',
  replaces: 'Converting decimals to clock times by hand',
  task: (d) => `The start times in \`C2:C${lastRow(d.rows)}\` show as decimals. Format them as times.`,
  concept: {
    summary: 'Excel stores a time as a fraction of a day, so 0.5 is noon. A time format shows the fraction as hours and minutes.',
    syntax: 'Home › Number Format › Time',
  },
  hints: ['Select the start times in column C.', 'Choose Time in the Number Format box on the Home tab.'],
  solution: (d) => `Select C2:C${lastRow(d.rows)} › Home › Number Format › Time`,
  make: shifts,
  blocks: (d) => [dataBlock('Shifts', 'A1', SHIFT_COLS, d.rows.map((r) => [r.employee, r.shift, r.start]), false)],
  inspections: (d) => [sheetCheck({ kind: 'numberFormat', range: `C2:C${lastRow(d.rows)}`, matches: TIME_FORMAT, describe: 'a time format' }, 'Start times show as times')],
  shortcut: { mac: '⌃⇧@', windows: 'Ctrl+Shift+@' },
});

// ---------- thousands separators ----------

export interface StockRow {
  item: string;
  category: string;
  warehouse: string;
  units: number;
}

const STOCK_COLS = headers('Item', 'Category', 'Warehouse', 'Units');

function stock(rng: Rng): Rows<StockRow> {
  return {
    rows: Array.from({ length: rng.int(8, 10) }, () => {
      const category = rng.pick(CATEGORIES);
      return { item: rng.pick(ITEMS.filter((it) => it.category === category)).item, category, warehouse: rng.pick(WAREHOUSES), units: rng.int(1200, 250_000) };
    }),
  };
}

const addThousands = drillItem<Rows<StockRow>>({
  id: 'drill-thousands-format',
  title: 'Add a thousands separator',
  replaces: 'Counting digits to read a large number',
  task: (d) => `Add a thousands separator to the units in \`D2:D${lastRow(d.rows)}\`.`,
  concept: {
    summary: 'A thousands separator groups digits so 48250 reads as 48,250. The value doesn’t change.',
    syntax: 'Home › Comma Style',
  },
  hints: ['Select the units in column D.', 'Select Comma Style on the Home tab.'],
  solution: (d) => `Select D2:D${lastRow(d.rows)} › Home › Comma Style`,
  make: stock,
  blocks: (d) => [dataBlock('Stock', 'A1', STOCK_COLS, d.rows.map((r) => [r.item, r.category, r.warehouse, r.units]), false)],
  inspections: (d) => [
    sheetCheck(
      { kind: 'numberFormat', range: `D2:D${lastRow(d.rows)}`, matches: THOUSANDS_FORMAT, describe: 'a number format with a thousands separator' },
      'Units have a thousands separator',
    ),
  ],
  shortcut: { mac: '⌃⇧!', windows: 'Ctrl+Shift+!' },
});

// ---------- negatives in red ----------

export interface VarianceRow {
  account: string;
  budget: number;
  actual: number;
  variance: number;
}

const VARIANCE_COLS = headers('Account', 'Budget', 'Actual', 'Variance');

/** Operating expense lines, at least two over budget and two under. */
function variances(rng: Rng): Rows<VarianceRow> {
  const accounts = rng.sample(
    GL_ACCOUNTS.filter((a) => a.type === 'Opex').map((a) => a.name),
    rng.int(8, 10),
  );
  const over = new Set(rng.sample(Array.from({ length: accounts.length }, (_, i) => i), rng.int(2, accounts.length - 2)));
  return {
    rows: accounts.map((account, i) => {
      const budget = rng.int(20, 400) * 100;
      const actual = round(budget * (over.has(i) ? rng.float(1.02, 1.25, 3) : rng.float(0.78, 0.98, 3)), 2);
      // Under budget is negative: Actual minus Budget.
      return { account, budget, actual, variance: round(actual - budget, 2) };
    }),
  };
}

const redNegatives = drillItem<Rows<VarianceRow>>({
  id: 'drill-red-negatives',
  title: 'Show negatives in red',
  replaces: 'Coloring negative cells red one at a time',
  task: (d) => `Give the variances in \`D2:D${lastRow(d.rows)}\` a number format that shows negatives in red.`,
  concept: {
    summary: 'A number format can show positive and negative numbers differently. With [Red] in its negative section, every negative shows in red, including ones added later.',
    syntax: 'Format Cells › Number › Negative numbers',
  },
  hints: ['Select the variances in column D, right-click and choose Format Cells.', 'On the Number tab, choose Number, pick a red option under Negative numbers, then select OK.'],
  solution: (d) => `Select D2:D${lastRow(d.rows)} › Format Cells › Number › a red Negative numbers option`,
  make: variances,
  blocks: (d) => [dataBlock('Variance', 'A1', VARIANCE_COLS, d.rows.map((r) => [r.account, r.budget, r.actual, r.variance]), false)],
  inspections: (d) => [
    sheetCheck(
      { kind: 'numberFormat', range: `D2:D${lastRow(d.rows)}`, matches: RED_NEGATIVE_FORMAT, describe: 'a format that shows negatives in red' },
      'Negative variances show in red',
    ),
  ],
  shortcut: { mac: '⌘1 › Number', windows: 'Ctrl+1 › Number' },
});

// ---------- back to General ----------

export interface OrderNoRow {
  order: number;
  item: string;
  qty: number;
}

/** The order numbers arrive formatted as dates, a common import mistake. */
export const MISTAKEN_DATE = 'm/d/yy';

const ORDER_NO_COLS: ColumnSpec[] = [{ header: 'Order no.', format: MISTAKEN_DATE }, { header: 'Item' }, { header: 'Qty' }];

function orderNumbers(rng: Rng): Rows<OrderNoRow> {
  const first = rng.int(45_000, 46_400);
  return { rows: Array.from({ length: rng.int(8, 10) }, (_, i) => ({ order: first + i * rng.int(1, 3), item: rng.pick(ITEMS).item, qty: rng.int(1, 60) })) };
}

const backToGeneral = drillItem<Rows<OrderNoRow>>({
  id: 'drill-general-format',
  title: 'Back to General',
  replaces: 'Retyping numbers that turned into dates',
  task: (d) => `The order numbers in \`A2:A${lastRow(d.rows)}\` show as dates by mistake. Set them back to the General format.`,
  concept: {
    summary: 'General is the format a new cell starts with. Numbers show as typed, without symbols, separators or dates.',
    syntax: 'Home › Number Format › General',
  },
  hints: ['Select the order numbers in column A.', 'Choose General in the Number Format box on the Home tab.'],
  solution: (d) => `Select A2:A${lastRow(d.rows)} › Home › Number Format › General`,
  make: orderNumbers,
  blocks: (d) => [dataBlock('OrderNumbers', 'A1', ORDER_NO_COLS, d.rows.map((r) => [r.order, r.item, r.qty]), false)],
  inspections: (d) => [sheetCheck({ kind: 'numberFormat', range: `A2:A${lastRow(d.rows)}`, matches: GENERAL_FORMAT, describe: 'General' }, 'Order numbers use General')],
  shortcut: { mac: '⌃⇧~', windows: 'Ctrl+Shift+~' },
});

export const numberFormats: Drill = {
  id: 'number-formats',
  title: 'Number formats',
  blurb: 'Percentages, dates, times and separators, mostly from the keyboard.',
  items: [showPercent, showDates, showTimes, addThousands, redNegatives, backToGeneral],
  parSeconds: 60,
};

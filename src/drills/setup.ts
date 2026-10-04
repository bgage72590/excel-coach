import { REGIONS } from '../engine/data';
import { round, type Rng } from '../engine/rng';
import type { Grid } from '../engine/types';
import { dataBlock } from '../exercises/common';
import { CURRENCY_FORMAT, drillItem, headers, lastRow, sheetCheck } from './common';
import type { Drill } from './types';

// ---------- data: a raw order export ----------

/** Customer names long enough that column B is cut off at its default width. */
const CUSTOMERS = [
  'Pinecrest Facilities Management',
  'Riverbend Unified School District',
  'Northgate Hardware and Supply',
  'Lakeside Builders Cooperative',
  'Summit Property Group Holdings',
  'Harbor Logistics International',
  'Maple Street Community Market',
  'Cedar Ridge Family Medical Clinic',
  'Granite Peak Construction',
  'Bluewater Marine Services',
] as const;

/**
 * Every rep has at least one of these. At 32+ characters they autofit to roughly 130 points or
 * more even in a narrow font, against a default column width of 48 to 65 points.
 */
const LONG_CUSTOMERS = CUSTOMERS.filter((c) => c.length >= 32);

/** Column B must end up at least this wide (points): well above any default, well below the autofit width. */
export const CUSTOMER_MIN_WIDTH = 100;

export interface OrderRow {
  order: string;
  customer: string;
  region: string;
  units: number;
  cost: number;
  amount: number;
}

export interface OrdersData {
  rows: OrderRow[];
}

/** Unformatted, like a fresh export: every number shows as General. */
export const ORDER_COLS = headers('Order', 'Customer', 'Region', 'Units', 'Cost', 'Amount');

function orders(rng: Rng): OrdersData {
  const count = rng.int(8, 12);
  const first = rng.int(40, 90) * 100;
  const long = rng.int(0, count - 1);
  const rows = Array.from({ length: count }, (_, i) => {
    const units = rng.int(2, 60);
    const cost = round(units * rng.float(4, 40), 2);
    return {
      order: `SO-${first + i}`,
      customer: rng.pick(i === long ? LONG_CUSTOMERS : CUSTOMERS),
      region: rng.pick(REGIONS),
      units,
      cost,
      amount: round(cost * rng.float(1.2, 1.6, 3), 2),
    };
  });
  return { rows };
}

export const orderGrid = (rows: OrderRow[]): Grid => rows.map((r) => [r.order, r.customer, r.region, r.units, r.cost, r.amount]);

const ordersBlocks = (d: OrdersData) => [dataBlock('Orders', 'A1', ORDER_COLS, orderGrid(d.rows), false)];

// ---------- items ----------

const freezeTopRow = drillItem<OrdersData>({
  id: 'drill-freeze-top-row',
  title: 'Freeze the top row',
  replaces: 'Scrolling back up to see which column is which',
  task: () => 'Freeze the top row so the headers stay in view when you scroll.',
  concept: {
    summary: 'Freeze Panes keeps rows and columns in view while the rest of the sheet scrolls. Freeze Top Row does it for the header row in one step.',
    syntax: 'View › Freeze Panes › Freeze Top Row',
  },
  hints: ['Open the View tab. Nothing needs to be selected first.', 'Select Freeze Panes › Freeze Top Row.'],
  solution: () => 'View › Freeze Panes › Freeze Top Row',
  make: orders,
  blocks: ordersBlocks,
  inspections: () => [sheetCheck({ kind: 'freeze', rows: 1, cols: 0 }, 'Row 1 is frozen, and nothing else')],
  shortcut: { mac: 'View › Freeze Panes › Freeze Top Row', windows: 'Alt, W, F, R' },
});

const turnOnFilters = drillItem<OrdersData>({
  id: 'drill-filter-on',
  title: 'Turn on filters',
  replaces: 'Copying rows to a new sheet to look at one region',
  task: () => 'Add filter buttons to the header row.',
  concept: {
    summary: 'Filter buttons in the header row sort the data or narrow it to the rows you need, and nothing else on the sheet changes.',
    syntax: 'Data › Filter',
  },
  hints: ['Click any cell in the data.', 'Select Data › Filter. An arrow appears in each header cell.'],
  solution: () => 'Click a cell in the data › Data › Filter',
  make: orders,
  blocks: ordersBlocks,
  inspections: () => [sheetCheck({ kind: 'filter', on: true }, 'Filter buttons are on')],
  shortcut: { mac: '⌘⇧F', windows: 'Ctrl+Shift+L' },
});

/**
 * The Table drill's name. Table names are workbook-wide, so it's one a learner's own workbook is
 * unlikely to hold already: Excel would refuse the rename, and the check would find their Table.
 */
export const TABLE_NAME = 'OrderExport';

const convertToTable = drillItem<OrdersData>({
  id: 'drill-convert-table',
  title: 'Convert to a Table',
  replaces: 'Reformatting and re-filtering a range every time rows are added',
  task: () => `Turn the data into a Table named \`${TABLE_NAME}\`.`,
  concept: {
    summary: 'A Table keeps its formatting, filter buttons and formulas in step as rows are added. A clear name makes formulas that use it readable.',
    syntax: '{tableKey}, then Table Name',
  },
  hints: [
    'Click any cell in the data and press {tableKey}. Leave “My table has headers” selected.',
    `On the Table tab (Table Design on Windows), type ${TABLE_NAME} in the Table Name box and press {enter}.`,
  ],
  solution: () => `{tableKey} › OK › Table Name: ${TABLE_NAME}`,
  make: orders,
  blocks: ordersBlocks,
  inspections: (d) => [{ kind: 'tableExists', table: TABLE_NAME, at: 'A1', rows: lastRow(d.rows), label: `A Table named ${TABLE_NAME} holds every row` }],
  shortcut: { mac: '⌘T, then Table › Table Name', windows: 'Ctrl+T, then Table Design › Table Name' },
});

const autofitColumn = drillItem<OrdersData>({
  id: 'drill-autofit-column',
  title: 'AutoFit a column',
  replaces: 'Dragging a column edge back and forth until the text fits',
  task: () => 'Widen column `B` to fit the longest customer name.',
  concept: {
    summary: 'AutoFit sizes a column to its longest entry, so nothing is cut off and no space is wasted.',
    syntax: 'Double-click the right edge of the column heading',
  },
  hints: [
    'Point at the line between the B and C column headings until the pointer becomes a double arrow.',
    'Double-click it. Or select column B, then Home › Format › AutoFit Column Width.',
  ],
  solution: () => 'Double-click the border between the B and C column headings',
  make: orders,
  blocks: ordersBlocks,
  inspections: (d) => [sheetCheck({ kind: 'minWidth', range: `B1:B${lastRow(d.rows)}`, points: CUSTOMER_MIN_WIDTH }, 'Column B fits the customer names')],
  // AutoFit Column Width fits the selected cells, not the whole column, so select the column first.
  shortcut: { mac: 'Double-click the column heading’s right edge', windows: 'Ctrl+Space in column B, then Alt, H, O, I' },
});

const hideColumn = drillItem<OrdersData>({
  id: 'drill-hide-column',
  title: 'Hide a column',
  replaces: 'Deleting a column for a customer copy, then rebuilding it',
  task: () => 'Hide the Cost column, `E`, so it doesn’t show when you share the sheet.',
  concept: {
    summary: 'A hidden column keeps its data, and formulas that use it still work. It’s out of view and off printouts until you unhide it.',
    syntax: 'Right-click the column heading › Hide',
  },
  hints: ['Right-click the E column heading.', 'Choose Hide. To bring it back later, select columns D to F, right-click and choose Unhide.'],
  solution: () => 'Right-click the E column heading › Hide',
  make: orders,
  blocks: ordersBlocks,
  inspections: () => [sheetCheck({ kind: 'hidden', columns: 'E', hidden: true }, 'Column E is hidden')],
  shortcut: { mac: '⌃0', windows: 'Ctrl+0' },
});

const currencyFormat = drillItem<OrdersData>({
  id: 'drill-currency-format',
  title: 'Format as currency',
  replaces: 'Typing dollar signs into a report by hand',
  task: (d) => `Format the amounts in \`F2:F${lastRow(d.rows)}\` as currency with two decimals.`,
  concept: {
    summary: 'A number format changes how a value looks, not the value. Currency adds the symbol, thousands separators and two decimals.',
    syntax: 'Home › Number Format › Currency',
  },
  hints: ['Select the amounts, or click the F column heading.', 'Choose Currency in the Number Format box on the Home tab.'],
  solution: (d) => `Select F2:F${lastRow(d.rows)} › Home › Number Format › Currency`,
  make: orders,
  blocks: ordersBlocks,
  inspections: (d) => [
    sheetCheck({ kind: 'numberFormat', range: `F2:F${lastRow(d.rows)}`, matches: CURRENCY_FORMAT, describe: 'currency with 2 decimals' }, 'Amounts show as currency'),
  ],
  shortcut: { mac: '⌃⇧$', windows: 'Ctrl+Shift+$' },
});

const boldHeader = drillItem<OrdersData>({
  id: 'drill-bold-header',
  title: 'Bold the header row',
  replaces: 'Squinting to tell the labels from the data',
  task: () => 'Make the header row, `A1:F1`, bold.',
  concept: {
    summary: 'A bold header row sets the labels apart from the data at a glance.',
    syntax: 'Home › Bold',
  },
  hints: ['Select `A1:F1`, or click the row 1 heading.', 'Select Bold on the Home tab.'],
  solution: () => 'Select A1:F1 › Home › Bold',
  make: orders,
  blocks: ordersBlocks,
  inspections: (d) => [
    sheetCheck({ kind: 'bold', range: 'A1:F1', bold: true }, 'The header row is bold'),
    sheetCheck({ kind: 'bold', range: `A2:F${lastRow(d.rows)}`, bold: false }, 'The data rows aren’t bold'),
  ],
  shortcut: { mac: '⌘B', windows: 'Ctrl+B' },
});

export const sheetSetup: Drill = {
  id: 'sheet-setup',
  title: 'Sheet setup',
  blurb: 'Turn a raw export into a sheet people can read.',
  items: [freezeTopRow, turnOnFilters, convertToTable, autofitColumn, hideColumn, currencyFormat, boldHeader],
  parSeconds: 75,
};

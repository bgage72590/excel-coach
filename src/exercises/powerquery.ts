import { FIRST_NAMES, GL_ACCOUNTS, ITEMS, LAST_NAMES, VENDORS, WAREHOUSES, serial } from '../engine/data';
import { round, type Rng } from '../engine/rng';
import type { Block, ColumnSpec, Exercise, Grid } from '../engine/types';
import { FMT, cells, dataBlock, defineExercise } from './common';

/** Shown in every Power Query exercise: the output never lands on the practice sheet. */
const FINDS_IT =
  '{macPqNote} The output loads to a new sheet; the coach finds it anywhere in the workbook. Before a second attempt, delete the last attempt’s output sheet and its query (Data › Queries & Connections) so the coach checks the new one.';

/** Where a query starts; expands per platform (Excel for Mac has no From Table/Range). */
const fromTable = (table: string) => `{fromTable:${table}}`;

/**
 * Loading a helper query that Merge or Append will pick by name. Excel for Mac has no Load To
 * dialog, so every query loads to its own new sheet.
 */
const helperLoad = (table: string) =>
  `{nameQuery:${table}}then Home › Close & Load. In Excel for Mac its copy lands on a new sheet and does no harm (on Windows, Home › Close & Load To… › Only Create Connection skips the copy).`;

// ---------- Power Query text semantics ----------

/** Text.Trim: removes whitespace at the start and end only. Unlike Excel's TRIM, inner runs stay. */
export function pqTrim(s: string): string {
  return s.replace(/^\s+|\s+$/g, '');
}

/** Text.Upper for plain ASCII text. */
export function pqUpper(s: string): string {
  return s.toUpperCase();
}

/**
 * Text.Proper for plain ASCII words: the first letter of each word upper case, the rest lower
 * ("the QUICK BrOwn fOx" → "The Quick Brown Fox", the documented example). The data only uses
 * letters and single spaces, so digits and punctuation never come into play.
 */
export function pqProper(s: string): string {
  let out = '';
  let prevLetter = false;
  for (const ch of s) {
    const letter = /[A-Za-z]/.test(ch);
    out += letter ? (prevLetter ? ch.toLowerCase() : ch.toUpperCase()) : ch;
    prevLetter = letter;
  }
  return out;
}

/** Remove Duplicates across every column: keeps the first copy of each identical row. */
export function distinctRows(rows: Grid): Grid {
  const seen = new Set<string>();
  return rows.filter((r) => {
    const key = JSON.stringify(r);
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

// ---------- Clean a messy system export ----------

interface ExportLine {
  invoice: string;
  vendor: string;
  warehouse: string;
  amount: number;
}

interface CleanData {
  /** Raw export lines in sheet order; null is a fully blank row. */
  rows: (ExportLine | null)[];
}

const EXPORT_COLS: ColumnSpec[] = [{ header: 'Invoice' }, { header: 'Vendor' }, { header: 'Warehouse' }, { header: 'Amount', format: FMT.currency }];

const lower = (s: string) => s.toLowerCase();
const upper = (s: string) => s.toUpperCase();
const asIs = (s: string) => s;

/** Wraps text in stray leading and trailing spaces, never doubling a space inside it. */
function pad(rng: Rng, s: string): string {
  const lead = rng.chance(0.35) ? ' '.repeat(rng.int(1, 2)) : '';
  const trail = rng.chance(0.4) ? ' '.repeat(rng.int(1, 3)) : '';
  return `${lead}${s}${trail}`;
}

function exportLines(rng: Rng): (ExportLine | null)[] {
  const count = rng.int(32, 40);
  let no = rng.int(20100, 27000);
  const base: ExportLine[] = Array.from({ length: count }, () => {
    no += rng.int(1, 9);
    return {
      invoice: pad(rng, (rng.chance(0.35) ? lower : upper)(`INV-${no}`)),
      vendor: pad(rng, rng.pick([lower, upper, asIs, asIs])(rng.pick(VENDORS))),
      warehouse: pad(rng, rng.pick([lower, upper, asIs])(rng.pick(WAREHOUSES))),
      amount: rng.float(60, 8400, 2),
    };
  });

  // Exact duplicates: a line exported twice, spaces and capitals included.
  const rows: (ExportLine | null)[] = base.slice();
  for (const src of rng.sample(base, rng.int(3, 5))) {
    const at = rows.indexOf(src);
    rows.splice(rng.int(at + 1, rows.length), 0, { ...src });
  }
  // Fully blank rows, never first or last.
  const blanks = rng.int(3, 4);
  for (let i = 0; i < blanks; i++) rows.splice(rng.int(2, rows.length - 2), 0, null);
  return rows;
}

const exportGrid = (rows: (ExportLine | null)[]): Grid =>
  rows.map((r) => (r ? [r.invoice, r.vendor, r.warehouse, r.amount] : ['', '', '', '']));

/** One raw line after the intended steps: Trim everywhere, UPPERCASE the invoice, Capitalize Each Word on the names. */
export function cleanLine(r: ExportLine): (string | number)[] {
  return [pqUpper(pqTrim(r.invoice)), pqProper(pqTrim(r.vendor)), pqProper(pqTrim(r.warehouse)), r.amount];
}

function firstLine(d: CleanData): ExportLine {
  return d.rows.find((r): r is ExportLine => r !== null)!;
}

export const pqCleanExport = defineExercise<CleanData>({
  id: 'pq-clean-export',
  module: 'powerquery',
  title: 'Clean a messy system export',
  replaces: 'Helper columns of TRIM and PROPER rebuilt for every new export',
  minutes: 8,
  task: (d) => {
    const [invoice, vendor] = cleanLine(firstLine(d));
    return (
      'The `Export` Table is a raw AP export with stray spaces, random capitals, blank rows and lines exported twice. ' +
      'Use Power Query to load a clean copy with the same four columns, `Invoice`, `Vendor`, `Warehouse` and `Amount`: ' +
      `no spaces before or after any text, \`Invoice\` in capitals (\`${invoice}\`), \`Vendor\` and \`Warehouse\` with each word capitalized (\`${vendor}\`), ` +
      'no blank rows and no duplicate rows. Row order doesn’t matter.'
    );
  },
  concept: {
    summary:
      'Power Query records each cleanup step you click and replays the whole list on demand. Next month you paste the new export into the Table and choose Data › Refresh All instead of cleaning it again.',
    syntax:
      `${fromTable('Export')}, then Transform › Format › Trim, UPPERCASE or Capitalize Each Word, then Home › Remove Rows › Remove Blank Rows and Remove Duplicates, then Home › Close & Load`,
    example: 'Each click adds a step to Applied Steps. Trim on Vendor becomes Table.TransformColumns(#"Changed Type", {{"Vendor", Text.Trim, type text}}).',
    tip: `Power Query’s Trim removes spaces only at the start and end of the text; Excel’s TRIM also squeezes doubled spaces inside it. ${FINDS_IT}`,
  },
  hints: [
    `Choose ${fromTable('Export')}. The Power Query editor opens with the raw rows.`,
    'Select the Invoice column, then Transform › Format › Trim and Transform › Format › UPPERCASE. Select Vendor and Warehouse together (⌘-click, or Ctrl-click on Windows) and choose Trim, then Capitalize Each Word.',
    'Choose Home › Remove Rows › Remove Blank Rows. Then select all four columns (click Invoice, Shift-click Amount) and Home › Remove Rows › Remove Duplicates. Finish with Home › Close & Load.',
  ],
  solution: () =>
    `${fromTable('Export')} · Invoice: Trim, UPPERCASE · Vendor and Warehouse: Trim, Capitalize Each Word · Home › Remove Rows › Remove Blank Rows · select all columns › Remove Rows › Remove Duplicates · Home › Close & Load`,
  make: (rng) => ({ rows: exportLines(rng) }),
  layout: (d) => ({
    blocks: [dataBlock('Export', 'A1', EXPORT_COLS, exportGrid(d.rows))],
    answer: { kind: 'query', columns: EXPORT_COLS.map((c) => c.header), order: 'any' },
  }),
  expected: (d) => distinctRows(d.rows.filter((r): r is ExportLine => r !== null).map(cleanLine)),
  inputs: () => [],
  variants: [],
});

// ---------- Unpivot a monthly grid ----------

const HALVES = [
  ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'],
  ['Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'],
] as const;

interface BudgetLine {
  account: string;
  amounts: number[];
}

interface UnpivotData {
  months: string[];
  rows: BudgetLine[];
}

const WHOLE_DOLLARS = '$#,##0';

function budgetLines(rng: Rng, months: number): BudgetLine[] {
  const accounts = rng.sample(GL_ACCOUNTS, 10).sort((a, b) => a.no - b.no);
  return accounts.map((a) => {
    const [lo, hi] = a.type === 'Revenue' ? [1500, 6000] : a.type === 'COGS' ? [200, 2500] : [10, 450];
    const base = rng.int(lo, hi) * 100;
    // Whole hundreds of dollars, never blank: Unpivot drops null cells.
    return { account: a.name, amounts: Array.from({ length: months }, () => Math.max(100, round((base * rng.float(0.85, 1.15, 3)) / 100, 0) * 100)) };
  });
}

const budgetCols = (d: UnpivotData): ColumnSpec[] => [{ header: 'Account' }, ...d.months.map((m) => ({ header: m, format: WHOLE_DOLLARS }))];
const budgetGrid = (d: UnpivotData): Grid => d.rows.map((r) => [r.account, ...r.amounts]);

export const pqUnpivot = defineExercise<UnpivotData>({
  id: 'pq-unpivot',
  module: 'powerquery',
  title: 'Unpivot a monthly grid',
  replaces: 'Stacking month columns by hand so a PivotTable can read them',
  minutes: 5,
  task: (d) =>
    'The `Budget` Table has one row per GL account and one column per month. Use Power Query to reshape it into a long list with one row per account per month and three columns: ' +
    `\`Account\`, \`Month\` (the month’s column header, such as \`${d.months[0]}\`) and \`Budget\` (the amount). ` +
    'Unpivot names the new columns Attribute and Value, so rename them to `Month` and `Budget`. Row order doesn’t matter.',
  concept: {
    summary:
      'A grid with months across the top reads well but is hard to summarize. Unpivot turns those month columns into rows, one per account per month, which is the shape PivotTables, SUMIFS and charts want.',
    syntax: 'Select the Account column › Transform › Unpivot Columns › Unpivot Other Columns, then double-click each new header to rename it',
    example: 'The row “Rent | 12,000 | 12,000 | 12,500 | …” becomes one row per month: Rent, Jul, 12,000 · Rent, Aug, 12,000 · Rent, Sep, 12,500 · …',
    tip: `Unpivot Other Columns keeps Account fixed and unpivots everything else, so a new month column is picked up on the next refresh. ${FINDS_IT}`,
  },
  hints: [
    `Choose ${fromTable('Budget')}.`,
    'Right-click the Account header and choose Unpivot Other Columns. (Selecting the month columns and choosing Transform › Unpivot Columns gives the same result.)',
    'Double-click Attribute and rename it Month, double-click Value and rename it Budget, then Home › Close & Load.',
  ],
  solution: () => `${fromTable('Budget')} · right-click Account › Unpivot Other Columns · rename Attribute to Month and Value to Budget · Home › Close & Load`,
  make: (rng) => {
    const months = [...rng.pick(HALVES)];
    return { months, rows: budgetLines(rng, months.length) };
  },
  layout: (d) => ({
    blocks: [dataBlock('Budget', 'A1', budgetCols(d), budgetGrid(d))],
    answer: { kind: 'query', columns: ['Account', 'Month', 'Budget'], order: 'any' },
  }),
  expected: (d) => d.rows.flatMap((r) => d.months.map((m, i) => [r.account, m, r.amounts[i]])),
  inputs: () => [],
  variants: [],
});

// ---------- Merge two tables ----------

interface ItemRef {
  sku: string;
  item: string;
  category: string;
  cost: number;
}

interface OrderLine {
  order: string;
  sku: string;
  qty: number;
}

interface MergeData {
  items: ItemRef[];
  orders: OrderLine[];
}

const ORDER_COLS: ColumnSpec[] = [{ header: 'Order' }, { header: 'SKU' }, { header: 'Qty', format: FMT.int }];
const ITEM_COLS: ColumnSpec[] = [{ header: 'SKU' }, { header: 'Item' }, { header: 'Category' }, { header: 'Unit cost', format: FMT.currency }];

function itemRefs(rng: Rng): ItemRef[] {
  const numbers = rng.sample(Array.from({ length: 400 }, (_, i) => 1100 + i), 14).sort((a, b) => a - b);
  return rng.sample(ITEMS, 14).map((it, i) => ({ sku: `SKU-${numbers[i]}`, item: it.item, category: it.category, cost: it.cost }));
}

function orderLines(rng: Rng, items: ItemRef[]): OrderLine[] {
  const target = rng.int(36, 44);
  const out: OrderLine[] = [];
  let no = rng.int(48100, 48900);
  while (out.length < target) {
    no++;
    // One to three lines per order, each a different SKU.
    for (const it of rng.sample(items, rng.int(1, 3))) out.push({ order: `SO-${no}`, sku: it.sku, qty: rng.int(2, 240) });
  }
  return out;
}

export const pqMerge = defineExercise<MergeData>({
  id: 'pq-merge',
  module: 'powerquery',
  title: 'Merge two tables',
  replaces: 'A VLOOKUP column added by hand to every new order export',
  minutes: 7,
  task: () =>
    'The `Orders` Table lists order lines by SKU, and the `Items` Table holds each SKU’s item, category and unit cost. ' +
    'Use Power Query to load the order lines with each line’s category added from `Items`, matched on SKU. ' +
    'The output must have exactly four columns, `Order`, `SKU`, `Qty` and `Category`, with one row per order line. Row order doesn’t matter.',
  concept: {
    summary:
      'Merge Queries is Power Query’s lookup: it matches rows from two queries on a shared column, like an XLOOKUP for the whole table at once. A left outer join keeps every row of the first query and brings in the matching details from the second.',
    syntax: 'Home › Combine › Merge Queries › choose Items › click SKU in both previews › Join Kind: Left Outer (all from first, matching from second) › OK, then expand the new column',
    example: 'After the merge, a new Items column shows Table in every row. The expand button in its header picks which Items columns to bring in.',
    tip: `Untick “Use original column name as prefix” when you expand, or the new column is named Items.Category. Merging can reorder rows, so the coach ignores row order. ${FINDS_IT}`,
  },
  hints: [
    `Merge Queries can only pick from existing queries, so start with \`Items\`: choose ${fromTable('Items')}, ${helperLoad('Items')}`,
    `Back on the practice sheet, choose ${fromTable('Orders')}. Choose Home › Combine › Merge Queries, pick Items in the second list, click the SKU header in both previews, keep Join Kind on Left Outer, then OK.`,
    'Click the expand button (two arrows) on the new Items column, tick only Category, untick “Use original column name as prefix”, then OK and Home › Close & Load.',
  ],
  solution: () =>
    `Items: ${fromTable('Items')} › {nameQuery:Items}Home › Close & Load (Windows: Close & Load To › Only Create Connection) · Orders: ${fromTable('Orders')} › Home › Combine › Merge Queries with Items on SKU, Left Outer › expand Category without the prefix › Home › Close & Load`,
  make: (rng) => {
    const items = itemRefs(rng);
    return { items, orders: orderLines(rng, items) };
  },
  layout: (d) => ({
    blocks: [
      dataBlock('Orders', 'A1', ORDER_COLS, d.orders.map((o) => [o.order, o.sku, o.qty])),
      dataBlock('Items', 'E1', ITEM_COLS, d.items.map((i) => [i.sku, i.item, i.category, i.cost])),
    ],
    answer: { kind: 'query', columns: ['Order', 'SKU', 'Qty', 'Category'], order: 'any' },
  }),
  expected: (d) => d.orders.map((o) => [o.order, o.sku, o.qty, d.items.find((i) => i.sku === o.sku)!.category]),
  inputs: () => [],
  variants: [],
});

// ---------- Append monthly files ----------

interface ApLine {
  date: number;
  invoice: string;
  vendor: string;
  amount: number;
  enteredBy: string;
}

interface ApMonth {
  table: string;
  label: string;
  rows: ApLine[];
}

interface AppendData {
  months: ApMonth[];
}

const AP_COLS: ColumnSpec[] = [
  { header: 'Date', format: FMT.date },
  { header: 'Invoice' },
  { header: 'Vendor' },
  { header: 'Amount', format: FMT.currency },
  { header: 'Entered by' },
];

const AP_MONTHS = [
  { table: 'AP_Jul', label: 'July', month: 7 },
  { table: 'AP_Aug', label: 'August', month: 8 },
  { table: 'AP_Sep', label: 'September', month: 9 },
] as const;

function apMonths(rng: Rng): ApMonth[] {
  const clerks = Array.from({ length: 3 }, () => `${rng.pick(FIRST_NAMES)} ${rng.pick(LAST_NAMES)}`);
  let no = rng.int(30100, 36000);
  return AP_MONTHS.map((m) => {
    const first = serial(2026, m.month, 1);
    const last = serial(2026, m.month + 1, 0);
    const dates = Array.from({ length: rng.int(12, 20) }, () => rng.int(first, last)).sort((a, b) => a - b);
    const rows = dates.map((date) => {
      no += rng.int(1, 6);
      return { date, invoice: `INV-${no}`, vendor: rng.pick(VENDORS), amount: rng.float(85, 9800, 2), enteredBy: rng.pick(clerks) };
    });
    return { table: m.table, label: m.label, rows };
  });
}

const apGrid = (rows: ApLine[]): Grid => rows.map((r) => [r.date, r.invoice, r.vendor, r.amount, r.enteredBy]);
const apTotal = (d: AppendData) => d.months.reduce((n, m) => n + m.rows.length, 0);

/** Stacks the three monthly Tables down column A, each under a short label, one blank row apart. */
function apBlocks(d: AppendData): Block[] {
  const blocks: Block[] = [];
  let row = 1;
  for (const m of d.months) {
    blocks.push(cells(`A${row}`, [[`${m.label} · ${m.table}`]], 'label'));
    blocks.push(dataBlock(m.table, `A${row + 1}`, AP_COLS, apGrid(m.rows)));
    row += m.rows.length + 3;
  }
  return blocks;
}

export const pqAppend = defineExercise<AppendData>({
  id: 'pq-append',
  module: 'powerquery',
  title: 'Append monthly files',
  replaces: 'Pasting each month’s AP file under the last one by hand',
  minutes: 7,
  task: (d) =>
    'The `AP_Jul`, `AP_Aug` and `AP_Sep` Tables are three monthly invoice files with the same five columns. ' +
    `Use Power Query to stack them into one output with all ${apTotal(d)} rows, then remove the \`Entered by\` column so the output has exactly four columns: ` +
    '`Date`, `Invoice`, `Vendor` and `Amount`. Row order doesn’t matter.',
  concept: {
    summary:
      'Append Queries stacks queries with the same columns on top of each other, matching columns by name. Build it once, and a refresh picks up every row added to the monthly files.',
    syntax: 'Home › Combine › Append Queries › Three or more tables › add each monthly query › OK',
    example: 'Three files of 15, 18 and 12 rows become one query of 45 rows with the same columns. In M it is Table.Combine({AP_Jul, AP_Aug, AP_Sep}).',
    tip: `In Excel for Mac each monthly query also loads to a sheet; the coach looks only for the four-column result. Dates stay real dates in the output. ${FINDS_IT}`,
  },
  hints: [
    `Append Queries can only pick from existing queries. Choose ${fromTable('AP_Aug')}, ${helperLoad('AP_Aug')} Do the same for \`AP_Sep\`.`,
    `Back on the practice sheet, choose ${fromTable('AP_Jul')}. In the editor, choose Home › Combine › Append Queries › Three or more tables, then add AP_Aug and AP_Sep to the list on the right.`,
    'After the append step, right-click the Entered by header › Remove (Remove columns on Mac), then Home › Close & Load.',
  ],
  solution: () =>
    `AP_Aug and AP_Sep: ${fromTable('AP_Aug')} (and the same for AP_Sep) › {nameQuery:AP_Aug}Home › Close & Load (Windows: Close & Load To › Only Create Connection) · AP_Jul: ${fromTable('AP_Jul')} › Home › Combine › Append Queries › Three or more tables (AP_Aug, AP_Sep) › remove Entered by › Home › Close & Load`,
  make: (rng) => ({ months: apMonths(rng) }),
  layout: (d) => ({
    blocks: apBlocks(d),
    answer: { kind: 'query', columns: ['Date', 'Invoice', 'Vendor', 'Amount'], order: 'any' },
  }),
  // Power Query loads the dates back as dates, so Excel reports them as serial numbers.
  expected: (d) => d.months.flatMap((m) => m.rows.map((r) => [r.date, r.invoice, r.vendor, r.amount])),
  inputs: () => [],
  variants: [],
});

// ---------- Group and sum ----------

type GroupField = 'Warehouse' | 'Category';

interface Shipment {
  id: string;
  warehouse: string;
  category: string;
  item: string;
  units: number;
}

interface GroupData {
  by: GroupField;
  rows: Shipment[];
}

const SHIP_COLS: ColumnSpec[] = [{ header: 'Ship ID' }, { header: 'Warehouse' }, { header: 'Category' }, { header: 'Item' }, { header: 'Units', format: FMT.int }];

function shipments(rng: Rng): Shipment[] {
  let no = rng.int(52000, 58000);
  return Array.from({ length: rng.int(70, 90) }, () => {
    const it = rng.pick(ITEMS);
    no += rng.int(1, 4);
    return { id: `SH-${no}`, warehouse: rng.pick(WAREHOUSES), category: it.category, item: it.item, units: rng.int(5, 400) };
  });
}

const groupKey = (s: Shipment, by: GroupField) => (by === 'Warehouse' ? s.warehouse : s.category);

/** Table.Group with a Sum: one row per value, in order of first appearance. */
export function groupTotals(d: GroupData): Grid {
  const totals = new Map<string, number>();
  for (const s of d.rows) totals.set(groupKey(s, d.by), (totals.get(groupKey(s, d.by)) ?? 0) + s.units);
  return [...totals].map(([k, v]) => [k, v]);
}

export const pqGroup = defineExercise<GroupData>({
  id: 'pq-group',
  module: 'powerquery',
  title: 'Group and sum',
  replaces: 'A unique list plus a SUMIF beside every name, rebuilt weekly',
  minutes: 5,
  task: (d) => {
    const noun = d.by.toLowerCase();
    return (
      `Use Power Query to total the units shipped for each ${noun} in the \`Shipments\` Table. ` +
      `Load an output with exactly two columns: \`${d.by}\`, with one row per ${noun}, and \`Total units\`, the sum of Units for that ${noun}. Row order doesn’t matter.`
    );
  },
  concept: {
    summary:
      'Group By collapses a table to one row per value in the column you choose and calculates a summary for each group, such as a sum or a count. It is the Power Query version of a SUMIFS summary, and it refreshes with the data.',
    syntax: 'Transform › Group By › Basic › group by the column › New column name, Operation: Sum, Column: Units › OK',
    example: 'Table.Group(Source, {"Warehouse"}, {{"Total units", each List.Sum([Units]), type number}}) turns 80 shipment lines into one row per warehouse.',
    tip: `Group By keeps only the grouping column and the new total; every other column drops out. ${FINDS_IT}`,
  },
  hints: [
    `Choose ${fromTable('Shipments')}.`,
    'Select the column you are totaling by, then choose Transform › Group By. (On Windows it is also on the Home tab.)',
    'In the dialog, keep Basic, set New column name to Total units, Operation to Sum and Column to Units, then OK and Home › Close & Load.',
  ],
  solution: (d) => `${fromTable('Shipments')} · Transform › Group By ${d.by}, New column name: Total units, Operation: Sum, Column: Units · Home › Close & Load`,
  make: (rng) => {
    const by: GroupField = rng.pick(['Warehouse', 'Category'] as const);
    return { by, rows: shipments(rng) };
  },
  layout: (d) => ({
    blocks: [dataBlock('Shipments', 'A1', SHIP_COLS, d.rows.map((s) => [s.id, s.warehouse, s.category, s.item, s.units]))],
    answer: { kind: 'query', columns: [d.by, 'Total units'], order: 'any' },
  }),
  expected: (d) => groupTotals(d),
  inputs: () => [],
  variants: [],
});

export const POWER_QUERY: Exercise<any>[] = [pqCleanExport, pqUnpivot, pqMerge, pqAppend, pqGroup];

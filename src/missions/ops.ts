import { CATEGORIES, ITEMS, WAREHOUSES, excelTextCompare, serial, sum, type ItemDef } from '../engine/data';
import { round, type Rng } from '../engine/rng';
import type { ColumnSpec, Grid } from '../engine/types';
import { FMT, cells, column, dataBlock, rangeWrite, tableWrite } from '../exercises/common';
import { defineMission, type Mission } from './types';

// =====================================================================================
// Mission 1: Month-end inventory report
// =====================================================================================

/** Extra stock items so the snapshot can hold 40+ distinct SKUs, one row each. */
const MORE_ITEMS: readonly ItemDef[] = [
  { item: 'Shipping box 12 in', category: 'Packaging', cost: 1.15 },
  { item: 'Shipping box 18 in', category: 'Packaging', cost: 1.85 },
  { item: 'Void fill paper', category: 'Packaging', cost: 38.0 },
  { item: 'Pallet label roll', category: 'Packaging', cost: 14.2 },
  { item: 'Rack beam 96 in', category: 'Hardware', cost: 48.5 },
  { item: 'Wire deck panel', category: 'Hardware', cost: 36.0 },
  { item: 'Anchor bolt kit', category: 'Hardware', cost: 22.4 },
  { item: 'Dock bumper', category: 'Hardware', cost: 74.0 },
  { item: 'Power strip', category: 'Electrical', cost: 18.9 },
  { item: 'Motion sensor', category: 'Electrical', cost: 31.5 },
  { item: 'Charging cable', category: 'Electrical', cost: 7.25 },
  { item: 'Scanner holster', category: 'Electrical', cost: 12.5 },
  { item: 'Safety glasses', category: 'Safety', cost: 3.9 },
  { item: 'Cut-resistant gloves', category: 'Safety', cost: 9.75 },
  { item: 'First aid kit', category: 'Safety', cost: 42.0 },
  { item: 'Fire extinguisher', category: 'Safety', cost: 58.0 },
  { item: 'Floor cleaner', category: 'Janitorial', cost: 16.4 },
  { item: 'Squeegee', category: 'Janitorial', cost: 11.8 },
  { item: 'Hand soap refill', category: 'Janitorial', cost: 8.6 },
  { item: 'Absorbent pads', category: 'Janitorial', cost: 27.0 },
];

const STOCK_ITEMS: readonly ItemDef[] = [...ITEMS, ...MORE_ITEMS];
const SNAPSHOT_ROWS = 42;
const ADDED_ROWS = 4;

export interface StockRow {
  sku: string;
  item: string;
  category: string;
  warehouse: string;
  onHand: number;
  cost: number;
  value: number;
  reorder: number;
  /** Typical full-shelf quantity; drives realistic on-hand levels. Not shown on the sheet. */
  cap: number;
}

export interface InventoryData {
  rows: StockRow[];
  /** Column order of the warehouse headers in K1:N1. */
  warehouses: string[];
  /** Row order of the category labels in J2:J6 and J11:J15. */
  categories: string[];
}

const INVENTORY_COLS: ColumnSpec[] = [
  { header: 'SKU' },
  { header: 'Item' },
  { header: 'Category' },
  { header: 'Warehouse' },
  { header: 'On hand', format: FMT.int },
  { header: 'Unit cost', format: FMT.currency },
  { header: 'Value', format: FMT.currency },
  { header: 'Reorder point', format: FMT.int },
];

const inventoryGrid = (rows: StockRow[]): Grid =>
  rows.map((r) => [r.sku, r.item, r.category, r.warehouse, r.onHand, r.cost, r.value, r.reorder]);

/** Cheap items are stocked by the thousand, expensive ones by the dozen. */
function capacity(rng: Rng, cost: number): number {
  if (cost < 1) return rng.int(1500, 6000);
  if (cost < 10) return rng.int(150, 800);
  if (cost < 40) return rng.int(40, 220);
  return rng.int(12, 70);
}

function reorderFor(rng: Rng, cap: number): number {
  const rp = Math.max(2, Math.round(cap * rng.float(0.15, 0.35, 3)));
  return rp >= 50 ? Math.round(rp / 5) * 5 : rp;
}

function onHandFor(rng: Rng, cap: number, reorder: number): number {
  return rng.chance(0.22) ? rng.int(0, reorder) : rng.int(reorder + 1, Math.max(reorder + 1, cap));
}

const withValue = (r: StockRow): StockRow => ({ ...r, value: round(r.onHand * r.cost, 2) });
export const isLow = (r: StockRow) => r.onHand <= r.reorder;

function stockRow(rng: Rng, it: ItemDef, sku: string): StockRow {
  const cap = capacity(rng, it.cost);
  const reorder = reorderFor(rng, cap);
  return withValue({
    sku,
    item: it.item,
    category: it.category,
    warehouse: rng.pick(WAREHOUSES),
    onHand: onHandFor(rng, cap, reorder),
    cost: it.cost,
    value: 0,
    reorder,
    cap,
  });
}

/**
 * Keeps the snapshot realistic and the checks meaningful: at least one SKU sits exactly on its
 * reorder point (so "at or below" matters), one is stocked out, and the reorder list is never short.
 */
export function ensureStockEdges(rows: StockRow[], rng: Rng): StockRow[] {
  const out = rows.map((r) => ({ ...r }));
  if (!out.some((r) => r.onHand === r.reorder)) {
    const r = out[rng.int(0, out.length - 1)];
    r.onHand = r.reorder;
  }
  if (!out.some((r) => r.onHand === 0)) {
    const candidates = out.filter((r) => r.onHand !== r.reorder);
    candidates[rng.int(0, candidates.length - 1)].onHand = 0;
  }
  while (out.filter(isLow).length < 6) {
    const high = out.filter((r) => !isLow(r));
    const r = high[rng.int(0, high.length - 1)];
    r.onHand = rng.int(1, r.reorder);
  }
  return out.map(withValue);
}

/** WMS exports come grouped by warehouse, then SKU. */
const byWarehouseThenSku = (a: StockRow, b: StockRow) => excelTextCompare(a.warehouse, b.warehouse) || excelTextCompare(a.sku, b.sku);

function newSkus(rng: Rng, taken: Set<string>, count: number): string[] {
  const out: string[] = [];
  while (out.length < count) {
    const sku = `SKU-${rng.int(2000, 2999)}`;
    if (!taken.has(sku)) {
      taken.add(sku);
      out.push(sku);
    }
  }
  return out;
}

function addStockLines(d: InventoryData, rng: Rng): InventoryData {
  const used = new Set(d.rows.map((r) => r.item));
  const spare = rng.sample(
    STOCK_ITEMS.filter((it) => !used.has(it.item)),
    ADDED_ROWS,
  );
  const skus = newSkus(rng, new Set(d.rows.map((r) => r.sku)), spare.length);
  const added = spare.map((it, i) => stockRow(rng, it, skus[i]));
  // One of the new lines is already low, so the reorder list has to pick it up.
  added[0] = withValue({ ...added[0], onHand: rng.int(0, added[0].reorder) });
  return { ...d, rows: [...d.rows, ...added] };
}

/** A shuffle that is guaranteed to change the order (items are compared by `key`). */
function reshuffle<T>(rng: Rng, items: readonly T[], key: (x: T) => string = String): T[] {
  const order = (xs: readonly T[]) => xs.map(key).join('|');
  let next = rng.shuffle(items);
  while (order(next) === order(items)) next = rng.shuffle(items);
  return next;
}

const reorderLabels = (rng: Rng, labels: readonly string[]) => reshuffle(rng, labels);

const stockLevelsChange = {
  label: 'stock levels change',
  apply: (d: InventoryData, rng: Rng): InventoryData => ({
    ...d,
    rows: ensureStockEdges(
      d.rows.map((r) => ({ ...r, onHand: onHandFor(rng, r.cap, r.reorder) })),
      rng,
    ),
  }),
};

const reorderPointsChange = {
  label: 'reorder points change',
  apply: (d: InventoryData, rng: Rng): InventoryData => ({
    ...d,
    rows: ensureStockEdges(
      d.rows.map((r) => ({ ...r, reorder: reorderFor(rng, r.cap) })),
      rng,
    ),
  }),
};

const sumValue = (rows: StockRow[]) => sum(rows.map((r) => r.value));

export const monthEndInventory = defineMission<InventoryData>({
  id: 'mission-month-end-inventory',
  title: 'Month-end inventory report',
  role: 'ops',
  summary: 'The month-end stock valuation, reorder counts and purchasing list from one snapshot',
  minutes: 20,
  skills: ['sumifs-warehouse', 'sumifs-grid', 'filter-late', 'unique-vendors'],
  brief: (d) => ({
    from: 'Dana Ruiz, Operations manager',
    subject: 'Month-end inventory pack for Friday',
    body: [
      'Hi,',
      '',
      `Finance needs the month-end inventory pack by Friday. The September 30 snapshot is in the \`Inventory\` Table: ${d.rows.length} SKUs across our four warehouses. Build everything off the Table so I can paste October’s snapshot over it and have the whole pack update.`,
      '',
      '1. Stock value by category and warehouse in `K2:N6`.',
      '2. How many SKUs each warehouse has at or below reorder point, in `K8:N8`.',
      '3. Each category’s share of total stock value in `K11:K15`.',
      '4. A reorder list for purchasing starting at `J18`, sorted by SKU.',
      '',
      'Thanks,',
      'Dana',
    ].join('\n'),
  }),
  make: (rng) => {
    const picked = rng.sample(STOCK_ITEMS, SNAPSHOT_ROWS);
    const skus = newSkus(rng, new Set(), picked.length);
    const rows = ensureStockEdges(
      picked.map((it, i) => stockRow(rng, it, skus[i])),
      rng,
    ).sort(byWarehouseThenSku);
    return { rows, warehouses: rng.shuffle(WAREHOUSES), categories: rng.shuffle(CATEGORIES) };
  },
  blocks: (d) => [
    dataBlock('Inventory', 'A1', INVENTORY_COLS, inventoryGrid(d.rows)),
    cells('J1', [['Category']], 'header'),
    cells('K1', [d.warehouses], 'header'),
    cells('J2', column(d.categories), 'label'),
    cells('J8', [['SKUs to reorder']], 'label'),
    cells('J10', [['Category', 'Share of value']], 'header'),
    cells('J11', column(d.categories), 'label'),
    cells('J17', [['SKU', 'Item', 'Category', 'Warehouse']], 'header'),
  ],
  inputs: (d) => [
    tableWrite('Inventory', INVENTORY_COLS, inventoryGrid(d.rows)),
    rangeWrite('K1:N1', [d.warehouses]),
    rangeWrite('J2:J6', column(d.categories)),
    rangeWrite('J11:J15', column(d.categories)),
  ],
  steps: [
    // ---------- 1. SUMIFS grid ----------
    {
      title: 'Stock value by category and warehouse',
      task: () =>
        'Fill `K2:N6` with the total stock Value for each category (column `J`) in each warehouse (row `1`). Write one formula in `K2` that still works when you fill it right and down.',
      hints: [
        'Use SUMIFS with two conditions: Inventory[Category] matched against the category in column J, and Inventory[Warehouse] matched against the warehouse in row 1.',
        'Lock the column on the category ($J2) and the row on the warehouse (K$1) so both stay put as you fill. Press {absKey} to cycle the $ signs.',
        '=SUMIFS(Inventory[Value], Inventory[Category], $J2, Inventory[Warehouse], K$1), then fill right to column N and down to row 6. To fill right, select the row and press {fillRight} (Home › Fill › Right) rather than dragging the fill handle, which slides Table column names along.',
      ],
      solution: () => '=SUMIFS(Inventory[Value],Inventory[Category],$J2,Inventory[Warehouse],K$1)',
      answer: () => ({ kind: 'cells', range: 'K2:N6', format: FMT.currency, consistency: 'all' }),
      expected: (d) =>
        d.categories.map((c) => d.warehouses.map((w) => sumValue(d.rows.filter((r) => r.category === c && r.warehouse === w)))),
      variants: [
        stockLevelsChange,
        {
          label: 'the warehouses in row 1 are reordered',
          explain: 'Lock the row with K$1 so every cell reads its warehouse from row 1.',
          apply: (d, rng) => ({ ...d, warehouses: reorderLabels(rng, d.warehouses) }),
        },
        {
          label: 'the categories in column J are reordered',
          explain: 'Lock the column with $J2 so every cell reads its category from column J.',
          apply: (d, rng) => ({ ...d, categories: reorderLabels(rng, d.categories) }),
        },
        {
          label: 'four new stock lines are added',
          explain: 'Use whole Table columns like Inventory[Value] so new rows are counted.',
          apply: addStockLines,
        },
      ],
      rules: {
        forbidText: { values: [...WAREHOUSES, ...CATEGORIES], advice: 'Point to the category in column J and the warehouse in row 1 instead.' },
        allowNumbers: [],
      },
    },

    // ---------- 2. Count by comparing two columns ----------
    {
      title: 'SKUs to reorder in each warehouse',
      task: () =>
        'In `K8:N8`, count the SKUs in each warehouse (row `1`) that are at or below their reorder point, meaning On hand is less than or equal to Reorder point. Write one formula in `K8` and fill it right to `N8`.',
      hints: [
        'COUNTIFS can’t compare one column with another row by row. Compare them yourself: Inventory[On hand]<=Inventory[Reorder point] gives TRUE or FALSE for every row.',
        'Multiply that by a second test, (Inventory[Warehouse]=K$1), so a row becomes 1 only when both are TRUE. SUMPRODUCT adds up the 1s.',
        '=SUMPRODUCT((Inventory[Warehouse]=K$1)*(Inventory[On hand]<=Inventory[Reorder point])), then fill right to N8.',
      ],
      solution: () => '=SUMPRODUCT((Inventory[Warehouse]=K$1)*(Inventory[On hand]<=Inventory[Reorder point]))',
      answer: () => ({ kind: 'cells', range: 'K8:N8', format: FMT.int, consistency: 'all' }),
      expected: (d) => [d.warehouses.map((w) => d.rows.filter((r) => r.warehouse === w && isLow(r)).length)],
      variants: [
        stockLevelsChange,
        reorderPointsChange,
        {
          label: 'the warehouses in row 1 are reordered',
          explain: 'Compare against the warehouse in row 1 instead of typing its name.',
          apply: (d, rng) => ({ ...d, warehouses: reorderLabels(rng, d.warehouses) }),
        },
        {
          label: 'four new stock lines are added',
          explain: 'Compare whole Table columns so new rows are counted.',
          apply: addStockLines,
        },
      ],
      rules: {
        forbidText: { values: [...WAREHOUSES], advice: 'Compare against the warehouse in row 1 instead.' },
        allowNumbers: [0, 1],
      },
    },

    // ---------- 3. Share of total ----------
    {
      title: 'Share of stock value by category',
      task: () =>
        'In `K11:K15`, show each category’s share of total stock value: the category’s Value divided by the Value of the whole `Inventory` Table. Write one formula in `K11` and fill it down. The five shares add up to 100%.',
      hints: [
        'The top of the fraction is a SUMIFS on Inventory[Value] for the category in J11.',
        'The bottom is the whole inventory: SUM(Inventory[Value]).',
        '=SUMIFS(Inventory[Value], Inventory[Category], J11)/SUM(Inventory[Value]), then fill down to K15.',
      ],
      solution: () => '=SUMIFS(Inventory[Value],Inventory[Category],J11)/SUM(Inventory[Value])',
      answer: () => ({ kind: 'cells', range: 'K11:K15', format: FMT.pct, consistency: 'all' }),
      expected: (d) => {
        const total = sumValue(d.rows);
        return d.categories.map((c) => [sumValue(d.rows.filter((r) => r.category === c)) / total]);
      },
      variants: [
        stockLevelsChange,
        {
          label: 'the categories in column J are reordered',
          explain: 'Point at the category in column J instead of typing it.',
          apply: (d, rng) => ({ ...d, categories: reorderLabels(rng, d.categories) }),
        },
        {
          label: 'four new stock lines are added',
          explain: 'Use whole Table columns in both parts of the fraction so new rows count.',
          apply: addStockLines,
        },
      ],
      rules: {
        forbidText: { values: [...CATEGORIES], advice: 'Point to the category in column J instead.' },
        allowNumbers: [],
      },
    },

    // ---------- 4. Sorted FILTER ----------
    {
      title: 'Reorder list for purchasing',
      task: () =>
        'With one formula in `J18`, list every SKU that is at or below its reorder point (On hand less than or equal to Reorder point). Return the SKU, Item, Category and Warehouse columns, sorted by SKU from A to Z. The headers are already in row `17`, so return the rows only.',
      hints: [
        'FILTER keeps the rows that pass a test. The test here is Inventory[On hand]<=Inventory[Reorder point].',
        'For the columns to return, Inventory[[SKU]:[Warehouse]] means every column from SKU through Warehouse.',
        'Wrap it in SORT, which sorts by the first column: =SORT(FILTER(Inventory[[SKU]:[Warehouse]], Inventory[On hand]<=Inventory[Reorder point]))',
      ],
      solution: () => '=SORT(FILTER(Inventory[[SKU]:[Warehouse]],Inventory[On hand]<=Inventory[Reorder point]))',
      answer: () => ({ kind: 'spill', anchor: 'J18' }),
      expected: (d) =>
        d.rows
          .filter(isLow)
          .sort((a, b) => excelTextCompare(a.sku, b.sku))
          .map((r) => [r.sku, r.item, r.category, r.warehouse]),
      variants: [
        stockLevelsChange,
        reorderPointsChange,
        {
          label: 'the Inventory Table is re-sorted',
          explain: 'Wrap the FILTER in SORT so the list comes out in SKU order whatever order the Table is in.',
          apply: (d, rng) => ({ ...d, rows: reshuffle(rng, d.rows, (r) => r.sku) }),
        },
        {
          label: 'four new stock lines are added',
          explain: 'Filter whole Table columns so new rows are included.',
          apply: addStockLines,
        },
      ],
      rules: {
        require: [
          { pattern: /FILTER\(/i, label: 'Uses FILTER', advice: 'FILTER(array, include) returns every matching row in one go.' },
          { pattern: /SORT(BY)?\(/i, label: 'Sorts the list by SKU', advice: 'Wrap the FILTER in SORT( … ).' },
        ],
      },
    },
  ],
});

// =====================================================================================
// Mission 2: Weekly reorder plan
// =====================================================================================

/** Days of demand history; also written to N1 for the learner to divide by. */
export const PLAN_DAYS = 28;
const PLAN_END = serial(2026, 9, 27);
const PLAN_START = PLAN_END - (PLAN_DAYS - 1);
const PLAN_ITEMS = 12;
const DEMAND_LINES = 96;

export interface DemandLine {
  date: number;
  sku: string;
  units: number;
}

export interface PlanItem {
  sku: string;
  lead: number;
  safety: number;
  onHand: number;
  /** How often the SKU ships (0 = no demand in the window). Not shown on the sheet. */
  weight: number;
  /** Largest single shipment line. Not shown on the sheet. */
  maxUnits: number;
}

export interface PlanData {
  demand: DemandLine[];
  items: PlanItem[];
}

const DEMAND_COLS: ColumnSpec[] = [{ header: 'Date', format: FMT.date }, { header: 'SKU' }, { header: 'Units', format: FMT.int }];

/**
 * No number formats here on purpose: when the learner types a new header in I1, Excel extends the
 * Table and the new column picks up the format of the column beside it. With General on On hand,
 * Avg daily demand shows its decimals instead of being rounded for display by #,##0.
 */
const ITEM_COLS: ColumnSpec[] = [{ header: 'SKU' }, { header: 'Lead days' }, { header: 'Safety days' }, { header: 'On hand' }];

const demandGrid = (rows: DemandLine[]): Grid => rows.map((l) => [l.date, l.sku, l.units]);
const itemGrid = (rows: PlanItem[]): Grid => rows.map((i) => [i.sku, i.lead, i.safety, i.onHand]);

/**
 * Lead and safety days whose sum is never a multiple of the history window, so the reorder point
 * can be kept off exact whole numbers (see fixDivisibility).
 */
function leadDays(rng: Rng, safety: number): number {
  let lead = rng.int(3, 21);
  while ((lead + safety) % PLAN_DAYS === 0) lead = rng.int(3, 21);
  return lead;
}

function safetyDays(rng: Rng, lead: number): number {
  let safety = rng.int(2, 10);
  while ((lead + safety) % PLAN_DAYS === 0) safety = rng.int(2, 10);
  return safety;
}

/** `days` plus one (or two), keeping days + other off a multiple of the window. */
function bumped(days: number, other: number): number {
  return (days + 1 + other) % PLAN_DAYS === 0 ? days + 2 : days + 1;
}

function planItem(rng: Rng, sku: string): PlanItem {
  const safety = rng.int(2, 10);
  return { sku, lead: leadDays(rng, safety), safety, onHand: 0, weight: rng.int(1, 6), maxUnits: rng.pick([6, 10, 18, 30, 48]) };
}

function demandLine(rng: Rng, items: PlanItem[]): DemandLine {
  const live = items.filter((i) => i.weight > 0);
  let pick = rng.int(1, sum(live.map((i) => i.weight)));
  const it = live.find((i) => (pick -= i.weight) <= 0)!;
  return { date: rng.int(PLAN_START, PLAN_END), sku: it.sku, units: rng.int(1, it.maxUnits) };
}

export const totalUnits = (d: PlanData, sku: string) => sum(d.demand.filter((l) => l.sku === sku).map((l) => l.units));

/** ROUNDUP(total ÷ days × (lead + safety), 0), done in whole numbers so float noise can't creep in. */
export function planReorderPoint(total: number, leadPlusSafety: number, days = PLAN_DAYS): number {
  return Math.ceil((total * leadPlusSafety) / days);
}

export const reorderPointOf = (d: PlanData, it: PlanItem) => planReorderPoint(totalUnits(d, it.sku), it.lead + it.safety);
export const orderQtyOf = (d: PlanData, it: PlanItem) => Math.max(0, reorderPointOf(d, it) - it.onHand);

/**
 * When total × (lead + safety) is an exact multiple of 28, the true reorder point is a whole
 * number and Excel could see 20.000000000000004 and round it up to 21. Nudging one demand line by
 * a unit moves every reorder point at least 1/28 away from a whole number, so ROUNDUP is exact.
 */
export function fixDivisibility(d: PlanData): PlanData {
  const demand = d.demand.map((l) => ({ ...l }));
  for (const it of d.items) {
    const lines = demand.filter((l) => l.sku === it.sku);
    const total = sum(lines.map((l) => l.units));
    if (total > 0 && (total * (it.lead + it.safety)) % PLAN_DAYS === 0) lines[0].units += 1;
  }
  return { ...d, demand };
}

/** On-hand stock: a mix of SKUs well below, well above and exactly at their reorder point. */
function freshOnHand(rng: Rng, d: PlanData): PlanItem[] {
  const withDemand = d.items.map((it, i) => ({ i, rp: reorderPointOf(d, it) })).filter((x) => x.rp > 0);
  const tie = withDemand.length ? rng.pick(withDemand).i : -1;
  return d.items.map((it, i) => {
    const rp = reorderPointOf(d, it);
    if (rp === 0) return { ...it, onHand: rng.int(4, 40) };
    if (i === tie) return { ...it, onHand: rp };
    return { ...it, onHand: rng.chance(0.5) ? rng.int(0, Math.floor(rp * 0.6)) : rng.int(rp + 1, rp * 2 + 10) };
  });
}

/** At least four SKUs need an order, so the step-4 list is never empty or trivially short. */
export function ensureNeeds(d: PlanData): PlanData {
  const items = d.items.map((it) => ({ ...it }));
  const next = { ...d, items };
  for (const it of items) {
    if (items.filter((x) => orderQtyOf(next, x) > 0).length >= 4) break;
    const rp = reorderPointOf(next, it);
    if (rp >= 1 && orderQtyOf(next, it) === 0) it.onHand = Math.floor(rp / 3);
  }
  return next;
}

const settle = (d: PlanData) => ensureNeeds(fixDivisibility(d));

function newPlanSkus(rng: Rng, taken: Set<string>, count: number): string[] {
  const out: string[] = [];
  while (out.length < count) {
    const sku = `SKU-${rng.int(3000, 3999)}`;
    if (!taken.has(sku)) {
      taken.add(sku);
      out.push(sku);
    }
  }
  return out;
}

function changed<T>(before: T[], after: T[], bump: (x: T) => T): T[] {
  return JSON.stringify(before) === JSON.stringify(after) ? [bump(after[0]), ...after.slice(1)] : after;
}

const demandChanges = {
  label: 'demand changes',
  apply: (d: PlanData, rng: Rng): PlanData => {
    const max = new Map(d.items.map((i) => [i.sku, i.maxUnits]));
    const demand = d.demand.map((l) => ({ ...l, units: rng.int(1, max.get(l.sku) ?? 12) }));
    return settle({ ...d, demand: changed(d.demand, demand, (l) => ({ ...l, units: l.units + 1 })) });
  },
};

const moreDemand = {
  label: 'eight new demand lines are added',
  explain: 'Use whole Table columns like Demand[Units] so new lines are counted.',
  apply: (d: PlanData, rng: Rng): PlanData =>
    settle({ ...d, demand: [...d.demand, ...Array.from({ length: 8 }, () => demandLine(rng, d.items))] }),
};

const newSku = {
  label: 'a new SKU is added to the Items Table',
  explain: 'A calculated column fills new rows by itself. Formulas typed cell by cell don’t.',
  apply: (d: PlanData, rng: Rng): PlanData => {
    const [sku] = newPlanSkus(rng, new Set(d.items.map((i) => i.sku)), 1);
    const it = { ...planItem(rng, sku), weight: 4 };
    // A new line with nothing on the shelf yet: it always needs an order.
    const lines = Array.from({ length: 6 }, () => demandLine(rng, [it]));
    return settle({ demand: [...d.demand, ...lines], items: [...d.items, it] });
  },
};

const onHandChanges = {
  label: 'stock on hand changes',
  apply: (d: PlanData, rng: Rng): PlanData => {
    const items = freshOnHand(rng, d);
    return ensureNeeds({ ...d, items: changed(d.items, items, (i) => ({ ...i, onHand: i.onHand + 7 })) });
  },
};

const sortedSkus = (d: PlanData) =>
  d.items
    .filter((it) => orderQtyOf(d, it) > 0)
    .map((it) => it.sku)
    .sort(excelTextCompare);

export const reorderPlan = defineMission<PlanData>({
  id: 'mission-reorder-plan',
  title: 'Weekly reorder plan',
  role: 'ops',
  summary: 'The weekly purchasing plan: daily demand, reorder points and what to order',
  minutes: 18,
  skills: ['tables-calc-column', 'sumifs-warehouse', 'let-reorder', 'filter-late', 'unique-vendors'],
  brief: (d) => ({
    from: 'Sam Okafor, Purchasing manager',
    subject: 'Reorder plan for Monday’s PO run',
    body: [
      'Morning,',
      '',
      `I place this week’s purchase orders on Monday. The last 28 days of shipments are in the \`Demand\` Table, and stock on hand, lead times and safety days for our ${d.items.length} replenished SKUs are in \`Items\`. Add the plan as three new columns of the \`Items\` Table, starting in \`I1\`, so it recalculates when I paste in next week’s history.`,
      '',
      '1. `Avg daily demand`: total units ÷ the 28 days in `N1`.',
      '2. `Reorder point`: that average × (lead days + safety days), rounded up to a whole unit.',
      '3. `Order qty`: how far on hand is below the reorder point, never less than 0.',
      '4. A sorted list of the SKUs I need to order, starting at `M4`.',
      '',
      'Thanks,',
      'Sam',
    ].join('\n'),
  }),
  make: (rng) => {
    const skus = newPlanSkus(rng, new Set(), PLAN_ITEMS);
    const items = skus.map((sku) => planItem(rng, sku));
    // One slow mover had no shipments in the window: its plan should come out as zeros.
    items[rng.int(0, items.length - 1)].weight = 0;
    const demand = Array.from({ length: DEMAND_LINES }, () => demandLine(rng, items)).sort((a, b) => a.date - b.date);
    const fixed = fixDivisibility({ demand, items });
    return ensureNeeds({ ...fixed, items: freshOnHand(rng, fixed) });
  },
  blocks: (d) => [
    dataBlock('Demand', 'A1', DEMAND_COLS, demandGrid(d.demand)),
    dataBlock('Items', 'E1', ITEM_COLS, itemGrid(d.items)),
    cells('M1', [['Days of history']], 'label'),
    cells('N1', [[PLAN_DAYS]], 'input', FMT.int),
    cells('M3', [['SKUs to order']], 'header'),
  ],
  inputs: (d) => [tableWrite('Demand', DEMAND_COLS, demandGrid(d.demand)), tableWrite('Items', ITEM_COLS, itemGrid(d.items))],
  steps: [
    // ---------- 1. Average daily demand ----------
    {
      title: 'Average daily demand',
      task: () =>
        'Add a column named `Avg daily demand` to the `Items` Table by typing the name in `I1`. Fill it with each SKU’s average daily demand: its total Units in the `Demand` Table divided by the days of history in `N1`. Divide by all 28 days, including days with no shipments; a SKU with no demand shows 0.',
      hints: [
        'Type Avg daily demand in I1 and press {enter}. The Table grows to include the new column.',
        'In I2, total the SKU’s units with SUMIFS(Demand[Units], Demand[SKU], [@SKU]), then divide by $N$1. The $ signs keep every row pointing at N1; without them the next row down would read N2.',
        '=SUMIFS(Demand[Units], Demand[SKU], [@SKU])/$N$1. Press {enter} once and the Table fills the rest of the column.',
      ],
      solution: () => '=SUMIFS(Demand[Units],Demand[SKU],[@SKU])/$N$1',
      answer: () => ({ kind: 'tableColumn', table: 'Items', column: 'Avg daily demand', format: '0.00' }),
      expected: (d) => d.items.map((it) => [totalUnits(d, it.sku) / PLAN_DAYS]),
      variants: [demandChanges, moreDemand, newSku],
      rules: { allowNumbers: [] },
    },

    // ---------- 2. Reorder point ----------
    {
      title: 'Reorder point',
      task: () =>
        'Add a column named `Reorder point` in `J1`. Fill it with average daily demand × (Lead days + Safety days), rounded up to a whole unit. A SKU with no demand shows 0.',
      hints: [
        'Type Reorder point in J1 and press {enter}. In J2, build on the column you made: [@[Avg daily demand]] is the average on the same row.',
        'Add lead and safety days in parentheses first, ([@[Lead days]]+[@[Safety days]]), multiply, then round up with ROUNDUP(…, 0).',
        '=ROUNDUP([@[Avg daily demand]]*([@[Lead days]]+[@[Safety days]]), 0)',
      ],
      solution: () => '=ROUNDUP([@[Avg daily demand]]*([@[Lead days]]+[@[Safety days]]),0)',
      answer: () => ({ kind: 'tableColumn', table: 'Items', column: 'Reorder point', format: FMT.int }),
      expected: (d) => d.items.map((it) => [reorderPointOf(d, it)]),
      variants: [
        demandChanges,
        {
          label: 'lead times change',
          apply: (d, rng) => {
            const items = d.items.map((i) => ({ ...i, lead: leadDays(rng, i.safety) }));
            return settle({ ...d, items: changed(d.items, items, (i) => ({ ...i, lead: bumped(i.lead, i.safety) })) });
          },
        },
        {
          label: 'safety days change',
          apply: (d, rng) => {
            const items = d.items.map((i) => ({ ...i, safety: safetyDays(rng, i.lead) }));
            return settle({ ...d, items: changed(d.items, items, (i) => ({ ...i, safety: bumped(i.safety, i.lead) })) });
          },
        },
      ],
      rules: { allowNumbers: [0, 1] },
    },

    // ---------- 3. Order quantity ----------
    {
      title: 'Order quantity',
      task: () =>
        'Add a column named `Order qty` in `K1` showing how many units to order: Reorder point minus On hand, or 0 when On hand already covers the reorder point.',
      hints: [
        'Type Order qty in K1 and press {enter}. In K2, the gap is [@[Reorder point]]-[@[On hand]].',
        'When stock is above the reorder point, that gap is negative. MAX(0, …) turns any negative into 0.',
        '=MAX(0, [@[Reorder point]]-[@[On hand]])',
      ],
      solution: () => '=MAX(0,[@[Reorder point]]-[@[On hand]])',
      answer: () => ({ kind: 'tableColumn', table: 'Items', column: 'Order qty', format: FMT.int }),
      expected: (d) => d.items.map((it) => [orderQtyOf(d, it)]),
      variants: [onHandChanges, demandChanges, newSku],
      rules: { allowNumbers: [0] },
    },

    // ---------- 4. Sorted FILTER ----------
    {
      title: 'SKUs to order this week',
      task: () => 'With one formula in `M4`, list the SKUs whose Order qty is above 0, sorted from A to Z. The header is already in `M3`, so return the SKUs only.',
      hints: [
        'FILTER returns the SKUs that pass a test: FILTER(Items[SKU], Items[Order qty]>0).',
        'Wrap the FILTER in SORT to put the list in A to Z order.',
        '=SORT(FILTER(Items[SKU], Items[Order qty]>0))',
      ],
      solution: () => '=SORT(FILTER(Items[SKU],Items[Order qty]>0))',
      answer: () => ({ kind: 'spill', anchor: 'M4' }),
      expected: (d) => sortedSkus(d).map((sku) => [sku]),
      variants: [
        onHandChanges,
        demandChanges,
        {
          label: 'the Items Table is re-sorted',
          explain: 'Wrap the FILTER in SORT so the list stays in A to Z order.',
          apply: (d, rng) => ({ ...d, items: reshuffle(rng, d.items, (i) => i.sku) }),
        },
        { ...newSku, explain: 'Filter whole Table columns so a new SKU is included.' },
      ],
      rules: {
        require: [
          { pattern: /FILTER\(/i, label: 'Uses FILTER', advice: 'FILTER(Items[SKU], include) returns every matching SKU in one go.' },
          { pattern: /SORT(BY)?\(/i, label: 'Sorts the list A to Z', advice: 'Wrap the FILTER in SORT( … ).' },
        ],
        allowNumbers: [0, 1],
      },
    },
  ],
});

/** Ops missions in the order they're offered. */
export const OPS_MISSIONS: Mission<any>[] = [monthEndInventory, reorderPlan];

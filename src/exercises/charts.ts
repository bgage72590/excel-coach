import { CARRIERS, CATEGORIES, ITEMS, WAREHOUSES, serial, skuCode, sum } from '../engine/data';
import { round, type Rng } from '../engine/rng';
import type { ColumnSpec, Exercise } from '../engine/types';
import { FMT, dataBlock, defineExercise } from './common';
import { checkStep, money } from './guides';
import { clickTableStep, dragStep, insertPivotStep } from './pivots';

/** Whole dollars, for monthly revenue. */
const USD0 = '$#,##0';

// ---------- Revenue and margin in one chart ----------

const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'] as const;
/** A distributor's year: soft January, busy autumn, quiet December. */
const SEASON = [0.86, 0.88, 0.97, 1.0, 1.04, 1.06, 1.02, 1.0, 1.05, 1.08, 1.12, 0.92] as const;

export interface MonthRow {
  month: string;
  revenue: number;
  margin: number;
}

const MONTH_COLS: ColumnSpec[] = [{ header: 'Month' }, { header: 'Revenue', format: USD0 }, { header: 'Margin %', format: FMT.pct }];

function monthlyResults(rng: Rng): MonthRow[] {
  const base = rng.int(18, 42) * 10_000;
  return MONTHS.map((month, i) => ({
    month,
    revenue: Math.round((base * SEASON[i] * rng.float(0.93, 1.07, 3)) / 100) * 100,
    margin: rng.float(0.22, 0.38, 3),
  }));
}

export const comboChart = defineExercise<{ rows: MonthRow[] }>({
  id: 'combo-chart',
  module: 'charts',
  title: 'Revenue and margin in one chart',
  replaces: 'Two separate charts pasted side by side in the monthly deck',
  minutes: 4,
  task: () =>
    'Chart Revenue and Margin % by month in one chart on this sheet: Revenue as columns and Margin % as a line on the secondary axis, so both stay readable despite their different scales. Give it a title that says what it shows. Graded: a chart on the practice sheet with at least two series, one of them a line on the secondary axis, and a title other than “Chart Title”.',
  concept: {
    summary:
      'A combo chart mixes chart types and can give one series its own axis. Revenue in dollars and margin in percent can then share a chart without the percentages flattening into a line along the bottom.',
    syntax: 'Insert › Combo Chart › Clustered Column – Line on Secondary Axis',
    example: 'Columns show revenue in dollars against the left axis; the margin line reads in percent against the right axis.',
    tip: 'To fix a chart you already made, select it, then Chart Design › Change Chart Type › Combo › Clustered Column – Line on Secondary Axis.',
  },
  hints: [
    'Click any cell in the Monthly Table so Excel picks up all three columns.',
    'Insert › Combo Chart (the icon with columns and a line) › Clustered Column – Line on Secondary Axis. Check that Margin % is the line.',
    'Click the Chart Title placeholder and type a title, such as Revenue and margin by month. If there’s no placeholder, use Chart Design › Add Chart Element › Chart Title.',
  ],
  solution: () =>
    'Click in the data › Insert › Combo Chart › Clustered Column – Line on Secondary Axis (Revenue as columns, Margin % as a line on the secondary axis) › type a title such as “Revenue and margin by month”.',
  guide: (d) => {
    const revenue = d.rows.map((r) => r.revenue);
    const margin = d.rows.map((r) => r.margin);
    const usd0 = (n: number) => `$${n.toLocaleString('en-US')}`;
    const pct1 = (n: number) => `${(n * 100).toFixed(1)}%`;
    return [
      {
        do: 'Meet the data. The Table named **Monthly** holds twelve months of Revenue (column `B`) and Margin % (column `C`).',
        why: `Revenue runs from ${usd0(Math.min(...revenue))} to ${usd0(Math.max(...revenue))}, while Margin % stays between ${pct1(Math.min(...margin))} and ${pct1(Math.max(...margin))}. On one shared axis the margin line would lie flat along the bottom. A **secondary axis**, a second scale on the right side of the chart, gives it room to move.`,
        show: [
          { label: 'Revenue column', at: 'Monthly[Revenue]', note: 'Dollars in the hundreds of thousands: these become the columns.' },
          { label: 'Margin % column', at: 'Monthly[Margin %]', note: 'Percentages under 100%: this becomes the line on the secondary axis.' },
        ],
      },
      {
        do: 'Click any cell in the **Monthly** Table, for example `B5`.',
        why: 'Excel then picks up all three columns for the chart: months along the bottom, and Revenue and Margin % as its two **series** (the sets of numbers the chart draws).',
        done: { kind: 'select', range: `A1:C${d.rows.length + 1}` },
      },
      {
        do: 'Click **Insert › Combo Chart** (the icon with columns and a line), then **Clustered Column – Line on Secondary Axis**.',
        why: 'Revenue becomes columns read against the left axis, and Margin % becomes a line read against a new axis on the right. The legend under the chart shows which is which: Margin % should be the line. If you already made another kind of chart, select it and use **Chart Design › Change Chart Type › Combo** instead.',
        done: { kind: 'inspect', inspection: { kind: 'chart', label: 'A combo chart with a secondary-axis line', minSeries: 2, secondaryLine: true } },
      },
      {
        do: 'Click the **Chart Title** placeholder at the top of the chart, type a title that says what it shows, such as `Revenue and margin by month`, and press {enter}.',
        why: 'If your words land beside the words “Chart Title” instead of replacing them, delete the leftovers. No placeholder? Use **Chart Design › Add Chart Element › Chart Title** first.',
        done: { kind: 'inspect', inspection: { kind: 'chart', label: 'A chart on the practice sheet', minSeries: 2, secondaryLine: true, title: true } },
      },
      checkStep('The coach looks for a chart on this sheet with at least two series, one of them a line on the secondary axis, and a title other than “Chart Title”.'),
    ];
  },
  make: (rng) => ({ rows: monthlyResults(rng) }),
  layout: (d) => ({
    blocks: [dataBlock('Monthly', 'A1', MONTH_COLS, d.rows.map((r) => [r.month, r.revenue, r.margin]))],
    answer: { kind: 'objects' },
  }),
  expected: () => [],
  inputs: () => [],
  variants: [],
  inspections: () => [{ kind: 'chart', label: 'A chart on the practice sheet', minSeries: 2, secondaryLine: true, title: true }],
});

// ---------- A PivotTable with a slicer ----------

export interface StockRow {
  sku: string;
  item: string;
  category: string;
  warehouse: string;
  onHand: number;
  cost: number;
  value: number;
}

const STOCK_COLS: ColumnSpec[] = [
  { header: 'SKU' },
  { header: 'Item' },
  { header: 'Category' },
  { header: 'Warehouse' },
  { header: 'On hand', format: FMT.int },
  { header: 'Unit cost', format: FMT.currency },
  { header: 'Value', format: FMT.currency },
];

/** The first rows cover every warehouse and category, so each slicer button and PivotTable row has stock. */
function inventory(rng: Rng): StockRow[] {
  return Array.from({ length: 48 }, (_, i) => {
    const warehouse = i < WAREHOUSES.length ? WAREHOUSES[i] : rng.pick(WAREHOUSES);
    const it = i < CATEGORIES.length ? rng.pick(ITEMS.filter((x) => x.category === CATEGORIES[i])) : rng.pick(ITEMS);
    const onHand = rng.int(10, 900);
    return {
      sku: skuCode(100 + ITEMS.indexOf(it) * 7),
      item: it.item,
      category: it.category,
      warehouse,
      onHand,
      cost: it.cost,
      value: round(onHand * it.cost, 2),
    };
  });
}

export const slicerDashboard = defineExercise<{ rows: StockRow[] }>({
  id: 'slicer-dashboard',
  module: 'charts',
  title: 'A PivotTable with a slicer',
  replaces: 'Re-filtering the stock report for each warehouse manager',
  minutes: 5,
  task: () =>
    'Build a small dashboard from the `Inventory` Table: a PivotTable with Category in Rows and the sum of Value in Values, plus a slicer for Warehouse, so anyone can click a warehouse and see its stock value by category. Graded: Category alone in Rows, Sum of Value in Values, nothing in Columns, and a Warehouse slicer.',
  concept: {
    summary:
      'A slicer is a panel of buttons that filters a PivotTable with one click. It always shows which items are selected, which a drop-down filter hides, so it suits reports other people use.',
    syntax: 'Click in the PivotTable › PivotTable Analyze › Insert Slicer › tick the field',
    example: 'Click Reno on the Warehouse slicer to see only Reno’s stock value by category. The Clear Filter button in the slicer’s corner brings every warehouse back.',
  },
  hints: [
    'Click inside the Inventory Table, then Insert › PivotTable › OK. A new sheet is fine.',
    'In the PivotTable Fields list, drag Category to Rows and Value to Values. The value field should read Sum of Value.',
    'Click inside the PivotTable, then PivotTable Analyze › Insert Slicer, tick Warehouse and click OK.',
  ],
  solution: () => 'Insert › PivotTable from Inventory · Rows: Category · Values: Sum of Value · PivotTable Analyze › Insert Slicer › Warehouse',
  guide: (d) => {
    // PivotTables and slicers list items A to Z, so these are the top row and the first button.
    const category = [...new Set(d.rows.map((r) => r.category))].sort()[0];
    const warehouse = [...new Set(d.rows.map((r) => r.warehouse))].sort()[0];
    const categoryValue = sum(d.rows.filter((r) => r.category === category).map((r) => r.value));
    const warehouseValue = sum(d.rows.filter((r) => r.warehouse === warehouse).map((r) => r.value));
    return [
      {
        do: 'Meet the data. The Table in columns `A` to `G` is named **Inventory**: one row per stock line, with its category, warehouse and value.',
        why: 'The PivotTable will total Value for each Category. The slicer will be a panel of Warehouse buttons that filters it with one click.',
        show: [
          { label: 'Category column', at: 'Inventory[Category]', note: 'Each category will get its own row in the PivotTable.' },
          { label: 'Value column', at: 'Inventory[Value]', note: 'The numbers the PivotTable adds up.' },
          { label: 'Warehouse column', at: 'Inventory[Warehouse]', note: 'The slicer will show one button for each warehouse here.' },
        ],
      },
      clickTableStep('Inventory', `A1:G${d.rows.length + 1}`, 'C5'),
      insertPivotStep('Inventory'),
      dragStep('Category', 'Rows', 'The category names appear down the left of the PivotTable, one row each.'),
      dragStep(
        'Value',
        'Values',
        `The field should read **Sum of Value**. ${category}’s row should come to ${money(categoryValue)} (the PivotTable may show it without the $ sign).`,
        { kind: 'pivot', rows: 'Category', valuesField: 'Value', summarizeBy: 'Sum', showAs: 'None', label: 'PivotTable layout' },
      ),
      {
        do: 'With a cell in the PivotTable selected, click **PivotTable Analyze › Insert Slicer**.',
        why: 'The **PivotTable Analyze** tab only shows while a cell in the PivotTable is selected. A box opens listing every field.',
      },
      {
        do: 'Tick **Warehouse** and click **OK**.',
        why: 'A panel of warehouse buttons appears on the sheet. Drag it beside the PivotTable if it covers the numbers.',
        done: { kind: 'inspect', inspection: { kind: 'slicer', field: 'Warehouse', label: 'A slicer for Warehouse' } },
      },
      {
        do: `Try it: click **${warehouse}** on the slicer.`,
        why: `The PivotTable now counts only ${warehouse}’s stock, and its **Grand Total** drops to ${money(warehouseValue)}. Click the **Clear Filter** button in the slicer’s corner to bring every warehouse back.`,
      },
      checkStep('The coach finds your PivotTable on any sheet and reads its layout: Category alone in Rows, Sum of Value in Values, nothing in Columns, plus a Warehouse slicer.'),
    ];
  },
  make: (rng) => ({ rows: inventory(rng) }),
  layout: (d) => ({
    blocks: [dataBlock('Inventory', 'A1', STOCK_COLS, d.rows.map((r) => [r.sku, r.item, r.category, r.warehouse, r.onHand, r.cost, r.value]))],
    answer: { kind: 'pivot' },
  }),
  expected: () => [],
  inputs: () => [],
  variants: [],
  inspections: () => [
    { kind: 'pivot', rows: 'Category', valuesField: 'Value', summarizeBy: 'Sum', showAs: 'None', label: 'PivotTable layout' },
    { kind: 'slicer', field: 'Warehouse', label: 'A slicer for Warehouse' },
  ],
});

// ---------- Average delay by carrier, filtered by warehouse ----------

export interface Shipment {
  id: string;
  date: number;
  carrier: string;
  warehouse: string;
  late: number;
}

const SHIP_COLS: ColumnSpec[] = [
  { header: 'Ship ID' },
  { header: 'Ship date', format: FMT.date },
  { header: 'Carrier' },
  { header: 'Warehouse' },
  { header: 'Days late', format: FMT.plain },
];

function shipments(rng: Rng): Shipment[] {
  // Each carrier has its own habit, so the averages differ in a way worth reporting.
  const habit = new Map(CARRIERS.map((c) => [c, rng.float(-0.5, 2.5, 1)]));
  return Array.from({ length: 80 }, (_, i) => {
    const carrier = i < CARRIERS.length ? CARRIERS[i] : rng.pick(CARRIERS);
    const warehouse = i < WAREHOUSES.length ? WAREHOUSES[i] : rng.pick(WAREHOUSES);
    const late = Math.max(-2, Math.min(9, Math.round(habit.get(carrier)! + rng.int(-2, 3))));
    return { id: `SH-${48200 + i}`, date: rng.int(serial(2026, 7, 1), serial(2026, 9, 30)), carrier, warehouse, late };
  });
}

export const pivotAverageFilter = defineExercise<{ rows: Shipment[] }>({
  id: 'pivot-average-filter',
  module: 'pivots',
  title: 'Average delay by carrier, filtered by warehouse',
  replaces: 'An AVERAGEIFS grid rebuilt whenever someone asks about one warehouse',
  minutes: 4,
  task: () =>
    'From the `Dispatch` Table of outbound deliveries, build a PivotTable that shows the average Days late for each carrier, with Warehouse as a filter so you can look at one warehouse at a time. Graded: Carrier alone in Rows, Average of Days late in Values, Warehouse in Filters and nothing in Columns.',
  concept: {
    summary:
      'Values in a PivotTable don’t have to be sums. Summarize Values By switches a field to an average, count, max or min, and a field in Filters adds a drop-down that limits the whole PivotTable to the items you pick.',
    syntax: 'Rows = Carrier · Values = Days late, summarized by Average · Filters = Warehouse',
    example: 'Right-click a value › Summarize Values By › Average turns Sum of Days late into Average of Days late.',
    tip: 'Early deliveries count as negative days late, so they pull a carrier’s average down.',
  },
  hints: [
    'Click inside the Dispatch Table, then Insert › PivotTable › OK.',
    'Drag Carrier to Rows, Days late to Values and Warehouse to Filters.',
    'Values starts as Sum of Days late. Right-click any number in the PivotTable › Summarize Values By › Average, or open the field’s settings from the Values box (the i button on a Mac, Value Field Settings on Windows).',
  ],
  solution: () => 'Insert › PivotTable from Dispatch · Rows: Carrier · Values: Days late, Summarize Values By › Average · Filters: Warehouse',
  guide: (d) => {
    // PivotTables list row labels A to Z, so this carrier is the top row.
    const carrier = [...new Set(d.rows.map((r) => r.carrier))].sort()[0];
    const late = d.rows.filter((r) => r.carrier === carrier).map((r) => r.late);
    const total = sum(late);
    const average = (total / late.length).toFixed(2);
    return [
      {
        do: 'Meet the data. The Table in columns `A` to `E` is named **Dispatch**: one row per delivery, with its carrier, warehouse and how many days late it arrived.',
        why: 'Early deliveries count as negative days late, so they pull a carrier’s average down.',
        show: [
          { label: 'Carrier column', at: 'Dispatch[Carrier]', note: 'Each carrier will get its own row in the PivotTable.' },
          { label: 'Days late column', at: 'Dispatch[Days late]', note: 'The numbers the PivotTable will average.' },
          { label: 'Warehouse column', at: 'Dispatch[Warehouse]', note: 'This becomes the filter: one warehouse at a time.' },
        ],
      },
      clickTableStep('Dispatch', `A1:E${d.rows.length + 1}`, 'C5'),
      insertPivotStep('Dispatch'),
      dragStep('Carrier', 'Rows', 'The carrier names appear down the left of the PivotTable, one row each.'),
      dragStep(
        'Days late',
        'Values',
        `It starts as **Sum of Days late**: ${carrier}’s row reads ${total}. A total of days says little about a carrier. An average says how late a typical delivery is, and you’ll switch to it in a moment.`,
      ),
      dragStep(
        'Warehouse',
        'Filters',
        'A **Warehouse** drop-down appears above the PivotTable, set to (All). It limits the whole PivotTable to the warehouses you pick.',
        { kind: 'pivot', rows: 'Carrier', valuesField: 'Days late', summarizeBy: 'Sum', showAs: 'None', filter: 'Warehouse', label: 'Carrier, Sum of Days late and the Warehouse filter are in place' },
      ),
      {
        do: 'Right-click any number in the PivotTable, then choose **Summarize Values By › Average**.',
        why: `The field becomes **Average of Days late**, and ${carrier}’s row should read about ${average}. Another way in: open the field’s settings from the **Values** box (the **i** button on a Mac, **Value Field Settings** on Windows).`,
        done: {
          kind: 'inspect',
          inspection: {
            kind: 'pivot',
            rows: 'Carrier',
            valuesField: 'Days late',
            summarizeBy: 'Average',
            showAs: 'None',
            filter: 'Warehouse',
            label: 'PivotTable layout',
          },
        },
      },
      {
        do: 'Try the filter: use the **Warehouse** drop-down above the PivotTable to pick one warehouse.',
        why: 'Every average now counts only that warehouse’s deliveries. Set it back to all warehouses the same way when you’re done looking.',
      },
      checkStep('The coach finds your PivotTable on any sheet and reads its layout: Carrier alone in Rows, Average of Days late in Values, Warehouse in Filters and nothing in Columns.'),
    ];
  },
  make: (rng) => ({ rows: shipments(rng) }),
  layout: (d) => ({
    blocks: [dataBlock('Dispatch', 'A1', SHIP_COLS, d.rows.map((r) => [r.id, r.date, r.carrier, r.warehouse, r.late]))],
    answer: { kind: 'pivot' },
  }),
  expected: () => [],
  inputs: () => [],
  variants: [],
  inspections: () => [
    {
      kind: 'pivot',
      rows: 'Carrier',
      valuesField: 'Days late',
      summarizeBy: 'Average',
      showAs: 'None',
      filter: 'Warehouse',
      label: 'PivotTable layout',
    },
  ],
});

export const CHARTS: Exercise<any>[] = [comboChart, slicerDashboard, pivotAverageFilter];

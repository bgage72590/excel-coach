import { ITEMS, WAREHOUSES, sum } from '../engine/data';
import type { Rng } from '../engine/rng';
import type { ColumnSpec, GuideStep, Inspection } from '../engine/types';
import { FMT, dataBlock, defineExercise } from './common';
import { checkStep } from './guides';

// ---------- walkthrough steps shared by the PivotTable skills ----------

/** "Click any cell in the Table": ticks off once the active cell is inside `range`. */
export function clickTableStep(table: string, range: string, example: string): GuideStep {
  return {
    do: `Click any cell in the **${table}** Table, for example \`${example}\`.`,
    why: 'Excel then knows which data to summarize, and picks up every row and column of the Table.',
    done: { kind: 'select', range },
  };
}

/** Insert › PivotTable › OK, as the hints word it. */
export function insertPivotStep(table: string): GuideStep {
  return {
    do: `Click **Insert › PivotTable**. The box that opens already names the ${table} Table. Leave **New Worksheet** selected and click **OK**.`,
    why: `Excel adds a sheet with an empty PivotTable on the left and the **PivotTable Fields** list on the right. Each column of ${table} is listed there as a field. The coach finds the PivotTable on any sheet.`,
  };
}

/** One drag in the PivotTable Fields list. */
export function dragStep(field: string, area: 'Rows' | 'Columns' | 'Values' | 'Filters', why: string, inspection?: Inspection): GuideStep {
  const step: GuideStep = { do: `In the **PivotTable Fields** list, drag **${field}** down into the **${area}** box.`, why };
  return inspection ? { ...step, done: { kind: 'inspect', inspection } } : step;
}

interface Outbound {
  month: string;
  warehouse: string;
  sku: string;
  units: number;
}

const OUTBOUND_COLS: ColumnSpec[] = [{ header: 'Month' }, { header: 'Warehouse' }, { header: 'SKU' }, { header: 'Units', format: FMT.int }];

function outbound(rng: Rng): Outbound[] {
  return Array.from({ length: 100 }, () => ({
    month: rng.pick(['2026-07', '2026-08', '2026-09']),
    warehouse: rng.pick(WAREHOUSES),
    sku: `SKU-${100 + ITEMS.indexOf(rng.pick(ITEMS)) * 7}`,
    units: rng.int(5, 600),
  }));
}

export const pivotShare = defineExercise<{ rows: Outbound[] }>({
  id: 'pivot-share',
  module: 'pivots',
  title: 'Each warehouse’s volume by month',
  replaces: 'A SUMIFS grid plus a second grid of percentages',
  minutes: 5,
  task: () =>
    'Build a PivotTable from the `Outbound` Table with Warehouse as rows, Month as columns and the sum of Units as values. Then show the values as % of row total, so each warehouse’s months add up to 100%.',
  concept: {
    summary:
      'A PivotTable summarizes a table by dragging fields into Rows, Columns and Values. Show Values As changes what each number means, such as its share of the row, without any formulas.',
    syntax: 'Insert › PivotTable, then drag fields: Rows = Warehouse, Columns = Month, Values = Units',
    example: 'Right-click any value › Show Values As › % of Row Total.',
    tip: 'The PivotTable can go on a new sheet; the coach finds it anywhere in the workbook.',
  },
  hints: [
    'Click inside the Outbound Table, then Insert › PivotTable › OK.',
    'In the field list, drag Warehouse to Rows, Month to Columns and Units to Values.',
    'Right-click a number in the PivotTable › Show Values As › % of Row Total.',
  ],
  solution: () => 'Rows: Warehouse · Columns: Month · Values: Sum of Units · Show Values As: % of Row Total',
  guide: (d) => {
    // PivotTables list row labels A to Z and months in order, so the first warehouse is the top row.
    const wh = [...new Set(d.rows.map((r) => r.warehouse))].sort()[0];
    const mine = d.rows.filter((r) => r.warehouse === wh);
    const total = sum(mine.map((r) => r.units));
    const firstMonth = [...new Set(mine.map((r) => r.month))].sort()[0];
    const share = ((100 * sum(mine.filter((r) => r.month === firstMonth).map((r) => r.units))) / total).toFixed(2);
    return [
      {
        do: 'Meet the data. The Table in columns `A` to `D` is named **Outbound**: each row is units of one SKU shipped from one warehouse in one month.',
        why: 'A PivotTable adds up Units for every warehouse and month at once, with no formulas. Each column name becomes a **field**: a piece you drag into place to lay out the summary.',
        show: [
          { label: 'Warehouse column', at: 'Outbound[Warehouse]', note: 'Each warehouse will get its own row in the PivotTable.' },
          { label: 'Month column', at: 'Outbound[Month]', note: 'Each month will get its own column.' },
          { label: 'Units column', at: 'Outbound[Units]', note: 'The numbers the PivotTable adds up.' },
        ],
      },
      clickTableStep('Outbound', `A1:D${d.rows.length + 1}`, 'B5'),
      insertPivotStep('Outbound'),
      dragStep('Warehouse', 'Rows', 'The warehouse names appear down the left of the PivotTable, one row each.'),
      dragStep('Month', 'Columns', 'The months spread across the top, with a **Grand Total** column at the right end.'),
      dragStep(
        'Units',
        'Values',
        `The field reads **Sum of Units**, and each cell totals one warehouse in one month: the SUMIFS grid, built by dragging. ${wh}’s **Grand Total**, at the right end of its row, should read ${total}.`,
        { kind: 'pivot', rows: 'Warehouse', columns: 'Month', valuesField: 'Units', summarizeBy: 'Sum', showAs: 'None', label: 'Warehouse, Month and Sum of Units are in place' },
      ),
      {
        do: 'Right-click any number in the PivotTable, then choose **Show Values As › % of Row Total**.',
        why: `Each number becomes its share of that warehouse’s total, so every row’s **Grand Total** reads 100.00%. ${wh}’s first month, the leftmost number in its row, should read ${share}%.`,
        done: {
          kind: 'inspect',
          inspection: {
            kind: 'pivot',
            rows: 'Warehouse',
            columns: 'Month',
            valuesField: 'Units',
            summarizeBy: 'Sum',
            showAs: 'PercentOfRowTotal',
            label: 'PivotTable layout',
          },
        },
      },
      checkStep('The coach finds your PivotTable on any sheet and reads its layout: Warehouse in Rows, Month in Columns, and Sum of Units in Values, shown as % of Row Total.'),
    ];
  },
  make: (rng) => ({ rows: outbound(rng) }),
  layout: (d) => ({
    blocks: [dataBlock('Outbound', 'A1', OUTBOUND_COLS, d.rows.map((r) => [r.month, r.warehouse, r.sku, r.units]))],
    answer: { kind: 'pivot' },
  }),
  expected: () => [],
  inputs: () => [],
  variants: [],
  inspections: () => [
    {
      kind: 'pivot',
      rows: 'Warehouse',
      columns: 'Month',
      valuesField: 'Units',
      summarizeBy: 'Sum',
      showAs: 'PercentOfRowTotal',
      label: 'PivotTable layout',
    },
  ],
});

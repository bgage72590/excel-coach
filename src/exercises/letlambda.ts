import { GL_ACCOUNTS, ITEMS } from '../engine/data';
import { round, type Rng } from '../engine/rng';
import type { ColumnSpec } from '../engine/types';
import { FMT, cells, dataBlock, defineExercise, tableWrite } from './common';

// ---------- Reorder point with LET ----------

interface StockRow {
  sku: string;
  item: string;
  demand: number;
  lead: number;
  safety: number;
}

const STOCK_COLS: ColumnSpec[] = [
  { header: 'SKU' },
  { header: 'Item' },
  { header: 'Avg daily demand', format: FMT.dec1 },
  { header: 'Lead days', format: FMT.int },
  { header: 'Safety days', format: FMT.int },
];

function stock(rng: Rng): StockRow[] {
  return rng.sample(ITEMS, 20).map((it, i) => ({
    sku: `SKU-${300 + i * 11}`,
    item: it.item,
    demand: rng.float(0.5, 40, 1),
    lead: rng.int(3, 21),
    safety: rng.int(2, 10),
  }));
}

const stockGrid = (rows: StockRow[]) => rows.map((r) => [r.sku, r.item, r.demand, r.lead, r.safety]);
/** ROUNDUP(x, 0), ignoring float noise such as 3.0000000000000004. */
const roundUp = (x: number) => Math.ceil(round(x, 9));

export const letReorder = defineExercise<{ rows: StockRow[] }>({
  id: 'let-reorder',
  module: 'letlambda',
  title: 'Reorder point with LET',
  replaces: 'Long formulas nobody, including you, can read a month later',
  minutes: 5,
  task: () =>
    'In `G2:G21`, calculate each item’s reorder point: average daily demand × (lead days + safety days), rounded up to a whole unit. Use LET to give each input a name.',
  concept: {
    summary:
      'LET names the pieces of a formula, then uses the names in a final calculation. The formula reads like a sentence, and anything used twice is only calculated once.',
    syntax: '=LET(name1, value1, name2, value2, …, calculation)',
    example: '=LET(demand, C2, lead, D2, safety, E2, ROUNDUP(demand*(lead+safety), 0))',
  },
  hints: [
    'Name the three inputs first: demand is C2, lead is D2, safety is E2.',
    'The last argument of LET is the calculation that uses those names.',
    'Round up with ROUNDUP(…, 0).',
  ],
  solution: () => '=LET(demand,C2,lead,D2,safety,E2,ROUNDUP(demand*(lead+safety),0))',
  make: (rng) => ({ rows: stock(rng) }),
  layout: (d) => ({
    blocks: [dataBlock('Stock', 'A1', STOCK_COLS, stockGrid(d.rows)), cells('G1', [['Reorder point']], 'header')],
    answer: { kind: 'cells', range: 'G2:G21', format: FMT.int, consistency: 'all' },
  }),
  expected: (d) => d.rows.map((r) => [roundUp(r.demand * (r.lead + r.safety))]),
  inputs: (d) => [tableWrite('Stock', STOCK_COLS, stockGrid(d.rows))],
  variants: [
    { label: 'demand changes', apply: (d, rng) => ({ rows: d.rows.map((r) => ({ ...r, demand: rng.float(0.5, 40, 1) })) }) },
    { label: 'lead times change', apply: (d, rng) => ({ rows: d.rows.map((r) => ({ ...r, lead: rng.int(3, 21) })) }) },
  ],
  rules: {
    require: [{ pattern: /LET\(/i, label: 'Names the parts with LET', advice: 'Wrap it in LET: =LET(demand, C2, lead, D2, safety, E2, …).' }],
    allowNumbers: [0],
  },
});

// ---------- Build your own VARPCT function ----------

interface BudgetRow {
  account: string;
  budget: number;
  actual: number;
}

const BUDGET_COLS: ColumnSpec[] = [{ header: 'Account' }, { header: 'Budget', format: FMT.currency }, { header: 'Actual', format: FMT.currency }];

function budget(rng: Rng): BudgetRow[] {
  return rng.sample(GL_ACCOUNTS, 12).map((a) => {
    const b = rng.float(2_000, 90_000, 0);
    return { account: a.name, budget: b, actual: round(b * rng.float(0.7, 1.35, 3), 2) };
  });
}

const budgetGrid = (rows: BudgetRow[]) => rows.map((r) => [r.account, r.budget, r.actual]);

export const lambdaVarpct = defineExercise<{ rows: BudgetRow[] }>({
  id: 'lambda-varpct',
  module: 'letlambda',
  title: 'Build your own VARPCT function',
  replaces: 'Pasting the same variance formula into every report, slightly differently each time',
  minutes: 6,
  task: () =>
    'Create a function named `VARPCT` that takes an actual and a budget and returns (actual − budget) ÷ budget. Then use it in `E2:E13` to show each account’s variance.',
  concept: {
    summary:
      'LAMBDA turns a formula into a reusable function. Save it under a name in Name Manager and you can call it from any cell in the workbook, like a built-in function.',
    syntax: 'Name: VARPCT    Refers to: =LAMBDA(actual, budget, (actual-budget)/budget)',
    example: '=VARPCT(C2, B2) returns 0.12 when actual is 12% over budget.',
    tip: 'Test a LAMBDA in a cell first by calling it right away: =LAMBDA(a, b, (a-b)/b)(C2, B2).',
  },
  hints: [
    'Open {nameManager} and create a name called VARPCT.',
    'In “Refers to”, enter =LAMBDA(actual, budget, (actual-budget)/budget).',
    'Back on the sheet, type =VARPCT(C2, B2) in E2 and fill down.',
  ],
  solution: () => 'Name VARPCT = LAMBDA(actual,budget,(actual-budget)/budget)    E2: =VARPCT(C2,B2)',
  make: (rng) => ({ rows: budget(rng) }),
  layout: (d) => ({
    // Leave column D empty: a label written right beside a Table makes Excel grow the Table over it.
    blocks: [dataBlock('Budget', 'A1', BUDGET_COLS, budgetGrid(d.rows)), cells('E1', [['Variance %']], 'header')],
    answer: { kind: 'cells', range: 'E2:E13', format: FMT.pct, consistency: 'all' },
  }),
  expected: (d) => d.rows.map((r) => [(r.actual - r.budget) / r.budget]),
  inputs: (d) => [tableWrite('Budget', BUDGET_COLS, budgetGrid(d.rows))],
  variants: [
    { label: 'actuals change', apply: (d, rng) => ({ rows: d.rows.map((r) => ({ ...r, actual: round(r.budget * rng.float(0.6, 1.5, 3), 2) })) }) },
    { label: 'budgets change', apply: (d, rng) => ({ rows: d.rows.map((r) => ({ ...r, budget: rng.float(2_000, 90_000, 0) })) }) },
  ],
  rules: {
    require: [{ pattern: /VARPCT\(/i, label: 'Uses your VARPCT function', advice: 'Call it like =VARPCT(C2, B2).' }],
    allowNumbers: [],
  },
  inspections: () => [{ kind: 'lambdaName', name: 'VARPCT', params: 2, label: 'VARPCT is defined as a LAMBDA with two inputs' }],
  ownsNames: ['VARPCT'],
});

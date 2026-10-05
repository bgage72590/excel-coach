import { GL_ACCOUNTS, ITEMS } from '../engine/data';
import { round, type Rng } from '../engine/rng';
import type { ColumnSpec } from '../engine/types';
import { FMT, cells, dataBlock, defineExercise, tableWrite } from './common';
import { checkStep, fillStep, money, part, raw, typeStep } from './guides';

/** 12.3%, the way an FMT.pct cell shows. */
const pct = (n: number) => `${(n * 100).toFixed(1)}%`;
/** Up to one decimal, the way the sheet shows a number: 1,234.5 or 96. */
const num = (n: number) => n.toLocaleString('en-US', { maximumFractionDigits: 1 });

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
  guide: (d) => {
    const r = d.rows[0];
    const exact = round(r.demand * (r.lead + r.safety), 1);
    const point = roundUp(r.demand * (r.lead + r.safety));
    const demand = r.demand.toFixed(1);
    return [
      {
        do: 'Meet the data. The blue block in columns `A` to `E` is a Table named **Stock**. Each row is one item, with the three inputs a reorder point needs.',
        why: 'A reorder point is how many units to have left when you place an order: enough to cover demand while the order arrives (lead days), plus a cushion (safety days).',
        show: [
          { label: 'Avg daily demand', at: 'Stock[Avg daily demand]', note: 'Column `C`: units sold on an average day.' },
          { label: 'Lead days', at: 'Stock[Lead days]', note: 'Column `D`: days between ordering and the stock arriving.' },
          { label: 'Safety days', at: 'Stock[Safety days]', note: 'Column `E`: extra days of stock kept as a cushion.' },
        ],
      },
      {
        do: `See what \`G2\` should show for ${r.item}.`,
        why: `${demand} a day × (${r.lead} lead days + ${r.safety} safety days) = ${num(exact)}. ${
          exact === point ? 'That’s already a whole unit, so rounding up leaves it as it is.' : `Rounded up to a whole unit, that’s ${point.toLocaleString('en-US')}.`
        }`,
        show: [
          {
            label: `Select the inputs for ${r.item}`,
            at: 'C2:E2',
            note: `Demand ${demand}, lead ${r.lead} days, safety ${r.safety} days. \`G2\` should show ${point.toLocaleString('en-US')}.`,
          },
        ],
      },
      typeStep({
        cell: 'G2',
        formula: [
          part('=LET(', 'Gives names to values, then works out a final answer using those names.'),
          part('demand', 'The first name you make up. Each name is followed by the value it stands for…'),
          raw(', '),
          part('C2', `…here the Avg daily demand on this row, in \`C2\` (${demand}).`, 'C2'),
          raw(', '),
          part('lead', 'The second name…'),
          raw(', '),
          part('D2', `…for the Lead days in \`D2\` (${r.lead}).`, 'D2'),
          raw(', '),
          part('safety', 'The third name…'),
          raw(', '),
          part('E2', `…for the Safety days in \`E2\` (${r.safety}).`, 'E2'),
          raw(', '),
          part('ROUNDUP(', 'The last argument is the answer, written with the names. ROUNDUP rounds up…'),
          part('demand*(lead+safety)', `…demand × (lead + safety), which is ${demand} × (${r.lead} + ${r.safety}) = ${num(exact)}…`),
          raw(', '),
          part('0', `…to 0 decimal places, a whole unit: ${point.toLocaleString('en-US')}.`),
          part('))', 'Closes ROUNDUP, then LET.'),
        ],
        why: 'Type the addresses `C2`, `D2` and `E2`. If you click the cells instead, Excel may write a longer Table name such as `Stock[@[Lead days]]`; that works too.',
      }),
      fillStep({
        from: 'G2',
        range: 'G2:G21',
        direction: 'down',
        why: 'Fill Down copies `G2`’s formula into the cells below. Excel moves `C2`, `D2` and `E2` down a row each time, so every item names its own inputs. The names stay the same.',
      }),
      checkStep('The coach changes demand and lead times behind the scenes, then puts them back. Your formula reads the cells, so every reorder point updates.'),
    ];
  },
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
  guide: (d) => {
    const r = d.rows[0];
    const v = (r.actual - r.budget) / r.budget;
    return [
      {
        do: 'Meet the data. The blue block in columns `A` to `C` is a Table named **Budget**: one row per account, with its Budget (column `B`) and Actual (column `C`).',
        why: 'You’ll build your own function, `VARPCT`, that works out how far actual is from budget as a percent: (actual − budget) ÷ budget. Then you’ll use it in cells like a built-in function.',
        show: [
          { label: 'Budget column', at: 'Budget[Budget]', note: 'Column `B`: what each account was meant to spend.' },
          { label: 'Actual column', at: 'Budget[Actual]', note: 'Column `C`: what it spent.' },
        ],
      },
      {
        do: `See what \`E2\` should show for ${r.account}.`,
        why: `(${money(r.actual)} − ${money(r.budget)}) ÷ ${money(r.budget)} = ${pct(v)}. ${v >= 0 ? 'Positive means over budget.' : 'Negative means under budget.'}`,
        show: [{ label: `Select the ${r.account} numbers`, at: 'B2:C2', note: `Budget ${money(r.budget)}, actual ${money(r.actual)}. \`E2\` should show ${pct(v)}.` }],
      },
      {
        do: 'Open **{nameManager}**. If a list of names opens rather than a form with a **Name** box, click **New…**.',
        why: 'A name can stand for a formula, not only a cell. Save a LAMBDA under a name and it becomes a function you can call from any cell in the workbook.',
      },
      {
        do: 'In the **Name** box, type `VARPCT`.',
        why: 'This is what you’ll type in cells to call your function, the way you type `SUM`. A name can’t contain spaces.',
      },
      {
        do: 'Click in the **Refers to** box, delete everything in it, and type this. Then click **OK**.',
        formula: [
          part('=LAMBDA(', 'Makes a function. Every argument but the last names an input; the last is the calculation.'),
          part('actual', 'Input 1, named `actual`. When you call the function, the first cell you give it fills this in.'),
          raw(', '),
          part('budget', 'Input 2, named `budget`: the second cell you give it.'),
          raw(', '),
          part('(actual-budget)/budget', 'The calculation, written with the input names: the gap between actual and budget, divided by budget.'),
          raw(')'),
        ],
        why: 'The box starts out holding a cell address; replace all of it. If the arrow keys start picking cells on the sheet, click in the box to move the cursor instead. If Name Manager is still open after **OK**, click **Close**.',
        done: { kind: 'inspect', inspection: { kind: 'lambdaName', name: 'VARPCT', params: 2, label: 'VARPCT is defined as a LAMBDA with two inputs' } },
      },
      typeStep({
        cell: 'E2',
        formula: [
          part('=VARPCT(', 'Your new function. Excel lists it as you type, like a built-in one.'),
          part('C2', `Input 1, \`actual\`: the Actual in \`C2\` (${money(r.actual)}).`, 'C2'),
          raw(', '),
          part('B2', `Input 2, \`budget\`: the Budget in \`B2\` (${money(r.budget)}).`, 'B2'),
          raw(')'),
        ],
        why: `Order matters: actual first, then budget, the same order as in the LAMBDA. \`E2\` should show ${pct(v)}. Type the addresses; if you click the cells, Excel may write longer Table names such as \`Budget[@Actual]\`, which work too.`,
      }),
      fillStep({
        from: 'E2',
        range: 'E2:E13',
        direction: 'down',
        why: 'Fill Down copies `E2`’s formula into the cells below. `C2` and `B2` move down a row each time, so every account gets its own variance.',
      }),
      checkStep('The coach changes the actuals and the budgets behind the scenes, then puts them back. Your function recalculates with them, like any built-in one.'),
    ];
  },
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

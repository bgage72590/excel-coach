import { WAREHOUSES } from '../engine/data';
import { round, type Rng } from '../engine/rng';
import type { CellCheck, Exercise } from '../engine/types';
import { FMT, cells, column, defineExercise, rangeWrite } from './common';
import { checkStep, money } from './guides';

/** "$1,234.50" for solution text. */
function usd(n: number): string {
  const s = Math.abs(n).toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  return `${n < 0 ? '-' : ''}$${s}`;
}

/** 38% for 0.38: whole percents, for walkthrough copy. */
const pct0 = (n: number) => `${Math.round(n * 100)}%`;

// ---------- Find the price for a target margin (Goal Seek) ----------

const GS_ITEMS = ['Packing tape', 'Floor tape', 'Bin divider', 'Shelf bracket', 'Safety glasses', 'Work gloves'] as const;
const GS_TARGETS = [0.3, 0.32, 0.34, 0.35, 0.36, 0.38, 0.4, 0.42] as const;

export interface GoalSeekData {
  item: string;
  units: number;
  cost: number;
  /** Starting price, below the one that hits the target. */
  price: number;
  target: number;
}

/** The exact price where Margin % = (P·U − C·U) / (P·U) = 1 − C/P equals the target. */
export const targetPrice = (d: GoalSeekData) => d.cost / (1 - d.target);

/** Margin % the sheet shows for a given price (the prewritten formulas, in Excel's order). */
export function marginAt(d: GoalSeekData, price: number): number {
  const revenue = price * d.units;
  const cogs = d.cost * d.units;
  return (revenue - cogs) / revenue;
}

/** How much Margin % moves per $1 of price at the solution: (1 − M)² / C. */
export const marginSlope = (d: GoalSeekData) => (1 - d.target) ** 2 / d.cost;

/** Goal Seek's default stop: the result is within 0.001 of the target (Maximum Change). */
export const GOAL_SEEK_STOP = 0.001;
/**
 * Margin tolerance, with headroom over Goal Seek's stop so a correct run never lands on the edge.
 * The price tolerance is wide enough that any margin within GS_MARGIN_TOL also passes the price
 * check (a test proves it), yet far tighter than the gap to the starting price.
 */
export const GS_MARGIN_TOL = 0.0015;
export const GS_PRICE_TOL = 0.02;

export const GS_FORMULAS = { revenue: '=B5*B3', cogs: '=B4*B3', margin: '=B8-B9', marginPct: '=B10/B8' } as const;

export const goalSeekPrice = defineExercise<GoalSeekData>({
  id: 'goal-seek-price',
  module: 'whatif',
  title: 'Find the price for a target margin',
  replaces: 'Trial-and-error price tweaks until the margin hits target',
  minutes: 4,
  task: () =>
    'Find the price per unit in `B5` that brings Margin % in `B11` to the target in `B13`. Use Data › What-If Analysis › Goal Seek instead of guessing. Graded: `B5` holds the typed price Goal Seek finds (within two cents of the exact answer), and `B11` still has its formula and shows the target margin.',
  concept: {
    summary:
      'Goal Seek works backward. You name the formula cell, the result you want and the one input it may change, and Excel tries values until the formula hits the target. It types the answer into the input cell, so the model keeps working afterward.',
    syntax: 'Data › What-If Analysis › Goal Seek: Set cell (the formula), To value (the target), By changing cell (the input)',
    example: 'Set cell B11, To value 0.38, By changing cell B5 finds the price that gives a 38% margin.',
    tip: 'In the Goal Seek Status box, OK keeps the new value and Cancel puts the old one back.',
  },
  hints: [
    'Goal Seek lives under Data › What-If Analysis. It needs the formula cell to aim at, the value to hit and the input it may change.',
    'Set cell is the Margin % cell, B11. By changing cell is the price, B5.',
    'In To value, type the target from B13 as a decimal (38% is 0.38), click OK, then OK again in the Goal Seek Status box to keep the price.',
  ],
  solution: (d) =>
    `Data › What-If Analysis › Goal Seek: Set cell B11, To value ${d.target}, By changing cell B5. The price comes out at about ${usd(targetPrice(d))}.`,
  guide: (d) => {
    const now = `${(marginAt(d, d.price) * 100).toFixed(1)}%`;
    return [
      {
        do: 'Meet the model. The price in `B5` feeds the formulas in `B8:B11`, and `B11` works out Margin %.',
        why: `Right now the price is ${money(d.price)} and Margin % shows ${now}. The target in \`B13\` is ${pct0(d.target)}. A higher price means a higher margin, but which price lands on ${pct0(d.target)} exactly? Goal Seek works backward from the target to find it.`,
        show: [
          { label: 'Price', at: 'B5', note: 'The input Goal Seek will change.' },
          { label: 'Margin %', at: 'B11', note: '`B11` holds `=B10/B8`: gross margin divided by revenue. Goal Seek needs a formula cell like this to aim at.' },
          { label: 'Target', at: 'B13', note: `The margin you want: ${pct0(d.target)}.` },
        ],
      },
      {
        do: 'Click `B11`, the Margin % cell.',
        why: 'Goal Seek starts with the selected cell in its first box, which saves a step.',
        done: { kind: 'select', range: 'B11' },
      },
      {
        do: 'Click **Data › What-If Analysis › Goal Seek**.',
        why: 'A small box opens with three fields: **Set cell**, **To value** and **By changing cell**.',
      },
      {
        do: 'Check that **Set cell** shows `B11` (Excel may write it as `$B$11`). If it shows another cell, click in the box, then click `B11`.',
        why: 'Set cell is the formula you want to land on a number: Margin %. It has to be a formula, not a typed number.',
      },
      {
        do: `In **To value**, type \`${d.target}\`.`,
        why: `That’s the ${pct0(d.target)} target from \`B13\`, written as a decimal. This box takes a typed number, not a cell.`,
      },
      {
        do: 'Click in **By changing cell**, then click `B5`.',
        why: 'This is the input Goal Seek may change: the price. It tries one price after another until `B11` shows the target.',
      },
      {
        do: 'Click **OK**.',
        why: `After a moment, a **Goal Seek Status** box says it found a solution. \`B5\` now shows about ${money(targetPrice(d))}, and \`B11\` shows ${pct0(d.target)}.`,
      },
      {
        do: 'In the **Goal Seek Status** box, click **OK** to keep the new price.',
        why: '**Cancel** would put the old price back. OK leaves the price typed into `B5`, so the model keeps working afterward.',
      },
      checkStep('The coach checks that `B5` holds a typed price within two cents of the exact answer, and that `B11` still has its formula and shows the target margin.'),
    ];
  },
  make: (rng) => {
    const item = rng.pick(GS_ITEMS);
    const target = rng.pick(GS_TARGETS);
    const cost = rng.float(1.8, 3, 2);
    const price = round(cost * rng.float(1.18, 1.32, 3), 2);
    const units = rng.int(30, 120) * 100;
    return { item, units, cost, price, target };
  },
  layout: (d) => ({
    blocks: [
      cells('A1', [['Input', 'Value']], 'header'),
      cells('A2', column(['Item', 'Units per month', 'Unit cost', 'Price per unit']), 'label'),
      cells('B2', [[d.item]], 'input'),
      cells('B3', [[d.units]], 'input', FMT.int),
      cells('B4', [[d.cost]], 'input', FMT.currency),
      cells('B5', [[d.price]], 'input', FMT.currency),
      cells('A7', [['Result', 'Value']], 'header'),
      cells('A8', column(['Revenue', 'Cost of goods', 'Gross margin', 'Margin %']), 'label'),
      cells('B8', [[GS_FORMULAS.revenue], [GS_FORMULAS.cogs], [GS_FORMULAS.margin]], 'formula', FMT.currency),
      cells('B11', [[GS_FORMULAS.marginPct]], 'formula', FMT.pct),
      cells('A13', [['Target margin %']], 'label'),
      cells('B13', [[d.target]], 'input', FMT.pct),
    ],
    answer: {
      kind: 'cellChecks',
      checks: [
        {
          cell: 'B5',
          label: 'B5 holds the price Goal Seek found',
          answer: true,
          holds: 'number',
          value: round(targetPrice(d), 6),
          tolerance: GS_PRICE_TOL,
          advice: 'Run Goal Seek with Set cell B11 and By changing cell B5.',
        },
        {
          cell: 'B11',
          label: 'Margin % in B11 still uses its formula and hits the target',
          formula: GS_FORMULAS.marginPct,
          value: d.target,
          tolerance: GS_MARGIN_TOL,
        },
      ],
    },
  }),
  expected: () => [],
  inputs: () => [],
  variants: [],
});

// ---------- Build a two-way data table ----------

export interface ProfitData {
  price: number;
  cost: number;
  units: number;
  fixed: number;
  /** Price options across C9:H9. */
  prices: number[];
  /** Volume options down B10:B16. */
  volumes: number[];
}

export const DT_PROFIT = '=(B2-B3)*B4-B5';

/** Profit for one price and volume, in the same order Excel evaluates =(B2-B3)*B4-B5. */
export const profit = (d: ProfitData, price: number, units: number) => (price - d.cost) * units - d.fixed;

const profitGrid = (d: ProfitData) => d.volumes.map((v) => d.prices.map((p) => profit(d, p, v)));

const dataTableExplain =
  'The grid didn’t follow the model, so it isn’t a live data table. Select B9:H16 and use Data › What-If Analysis › Data Table (typed or pasted numbers never update). If it is a data table, check that B9 still points at B6 and Formulas › Calculation Options is Automatic.';

export const dataTableTwoWay = defineExercise<ProfitData>({
  id: 'data-table-two-way',
  module: 'whatif',
  title: 'Build a two-way data table',
  replaces: 'Copying the model once per scenario and editing each copy',
  minutes: 6,
  task: () =>
    'Fill `C10:H16` with the monthly profit for every price in `C9:H9` and every volume in `B10:B16`. `B9` already points at Profit in `B6`. Use Data › What-If Analysis › Data Table so the grid stays live. Graded: every cell in `C10:H16` holds the data table’s result, and the grid still matches when Unit cost or Fixed cost changes.',
  concept: {
    summary:
      'A data table reruns one model for many inputs at once. Put the options for one input across the top, the options for another down the side, and point the corner cell at the result. Excel fills the grid and keeps it up to date.',
    syntax: 'Select the grid including the corner › Data › What-If Analysis › Data Table › Row input cell, Column input cell',
    example: 'Row input cell B2, because prices run across row 9. Column input cell B4, because units run down column B.',
    tip: 'Excel shows {=TABLE(B2,B4)} in every result cell. You can’t edit one cell of a data table; to redo it, select all of C10:H16 and delete.',
  },
  hints: [
    'Select the whole grid, B9:H16, including the corner cell B9 and both sets of options.',
    'Open Data › What-If Analysis › Data Table. Prices run across the top, so they stand in for the price cell; units run down the side, so they stand in for the units cell.',
    'Row input cell: B2. Column input cell: B4. Click OK and the grid fills in.',
  ],
  solution: () => 'Select B9:H16 › Data › What-If Analysis › Data Table › Row input cell B2, Column input cell B4. Every result cell shows =TABLE(B2,B4).',
  guide: (d) => {
    const units = (n: number) => n.toLocaleString('en-US');
    const firstCell = profit(d, d.prices[0], d.volumes[0]);
    return [
      {
        do: 'Meet the model. The inputs are in `B2:B5`, and `B6` works out monthly Profit from them.',
        show: [
          { label: 'Inputs', at: 'B2:B5', note: 'Price per unit, Unit cost, Units per month and Fixed cost per month.' },
          {
            label: 'Profit',
            at: 'B6',
            note: `\`B6\` holds \`=(B2-B3)*B4-B5\`: (price − unit cost) × units − fixed cost. Right now it’s ${money(profit(d, d.price, d.units))}.`,
          },
        ],
      },
      {
        do: 'Meet the grid. Prices run across `C9:H9`, volumes run down `B10:B16`, and the corner cell `B9` points at Profit.',
        why: 'A **data table** reruns the model once for every cell of the grid. For each cell it puts that column’s price into `B2` and that row’s volume into `B4`, then records what `B6` would show. The corner tells it which result to record.',
        show: [
          { label: 'Prices', at: 'C9:H9', note: 'Each column of the grid tries one of these prices.' },
          { label: 'Volumes', at: 'B10:B16', note: 'Each row of the grid tries one of these volumes.' },
          { label: 'Corner', at: 'B9', note: '`B9` holds `=B6`, so the data table records Profit.' },
        ],
      },
      {
        do: `See what \`C10\` will show: profit at a price of ${money(d.prices[0])} and ${units(d.volumes[0])} units a month.`,
        why: `(${money(d.prices[0])} − ${money(d.cost)}) × ${units(d.volumes[0])} − ${money(d.fixed)} = ${money(firstCell)}. And \`E13\` will match \`B6\`, because that’s the model’s own price and volume.`,
        show: [{ label: 'Select its price and volume', at: 'C9,B10', note: '`C9` is the price and `B10` the volume that `C10` stands for.' }],
      },
      {
        do: 'Select `B9:H16`: click `B9`, then Shift-click `H16`.',
        why: 'Include the corner and both sets of options. The empty cells `C10:H16` are where the results go.',
      },
      {
        do: 'Click **Data › What-If Analysis › Data Table**.',
        why: 'A small box opens with two fields: **Row input cell** and **Column input cell**.',
      },
      {
        do: 'Click in **Row input cell**, then click `B2`.',
        why: 'The prices sit in a row across the top, so they stand in for the price cell, `B2`.',
      },
      {
        do: 'Click in **Column input cell**, then click `B4`.',
        why: 'The volumes sit in a column down the side, so they stand in for the units cell, `B4`.',
      },
      {
        do: 'Click **OK**.',
        why: `The grid fills in, and \`C10\` shows ${money(firstCell)}. Click any result: the formula bar shows a \`TABLE\` formula pointing at \`B2\` and \`B4\`. You can’t edit one cell of a data table; to redo it, select all of \`C10:H16\` and delete.`,
        done: { kind: 'answer' },
      },
      checkStep('The coach changes Unit cost and Fixed cost behind the scenes. A live data table follows the model; typed numbers wouldn’t.'),
    ];
  },
  make: (rng) => {
    const price = rng.int(30, 60);
    const step = rng.pick([1, 2, 2.5]);
    const units = rng.int(8, 16) * 100;
    const vstep = rng.pick([100, 150, 200]);
    const cost = round(price * rng.float(0.5, 0.68, 3), 2);
    const fixed = Math.round(((price - cost) * units * rng.float(0.55, 0.95, 3)) / 500) * 500;
    return {
      price,
      cost,
      units,
      fixed,
      prices: [-2, -1, 0, 1, 2, 3].map((k) => price + k * step),
      volumes: [-3, -2, -1, 0, 1, 2, 3].map((k) => units + k * vstep),
    };
  },
  layout: (d) => ({
    blocks: [
      cells('A1', [['Input', 'Value']], 'header'),
      cells('A2', column(['Price per unit', 'Unit cost', 'Units per month', 'Fixed cost per month', 'Profit']), 'label'),
      cells('B2', [[d.price], [d.cost]], 'input', FMT.currency),
      cells('B4', [[d.units]], 'input', FMT.int),
      cells('B5', [[d.fixed]], 'input', FMT.currency),
      cells('B6', [[DT_PROFIT]], 'formula', FMT.currency),
      cells('C8', [['Price per unit']], 'label'),
      cells('B9', [['=B6']], 'formula', FMT.currency),
      cells('C9', [d.prices], 'input', FMT.currency),
      cells('A10', [['Units per month']], 'label'),
      cells('B10', column(d.volumes), 'input', FMT.int),
    ],
    answer: { kind: 'cells', range: 'C10:H16', format: FMT.currency, consistency: 'none', liveValues: true },
  }),
  expected: profitGrid,
  inputs: (d) => [rangeWrite('B3', [[d.cost]]), rangeWrite('B5', [[d.fixed]])],
  variants: [
    {
      label: 'the unit cost changes',
      explain: dataTableExplain,
      apply: (d, rng) => {
        let cost = round(d.cost * rng.float(0.85, 1.15, 3), 2);
        if (cost === d.cost) cost = round(d.cost + 0.75, 2);
        return { ...d, cost };
      },
    },
    {
      label: 'the fixed cost changes',
      explain: dataTableExplain,
      apply: (d, rng) => ({ ...d, fixed: d.fixed + rng.pick([-3000, -2000, -1500, 1500, 2000, 3000]) }),
    },
  ],
  // No formula rule: Office.js reports data-table cells as values. The variants prove the grid is live.
});

// ---------- Cheapest shipping plan with Solver ----------

export interface ShipPlanData {
  warehouses: string[];
  stores: string[];
  /** Cost per pallet, warehouse × store. */
  cost: number[][];
  supply: number[];
  demand: number[];
  /** The cheapest whole-pallet plan, warehouse × store. */
  plan: number[][];
  best: number;
}

export interface ShippingSolution {
  plan: number[][];
  best: number;
  /** How many whole-pallet plans reach the minimum. */
  ties: number;
}

/**
 * Exact integer optimum for 2 warehouses × 3 stores. Each store's warehouse-1 amount fixes its
 * warehouse-2 amount (demand − x), so enumerating warehouse 1's three amounts covers every plan.
 */
export function solveShipping(cost: number[][], supply: number[], demand: number[]): ShippingSolution {
  const [d0, d1, d2] = demand;
  let best = Infinity;
  let ties = 0;
  let plan: number[][] = [];
  for (let a = 0; a <= d0; a++) {
    for (let b = 0; b <= d1; b++) {
      if (a + b > supply[0]) break;
      for (let c = 0; c <= d2; c++) {
        if (a + b + c > supply[0]) break;
        if (d0 - a + (d1 - b) + (d2 - c) > supply[1]) continue;
        const total =
          cost[0][0] * a + cost[0][1] * b + cost[0][2] * c + cost[1][0] * (d0 - a) + cost[1][1] * (d1 - b) + cost[1][2] * (d2 - c);
        if (total < best) {
          best = total;
          ties = 1;
          plan = [
            [a, b, c],
            [d0 - a, d1 - b, d2 - c],
          ];
        } else if (total === best) {
          ties++;
        }
      }
    }
  }
  return { plan, best, ties };
}

/** True when sending every store's pallets from its cheaper warehouse would break a supply limit. */
export function greedyBreaksSupply(cost: number[][], supply: number[], demand: number[]): boolean {
  const fromFirst = demand.reduce((s, dj, j) => s + (cost[0][j] < cost[1][j] ? dj : 0), 0);
  const total = demand.reduce((s, dj) => s + dj, 0);
  return fromFirst > supply[0] || total - fromFirst > supply[1];
}

function shippingProblem(rng: Rng): ShipPlanData {
  const warehouses = rng.sample(WAREHOUSES, 2);
  const stores = rng
    .sample(Array.from({ length: 80 }, (_, i) => 10 + i), 3)
    .sort((a, b) => a - b)
    .map((n) => `Store ${n}`);
  let last: ShipPlanData | undefined;
  for (let attempt = 0; attempt < 40; attempt++) {
    const demand = [0, 1, 2].map(() => rng.int(4, 12) * 10);
    const total = demand.reduce((s, x) => s + x, 0);
    const capacity = total + rng.int(1, 4) * 10;
    const first = Math.round((capacity * rng.float(0.35, 0.65, 3)) / 10) * 10;
    const supply = [first, capacity - first];
    const cost = [0, 1].map(() => [0, 1, 2].map(() => rng.int(35, 140)));
    const sol = solveShipping(cost, supply, demand);
    last = { warehouses, stores, cost, supply, demand, plan: sol.plan, best: sol.best };
    const distinct = cost[0].every((c, j) => c !== cost[1][j]);
    if (sol.ties === 1 && distinct && supply.every((s) => s > 0 && s < total) && greedyBreaksSupply(cost, supply, demand)) return last;
  }
  return last!;
}

/** Decision cells B6:D7, one per warehouse (row) and store (column). */
const DECISION_COLS = ['B', 'C', 'D'];
export const SHIP_FORMULAS = {
  shipped: ['=SUM(B6:D6)', '=SUM(B7:D7)'],
  received: ['=SUM(B6:B7)', '=SUM(C6:C7)', '=SUM(D6:D7)'],
  total: '=SUMPRODUCT(B2:D3,B6:D7)',
} as const;

function shippingChecks(d: ShipPlanData): CellCheck[] {
  const decisions: CellCheck[] = d.warehouses.flatMap((w, i) =>
    d.stores.map((s, j) => ({
      cell: `${DECISION_COLS[j]}${6 + i}`,
      label: `${w} to ${s}: a whole number of pallets, 0 or more`,
      answer: true,
      holds: 'number' as const,
      min: 0,
      integer: true,
      advice: 'Tick Make Unconstrained Variables Non-Negative and solve with Simplex LP.',
    })),
  );
  const received: CellCheck[] = d.stores.map((s, j) => ({
    cell: `${DECISION_COLS[j]}8`,
    label: `${s} receives exactly ${d.demand[j]} pallets`,
    formula: SHIP_FORMULAS.received[j],
    value: d.demand[j],
    tolerance: 0.001,
    advice: 'Add the constraint B8:D8 = B9:D9.',
  }));
  const shipped: CellCheck[] = d.warehouses.map((w, i) => ({
    cell: `E${6 + i}`,
    label: `${w} ships no more than its ${d.supply[i]} pallets`,
    formula: SHIP_FORMULAS.shipped[i],
    max: d.supply[i],
    advice: 'Add the constraint E6:E7 <= F6:F7.',
  }));
  const total: CellCheck = {
    cell: 'B11',
    label: 'Total cost in B11 is the lowest possible',
    formula: SHIP_FORMULAS.total,
    value: d.best,
    tolerance: 0.01,
    advice: 'Set the objective to Min and solve with Simplex LP.',
  };
  return [...decisions, ...received, ...shipped, total];
}

export const solverShipping = defineExercise<ShipPlanData>({
  id: 'solver-shipping',
  module: 'whatif',
  title: 'Cheapest shipping plan with Solver',
  replaces: 'Shuffling pallet counts by hand until the freight bill looks low',
  minutes: 8,
  task: () =>
    'Plan how many pallets each warehouse sends to each store by changing only the yellow cells `B6:D7`. Every store must receive exactly its demand (row `9`), no warehouse may ship more than its supply (column `F`), and Total cost in `B11` should be as low as possible. Use Solver. Graded: whole, non-negative pallet counts in `B6:D7`; the formulas in `E6:E7`, `B8:D8` and `B11` unchanged; every store receives its demand; no warehouse goes over supply; and `B11` matches the lowest possible cost.',
  concept: {
    summary:
      'Solver finds the best values for several input cells at once, within limits you set. You give it a cell to minimize or maximize, the cells it may change, and constraints such as “received = demand”.',
    syntax: 'Data › Solver: Set Objective, To Min or Max, By Changing Variable Cells, Subject to the Constraints, Select a Solving Method',
    example: 'Set Objective $B$11 · To Min · By Changing $B$6:$D$7 · Constraints $B$8:$D$8 = $B$9:$D$9 and $E$6:$E$7 <= $F$6:$F$7 · Simplex LP',
    tip: 'Simplex LP fits any model that only adds up inputs multiplied by fixed numbers, like this one. It finds the true minimum rather than a good guess. On a Mac, Solver can show “No cells were found” a few times while it sets up; click OK each time and it carries on.',
  },
  hints: [
    'Turn Solver on once: Tools › Excel Add-ins › Solver Add-in on a Mac (File › Options › Add-ins › Go on Windows). It then appears as Data › Solver.',
    'Set Objective B11, To Min, By Changing Variable Cells B6:D7. Tick Make Unconstrained Variables Non-Negative and choose Simplex LP.',
    'Add two constraints: B8:D8 = B9:D9 (each store gets its demand) and E6:E7 <= F6:F7 (no warehouse ships more than it has). Click Solve, then keep the Solver solution.',
  ],
  solution: (d) => {
    const routes = d.warehouses
      .flatMap((w, i) => d.stores.map((s, j) => (d.plan[i][j] ? `${w} to ${s}: ${d.plan[i][j]}` : '')))
      .filter(Boolean)
      .join(', ');
    return `Data › Solver: Set Objective $B$11, To Min, By Changing $B$6:$D$7. Constraints: $B$8:$D$8 = $B$9:$D$9 and $E$6:$E$7 <= $F$6:$F$7. Make Unconstrained Variables Non-Negative, Simplex LP. Lowest cost ${usd(d.best)} (${routes}).`;
  },
  guide: (d) => {
    // Why trial and error is hard: each store's cheaper warehouse can't cover everything.
    const fromFirst = d.demand.reduce((s, dj, j) => s + (d.cost[0][j] < d.cost[1][j] ? dj : 0), 0);
    const totalDemand = d.demand.reduce((s, dj) => s + dj, 0);
    const over = fromFirst > d.supply[0] ? 0 : 1;
    const need = over === 0 ? fromFirst : totalDemand - fromFirst;
    const squeeze = greedyBreaksSupply(d.cost, d.supply, d.demand)
      ? ` Sending every store its pallets from its cheaper warehouse would need ${need} pallets from ${d.warehouses[over]}, which only has ${d.supply[over]}. That’s why guessing is slow.`
      : '';
    return [
      {
        do: 'Meet the model. Each yellow cell in `B6:D7` is how many pallets one warehouse sends to one store. They all start at 0.',
        why: 'Solver will change those six cells, and only those, to make Total cost in `B11` as low as it can.',
        show: [
          { label: 'Cost per pallet', at: 'B2:D3', note: 'What one pallet costs on each route: warehouse down the side, store across the top.' },
          { label: 'Pallets to ship', at: 'B6:D7', note: 'The cells Solver may change.' },
          { label: 'Total cost', at: 'B11', note: '`B11` holds `=SUMPRODUCT(B2:D3,B6:D7)`: each route’s pallets times its cost, added up.' },
        ],
      },
      {
        do: 'Meet the limits. Received in row `8` must equal Demand in row `9`, and Shipped in column `E` can’t go over Supply in column `F`.',
        why: `Limits like these are called **constraints**.${squeeze}`,
        show: [
          { label: 'Received and Demand', at: 'B8:D9', note: `${d.stores.map((s, j) => `${s} needs ${d.demand[j]}`).join(', ')} pallets.` },
          { label: 'Shipped and Supply', at: 'E6:F7', note: `${d.warehouses[0]} has ${d.supply[0]} pallets and ${d.warehouses[1]} has ${d.supply[1]}.` },
        ],
      },
      {
        do: 'If **Solver** isn’t on the **Data** tab yet, turn it on once: **Tools › Excel Add-ins** on a Mac (**File › Options › Add-ins › Go** on Windows), tick **Solver Add-in** and click **OK**.',
        why: 'Solver comes with Excel but starts switched off. Once it’s on, it stays on.',
      },
      {
        do: 'Click **Data › Solver**.',
        why: 'The **Solver Parameters** box opens. On a Mac, Solver can show “No cells were found” a few times while it sets up; click **OK** each time and it carries on.',
      },
      {
        do: 'Click in **Set Objective**, then click `B11`.',
        why: 'The **objective** is the cell Solver works on: Total cost. Solver writes it as `$B$11`.',
      },
      {
        do: 'Under **To**, choose **Min**.',
        why: 'You want the lowest total cost, not the highest.',
      },
      {
        do: 'Click in **By Changing Variable Cells**, then select `B6:D7`.',
        why: 'These are the six yellow cells: the only ones Solver may change. Solver writes them as `$B$6:$D$7`.',
      },
      {
        do: 'Click **Add**. Set **Cell Reference** to `B8:D8`, pick `=` in the middle, set **Constraint** to `B9:D9`, then click **OK**.',
        why: 'The first constraint: each store receives exactly its demand.',
      },
      {
        do: 'Click **Add** again. Set **Cell Reference** to `E6:E7`, pick `<=` in the middle, set **Constraint** to `F6:F7`, then click **OK**.',
        why: 'The second constraint: no warehouse ships more pallets than it has.',
      },
      {
        do: 'Make sure **Make Unconstrained Variables Non-Negative** is ticked.',
        why: 'A pallet count can’t go below 0.',
      },
      {
        do: 'In **Select a Solving Method**, choose **Simplex LP**.',
        why: 'Simplex LP fits any model that only adds up inputs multiplied by fixed numbers, like this one. It finds the true minimum rather than a good guess.',
      },
      {
        do: 'Click **Solve**.',
        why: `Solver tries plans until it finds the cheapest one that meets every constraint. Total cost in \`B11\` should come to ${money(d.best)}.`,
      },
      {
        do: 'In the **Solver Results** box, leave **Keep Solver Solution** selected and click **OK**.',
        why: 'Solver leaves its pallet counts typed into `B6:D7`. Received now matches Demand, and Shipped stays within Supply.',
      },
      checkStep('The coach checks that `B6:D7` hold whole pallet counts of 0 or more, the formulas around them are untouched, every store gets its demand, no warehouse goes over supply, and `B11` is the lowest possible cost.'),
    ];
  },
  make: shippingProblem,
  layout: (d) => ({
    blocks: [
      cells('A1', [['Cost per pallet', ...d.stores]], 'header'),
      cells('A2', column(d.warehouses), 'label'),
      cells('B2', d.cost, 'input', FMT.currency),
      cells('A5', [['Pallets to ship', ...d.stores, 'Shipped', 'Supply']], 'header'),
      cells('A6', column(d.warehouses), 'label'),
      cells('B6', [
        [0, 0, 0],
        [0, 0, 0],
      ], 'input', FMT.int),
      cells('E6', column(SHIP_FORMULAS.shipped), 'formula', FMT.int),
      cells('F6', column(d.supply), 'input', FMT.int),
      cells('A8', column(['Received', 'Demand']), 'label'),
      cells('B8', [[...SHIP_FORMULAS.received]], 'formula', FMT.int),
      cells('B9', [d.demand], 'input', FMT.int),
      cells('A11', [['Total cost']], 'label'),
      cells('B11', [[SHIP_FORMULAS.total]], 'formula', FMT.currency),
    ],
    answer: { kind: 'cellChecks', checks: shippingChecks(d) },
  }),
  expected: () => [],
  inputs: () => [],
  variants: [],
});

export const WHAT_IF: Exercise<any>[] = [goalSeekPrice, dataTableTwoWay, solverShipping];

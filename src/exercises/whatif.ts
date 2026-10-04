import { WAREHOUSES } from '../engine/data';
import { round, type Rng } from '../engine/rng';
import type { CellCheck, Exercise } from '../engine/types';
import { FMT, cells, column, defineExercise, rangeWrite } from './common';

/** "$1,234.50" for solution text. */
function usd(n: number): string {
  const s = Math.abs(n).toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  return `${n < 0 ? '-' : ''}$${s}`;
}

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

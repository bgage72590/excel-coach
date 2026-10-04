import { describe, expect, it } from 'vitest';
import { parseCell } from '../src/engine/address';
import { CARRIERS, CATEGORIES, WAREHOUSES } from '../src/engine/data';
import { Rng, round } from '../src/engine/rng';
import type { Cell, CellCheck, Exercise, Layout } from '../src/engine/types';
import { CHARTS, comboChart, pivotAverageFilter, slicerDashboard } from '../src/exercises/charts';
import {
  DT_PROFIT,
  GOAL_SEEK_STOP,
  GS_FORMULAS,
  GS_MARGIN_TOL,
  GS_PRICE_TOL,
  WHAT_IF,
  dataTableTwoWay,
  goalSeekPrice,
  greedyBreaksSupply,
  marginAt,
  solveShipping,
  solverShipping,
  targetPrice,
} from '../src/exercises/whatif';
import { exerciseSuite } from './helpers/suite';

exerciseSuite([...WHAT_IF, ...CHARTS]);

/** What setup writes into one cell, read back from the layout's cell blocks. */
function cellAt(layout: Layout, address: string): Cell | undefined {
  const target = parseCell(address);
  for (const b of layout.blocks) {
    if (b.kind !== 'cells') continue;
    const s = parseCell(b.at);
    const v = b.values[target.row - s.row]?.[target.col - s.col];
    if (v !== undefined) return v;
  }
  return undefined;
}

function checksOf(layout: Layout): CellCheck[] {
  return layout.answer.kind === 'cellChecks' ? layout.answer.checks : [];
}

const MANY_SEEDS = Array.from({ length: 150 }, (_, i) => i * 7919 + 3);

describe('area basics', () => {
  it('uses unique ids and the right modules', () => {
    const all = [...WHAT_IF, ...CHARTS];
    expect(new Set(all.map((e) => e.id)).size).toBe(all.length);
    expect(WHAT_IF.every((e) => e.module === 'whatif')).toBe(true);
    expect([comboChart.module, slicerDashboard.module, pivotAverageFilter.module]).toEqual(['charts', 'charts', 'pivots']);
  });

  it('every formula a cell check protects is the formula setup writes there', () => {
    for (const ex of [goalSeekPrice, solverShipping] as Exercise<any>[]) {
      for (const seed of [1, 42, 987654]) {
        const d = ex.make(new Rng(seed));
        const layout = ex.layout(d);
        for (const c of checksOf(layout)) if (c.formula) expect(cellAt(layout, c.cell), `${ex.id} ${c.cell}`).toBe(c.formula);
        for (const c of checksOf(layout)) if (c.answer) expect(typeof cellAt(layout, c.cell), `${ex.id} ${c.cell} starts with a number`).toBe('number');
      }
    }
  });
});

describe('goal-seek-price', () => {
  it('the exact price brings the sheet’s margin to the target', () => {
    for (const seed of MANY_SEEDS) {
      const d = goalSeekPrice.make(new Rng(seed));
      expect(Math.abs(marginAt(d, targetPrice(d)) - d.target)).toBeLessThan(1e-12);
    }
  });

  it('any price Goal Seek accepts (margin within its 0.001 stop) is within the price tolerance', () => {
    for (const seed of MANY_SEEDS) {
      const d = goalSeekPrice.make(new Rng(seed));
      const exact = targetPrice(d);
      for (const m of [d.target - GS_MARGIN_TOL, d.target + GS_MARGIN_TOL]) {
        const price = d.cost / (1 - m);
        expect(Math.abs(price - exact), `seed ${seed}`).toBeLessThanOrEqual(GS_PRICE_TOL);
      }
    }
  });

  it('leaves headroom over Goal Seek’s own 0.001 stop', () => {
    expect(GS_MARGIN_TOL).toBeGreaterThan(GOAL_SEEK_STOP * 1.2);
  });

  it('the untouched starting price fails both checks', () => {
    for (const seed of MANY_SEEDS) {
      const d = goalSeekPrice.make(new Rng(seed));
      expect(Math.abs(d.price - targetPrice(d)), `seed ${seed}`).toBeGreaterThan(GS_PRICE_TOL * 3);
      expect(Math.abs(marginAt(d, d.price) - d.target)).toBeGreaterThan(GS_MARGIN_TOL * 3);
    }
  });

  it('starts well short of the target, so Goal Seek has work to do', () => {
    for (const seed of MANY_SEEDS) {
      const d = goalSeekPrice.make(new Rng(seed));
      expect(marginAt(d, d.price)).toBeLessThan(d.target - 0.04);
      expect(d.price).toBeLessThan(targetPrice(d));
    }
  });

  it('grades B5 against the exact price and B11 against the target', () => {
    const d = goalSeekPrice.make(new Rng(7));
    const [price, margin] = checksOf(goalSeekPrice.layout(d));
    expect(price).toMatchObject({ cell: 'B5', holds: 'number', answer: true, tolerance: GS_PRICE_TOL });
    expect(price.value).toBeCloseTo(targetPrice(d), 6);
    expect(margin).toMatchObject({ cell: 'B11', formula: GS_FORMULAS.marginPct, value: d.target, tolerance: GS_MARGIN_TOL });
  });

  it('works through a hand-checked case', () => {
    const d = { item: 'Packing tape', units: 5000, cost: 2.48, price: 3, target: 0.38 };
    expect(targetPrice(d)).toBeCloseTo(4, 12);
    expect(marginAt(d, 4)).toBeCloseTo(0.38, 12);
  });
});

describe('data-table-two-way', () => {
  it('expects profit = (price − cost) × units − fixed for every pair', () => {
    for (const seed of [1, 7, 42, 1234]) {
      const d = dataTableTwoWay.make(new Rng(seed));
      const grid = dataTableTwoWay.expected(d);
      expect(grid.length).toBe(7);
      grid.forEach((row, i) => {
        expect(row.length).toBe(6);
        row.forEach((v, j) => expect(v).toBe((d.prices[j] - d.cost) * d.volumes[i] - d.fixed));
      });
    }
  });

  it('includes the model’s own price and volume, matching the Profit cell', () => {
    const d = dataTableTwoWay.make(new Rng(42));
    const i = d.volumes.indexOf(d.units);
    const j = d.prices.indexOf(d.price);
    expect(i).toBeGreaterThanOrEqual(0);
    expect(j).toBeGreaterThanOrEqual(0);
    expect(dataTableTwoWay.expected(d)[i][j]).toBe((d.price - d.cost) * d.units - d.fixed);
  });

  it('keeps volumes positive and shows both losses and profits on some seeds', () => {
    let sawLoss = false;
    let sawProfit = false;
    for (const seed of MANY_SEEDS) {
      const d = dataTableTwoWay.make(new Rng(seed));
      expect(Math.min(...d.volumes)).toBeGreaterThan(0);
      expect(Math.min(...d.prices)).toBeGreaterThan(d.cost);
      const flat = dataTableTwoWay.expected(d).flat() as number[];
      sawLoss ||= flat.some((v) => v < 0);
      sawProfit ||= flat.some((v) => v > 0);
    }
    expect(sawLoss && sawProfit).toBe(true);
  });

  it('writes the model and corner formulas where the hints say', () => {
    const d = dataTableTwoWay.make(new Rng(1));
    const layout = dataTableTwoWay.layout(d);
    expect(cellAt(layout, 'B6')).toBe(DT_PROFIT);
    expect(cellAt(layout, 'B9')).toBe('=B6');
    expect(cellAt(layout, 'B2')).toBe(d.price);
    expect(cellAt(layout, 'B3')).toBe(d.cost);
    expect(cellAt(layout, 'B4')).toBe(d.units);
    expect(cellAt(layout, 'B5')).toBe(d.fixed);
  });

  it('treats the grid as a live data table: values pass, hand-built formulas fail', async () => {
    const { gradeStructure } = await import('../src/engine/grade');
    const d = dataTableTwoWay.make(new Rng(3));
    const area = dataTableTwoWay.layout(d).answer;
    expect(area.kind === 'cells' && area.liveValues).toBe(true);
    const values = dataTableTwoWay.expected(d) as number[][];
    const asValues = { address: 'C10:H16', values, formulas: values, r1c1: values };
    expect(gradeStructure(area, asValues, 7, 6).items[0].status).toBe('pass');
    const handBuilt = values.map((r) => r.map(() => '=(C$9-$B$3)*$B10-$B$5'));
    expect(gradeStructure(area, { ...asValues, formulas: handBuilt, r1c1: handBuilt }, 7, 6).items[0].status).toBe('fail');
  });


  it('variants rewrite only the cost cells', () => {
    const d = dataTableTwoWay.make(new Rng(3));
    expect(dataTableTwoWay.inputs(d).map((w) => (w.kind === 'range' ? w.address : w.table))).toEqual(['B3', 'B5']);
    for (const [i, v] of dataTableTwoWay.variants.entries()) {
      const changed = v.apply(d, new Rng(100 + i));
      expect(changed.prices).toEqual(d.prices);
      expect(changed.volumes).toEqual(d.volumes);
      expect(changed.cost !== d.cost || changed.fixed !== d.fixed).toBe(true);
    }
  });
});

/** Independent check: with x = warehouse 1's amounts, cost is linear in x; fill by cheapest saving first. */
function greedyOptimum(cost: number[][], supply: number[], demand: number[]): number {
  const total = demand.reduce((a, b) => a + b, 0);
  const lower = Math.max(0, total - supply[1]);
  const delta = demand.map((_, j) => cost[0][j] - cost[1][j]);
  const order = [0, 1, 2].sort((a, b) => delta[a] - delta[b]);
  const x = [0, 0, 0];
  let sent = 0;
  for (const j of order) {
    if (delta[j] >= 0 && sent >= lower) break;
    const room = delta[j] < 0 ? supply[0] - sent : lower - sent;
    const take = Math.min(demand[j], room);
    x[j] = take;
    sent += take;
  }
  return demand.reduce((s, dj, j) => s + cost[0][j] * x[j] + cost[1][j] * (dj - x[j]), 0);
}

describe('solver-shipping', () => {
  it('solves a hand-checked case with and without a binding supply limit', () => {
    const cost = [
      [10, 20, 30],
      [25, 15, 12],
    ];
    const loose = solveShipping(cost, [50, 100], [40, 40, 40]);
    expect(loose.best).toBe(1480);
    expect(loose.plan).toEqual([
      [40, 0, 0],
      [0, 40, 40],
    ]);
    const tight = solveShipping(cost, [30, 100], [40, 40, 40]);
    expect(tight.best).toBe(1630);
    expect(tight.plan).toEqual([
      [30, 0, 0],
      [10, 40, 40],
    ]);
    expect(greedyBreaksSupply(cost, [30, 100], [40, 40, 40])).toBe(true);
    expect(greedyBreaksSupply(cost, [50, 100], [40, 40, 40])).toBe(false);
  });

  it('generated plans are feasible, optimal, unique and need the supply limits', () => {
    for (const seed of MANY_SEEDS.slice(0, 80)) {
      const d = solverShipping.make(new Rng(seed));
      const { plan, best } = d;
      for (const row of plan) for (const x of row) expect(Number.isInteger(x) && x >= 0).toBe(true);
      plan.forEach((row, i) => expect(row.reduce((a, b) => a + b, 0)).toBeLessThanOrEqual(d.supply[i]));
      d.demand.forEach((dj, j) => expect(plan[0][j] + plan[1][j]).toBe(dj));
      expect(plan.flatMap((row, i) => row.map((x, j) => x * d.cost[i][j])).reduce((a, b) => a + b, 0)).toBe(best);
      expect(best).toBe(greedyOptimum(d.cost, d.supply, d.demand));
      expect(solveShipping(d.cost, d.supply, d.demand).ties, `seed ${seed} has one cheapest plan`).toBe(1);
      expect(greedyBreaksSupply(d.cost, d.supply, d.demand), `seed ${seed} needs Solver`).toBe(true);
      const total = d.demand.reduce((a, b) => a + b, 0);
      expect(d.supply[0] + d.supply[1]).toBeGreaterThanOrEqual(total);
      expect(Math.max(...d.supply)).toBeLessThan(total);
    }
  });

  it('checks six decisions, three receipts, two shipments and the total', () => {
    const d = solverShipping.make(new Rng(42));
    const checks = checksOf(solverShipping.layout(d));
    expect(checks.length).toBe(12);
    const decisions = checks.filter((c) => c.answer);
    expect(decisions.map((c) => c.cell)).toEqual(['B6', 'C6', 'D6', 'B7', 'C7', 'D7']);
    for (const c of decisions) expect(c).toMatchObject({ holds: 'number', min: 0, integer: true });
    expect(checks.filter((c) => /^[BCD]8$/.test(c.cell)).map((c) => c.value)).toEqual(d.demand);
    expect(checks.filter((c) => /^E[67]$/.test(c.cell)).map((c) => c.max)).toEqual(d.supply);
    const total = checks.find((c) => c.cell === 'B11')!;
    expect(total.value).toBe(d.best);
    expect(total.tolerance).toBe(0.01);
  });

  it('the optimal plan, typed into the decision cells, passes every value check', () => {
    const d = solverShipping.make(new Rng(1234));
    const checks = checksOf(solverShipping.layout(d));
    const value = (cell: string): number => {
      const { row, col } = parseCell(cell);
      const x = (r: number, c: number) => d.plan[r - 6][c - 2];
      if (row >= 6 && row <= 7 && col <= 4) return x(row, col);
      if (row === 8) return x(6, col) + x(7, col);
      if (col === 5) return [2, 3, 4].reduce((s, c) => s + x(row, c), 0);
      return d.best;
    };
    for (const c of checks) {
      const v = value(c.cell);
      if (c.value !== undefined) expect(Math.abs(v - c.value)).toBeLessThanOrEqual(c.tolerance ?? 0.005);
      if (c.max !== undefined) expect(v).toBeLessThanOrEqual(c.max);
      if (c.min !== undefined) expect(v).toBeGreaterThanOrEqual(c.min);
    }
  });
});

describe('charts and pivots', () => {
  it('combo chart data has twelve months of revenue and margin', () => {
    const d = comboChart.make(new Rng(5));
    expect(d.rows.map((r) => r.month)).toEqual(['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']);
    for (const r of d.rows) {
      expect(r.revenue % 100).toBe(0);
      expect(r.margin).toBeGreaterThan(0.2);
      expect(r.margin).toBeLessThan(0.4);
    }
    expect(comboChart.inspections!(d)).toEqual([expect.objectContaining({ kind: 'chart', minSeries: 2, secondaryLine: true, title: true })]);
  });

  it('inventory covers every warehouse and category, with value = on hand × cost', () => {
    for (const seed of [1, 7, 42]) {
      const d = slicerDashboard.make(new Rng(seed));
      expect(new Set(d.rows.map((r) => r.warehouse))).toEqual(new Set(WAREHOUSES));
      expect(new Set(d.rows.map((r) => r.category))).toEqual(new Set(CATEGORIES));
      for (const r of d.rows) expect(r.value).toBe(round(r.onHand * r.cost, 2));
    }
    const kinds = slicerDashboard.inspections!(slicerDashboard.make(new Rng(1))).map((i) => i.kind);
    expect(kinds).toEqual(['pivot', 'slicer']);
  });

  it('deliveries cover every carrier and warehouse, with whole days late', () => {
    for (const seed of [1, 7, 42]) {
      const d = pivotAverageFilter.make(new Rng(seed));
      expect(new Set(d.rows.map((r) => r.carrier))).toEqual(new Set(CARRIERS));
      expect(new Set(d.rows.map((r) => r.warehouse))).toEqual(new Set(WAREHOUSES));
      for (const r of d.rows) expect(Number.isInteger(r.late) && r.late >= -2 && r.late <= 9).toBe(true);
    }
    expect(pivotAverageFilter.inspections!(pivotAverageFilter.make(new Rng(1)))[0]).toMatchObject({
      kind: 'pivot',
      rows: 'Carrier',
      valuesField: 'Days late',
      summarizeBy: 'Average',
      filter: 'Warehouse',
    });
  });
});

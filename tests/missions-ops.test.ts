import { describe, expect, it } from 'vitest';
import { parseCell, parseRange, type RangeRef } from '../src/engine/address';
import type { Block, Cell, DataBlock } from '../src/engine/types';
import { gradeRules } from '../src/engine/grade';
import { Rng } from '../src/engine/rng';
import { EXERCISES } from '../src/exercises';
import {
  OPS_MISSIONS,
  PLAN_DAYS,
  fixDivisibility,
  isLow,
  monthEndInventory,
  orderQtyOf,
  planReorderPoint,
  reorderPlan,
  totalUnits,
  type InventoryData,
  type PlanData,
} from '../src/missions/ops';
import { SEEDS, blockRange, missionSuite, overlaps } from './helpers/suite';

missionSuite(OPS_MISSIONS);

/**
 * The base data for a seed plus every variant of every step, seeded exactly as ExcelHost.check
 * seeds them (seed + 7919 × (variant index + 1)), so these tests see the states a learner meets.
 */
function allStates<D>(m: { make(rng: Rng): D; steps: { variants: { apply(d: D, rng: Rng): D }[] }[] }, seed: number): D[] {
  const base = m.make(new Rng(seed));
  return [base, ...m.steps.flatMap((s) => s.variants.map((v, i) => v.apply(base, new Rng(seed + 7919 * (i + 1)))))];
}

/** Session seeds are random 32-bit numbers, so invariants are checked over many of them. */
const MANY_SEEDS = [...SEEDS, ...Array.from({ length: 300 }, (_, k) => ((k + 1) * 2654435761) >>> 0)];

/** Reads a Table's rows back out of the sheet blocks, keyed by header: what Excel will hold. */
function tableRows(blocks: Block[], table: string): Record<string, Cell>[] {
  const b = blocks.find((x): x is DataBlock => x.kind === 'data' && x.table === table)!;
  return b.rows.map((r) => Object.fromEntries(b.columns.map((c, i) => [c.header, r[i]])));
}

function cellsAt(blocks: Block[], at: string): Cell[][] {
  const b = blocks.find((x) => x.kind === 'cells' && x.at === at)!;
  return b.kind === 'cells' ? b.values : [];
}

/** Excel's SORT on plain ASCII text: case-insensitive. */
const excelSort = (xs: string[]) => [...xs].sort((a, b) => (a.toLowerCase() < b.toLowerCase() ? -1 : a.toLowerCase() > b.toLowerCase() ? 1 : 0));

describe('ops missions: metadata', () => {
  it('only lists skills that exist', () => {
    const ids = new Set(EXERCISES.map((e) => e.id));
    for (const m of OPS_MISSIONS) for (const s of m.skills) expect(ids.has(s), `${m.id} skill ${s}`).toBe(true);
  });

  it('are ops missions with four steps and a named manager', () => {
    for (const m of OPS_MISSIONS) {
      expect(m.role).toBe('ops');
      expect(m.steps.length).toBe(4);
      expect(m.brief(m.make(new Rng(1))).from).toMatch(/^[A-Z][a-z]+ [A-Z][a-z]+, /);
    }
  });
});

describe('ops missions: sheet layout beyond the shared checks', () => {
  it.each(SEEDS)('seed %i: nothing sits in the 80 rows the host formats below a spill anchor', (seed) => {
    for (const m of OPS_MISSIONS) {
      const d = m.make(new Rng(seed));
      for (const [i, step] of m.steps.entries()) {
        const a = step.answer(d);
        if (a.kind !== 'spill') continue;
        const start = parseCell(a.anchor);
        const width = step.expected(d)[0].length;
        const zone: RangeRef = { start, end: { row: start.row + 79, col: start.col + width } };
        for (const b of m.blocks(d)) expect(overlaps(zone, blockRange(b)), `${m.id} block at ${b.at}`).toBe(false);
        m.steps.forEach((other, j) => {
          const o = other.answer(d);
          if (j !== i && o.kind === 'cells') expect(overlaps(zone, parseRange(o.range)), `${m.id} step ${j + 1}`).toBe(false);
        });
      }
    }
  });

  it.each(SEEDS)('seed %i: the Items Table has room for the three plan columns plus a gap', (seed) => {
    const d = reorderPlan.make(new Rng(seed));
    const items = reorderPlan.blocks(d).find((b) => b.kind === 'data' && b.table === 'Items')!;
    const r = blockRange(items);
    expect(r.end.col).toBe(8); // E:H, the learner adds I, J and K
    // I:L must stay clear: three new columns and the gap Excel needs after them.
    const room: RangeRef = { start: { row: 1, col: 9 }, end: { row: r.end.row + 2, col: 12 } };
    for (const b of reorderPlan.blocks(d)) if (b !== items) expect(overlaps(room, blockRange(b)), `block at ${b.at}`).toBe(false);
    const spill = reorderPlan.steps[3].answer(d);
    expect(spill.kind === 'spill' && parseCell(spill.anchor).col).toBeGreaterThan(12);
    const plan = reorderPlan.steps.slice(0, 3).map((s) => s.answer(d));
    expect(plan.map((a) => (a.kind === 'tableColumn' ? a.column : ''))).toEqual(['Avg daily demand', 'Reorder point', 'Order qty']);
  });
});

describe('month-end inventory data', () => {
  it('snapshot is consistent and every check stays meaningful in every state', () => {
    for (const seed of MANY_SEEDS) for (const d of allStates<InventoryData>(monthEndInventory, seed)) {
      const skus = d.rows.map((r) => r.sku);
      expect(new Set(skus).size, 'SKUs are unique').toBe(skus.length);
      expect(new Set(d.rows.map((r) => r.item)).size, 'items are unique').toBe(d.rows.length);
      expect(d.rows.length).toBeGreaterThanOrEqual(40);
      expect(d.rows.length).toBeLessThanOrEqual(50);
      for (const r of d.rows) {
        expect(r.value).toBeCloseTo(r.onHand * r.cost, 2);
        expect(Number.isInteger(r.onHand) && r.onHand >= 0).toBe(true);
        expect(Number.isInteger(r.reorder) && r.reorder >= 2).toBe(true);
        expect(r.sku).toMatch(/^SKU-\d{4}$/);
      }
      expect(d.rows.filter(isLow).length, 'reorder list is never short').toBeGreaterThanOrEqual(6);
      expect(d.rows.some((r) => r.onHand === r.reorder), 'a SKU sits exactly on its reorder point').toBe(true);
      expect(d.rows.some((r) => r.onHand === 0), 'a SKU is stocked out').toBe(true);
    }
  });

  it.each(SEEDS)('seed %i: grid, counts and shares reconcile', (seed) => {
    const d = monthEndInventory.make(new Rng(seed));
    const [grid, counts, shares, list] = monthEndInventory.steps.map((s) => s.expected(d));
    const total = d.rows.reduce((a, r) => a + r.value, 0);
    const gridTotal = grid.flat().reduce<number>((a, v) => a + (v as number), 0);
    expect(gridTotal).toBeCloseTo(total, 6);
    expect(shares.flat().reduce<number>((a, v) => a + (v as number), 0)).toBeCloseTo(1, 9);
    expect(counts[0].reduce<number>((a, v) => a + (v as number), 0)).toBe(list.length);
    const listed = list.map((r) => r[0] as string);
    expect(listed).toEqual([...listed].sort());
  });

  it('base rows are grouped by warehouse, so the reorder list needs SORT', () => {
    const d = monthEndInventory.make(new Rng(42));
    const order = d.rows.map((r) => r.warehouse);
    expect(order).toEqual([...order].sort());
    const unsorted = d.rows.filter(isLow).map((r) => r.sku);
    expect(unsorted).not.toEqual([...unsorted].sort());
  });
});

describe('reorder plan data', () => {
  it('rounds up like ROUNDUP(total / 28 * (lead + safety), 0)', () => {
    expect(planReorderPoint(50, 10)).toBe(18); // 17.857…
    expect(planReorderPoint(29, 5)).toBe(6); // 5.178…
    expect(planReorderPoint(1, 5)).toBe(1); // 0.178…
    expect(planReorderPoint(0, 12)).toBe(0);
    expect(planReorderPoint(56, 10)).toBe(20); // exact: the case fixDivisibility steers away from
  });

  it('nudges demand off exact whole-number reorder points', () => {
    const d: PlanData = {
      demand: [
        { date: 46280, sku: 'SKU-3001', units: 10 },
        { date: 46281, sku: 'SKU-3001', units: 4 },
        { date: 46282, sku: 'SKU-3002', units: 9 },
      ],
      items: [
        { sku: 'SKU-3001', lead: 4, safety: 2, onHand: 0, weight: 1, maxUnits: 10 }, // 14 × 6 = 84 = 3 × 28
        { sku: 'SKU-3002', lead: 4, safety: 2, onHand: 0, weight: 1, maxUnits: 10 }, // 9 × 6 = 54, fine
      ],
    };
    const fixed = fixDivisibility(d);
    expect(totalUnits(fixed, 'SKU-3001')).toBe(15);
    expect(totalUnits(fixed, 'SKU-3002')).toBe(9);
    expect(d.demand[0].units, 'input is not mutated').toBe(10);
  });

  it('every state keeps ROUNDUP unambiguous and the order list non-empty', () => {
    for (const seed of MANY_SEEDS) for (const d of allStates<PlanData>(reorderPlan, seed)) {
      const skus = d.items.map((i) => i.sku);
      expect(new Set(skus).size).toBe(skus.length);
      expect(d.demand.length).toBeLessThanOrEqual(120);
      for (const l of d.demand) {
        expect(skus).toContain(l.sku);
        expect(Number.isInteger(l.units) && l.units > 0).toBe(true);
      }
      for (const it of d.items) {
        expect((it.lead + it.safety) % PLAN_DAYS).not.toBe(0);
        const total = totalUnits(d, it.sku);
        const product = total * (it.lead + it.safety);
        if (total > 0) expect(product % PLAN_DAYS, `${it.sku} reorder point is not a whole number`).not.toBe(0);
        // Excel's route (average first, then multiply) lands on the same side of every whole number.
        const viaAverage = (total / PLAN_DAYS) * (it.lead + it.safety);
        expect(Math.ceil(viaAverage)).toBe(planReorderPoint(total, it.lead + it.safety));
        if (total > 0) expect(Math.abs(viaAverage - Math.round(viaAverage))).toBeGreaterThan(0.03);
      }
      const needs = d.items.filter((it) => orderQtyOf(d, it) > 0).length;
      expect(needs).toBeGreaterThanOrEqual(4);
    }
  });

  it.each(SEEDS)('seed %i: history spans the 28-day window and has a SKU with no demand', (seed) => {
    const d = reorderPlan.make(new Rng(seed));
    const dates = d.demand.map((l) => l.date);
    expect(Math.max(...dates) - Math.min(...dates)).toBeLessThan(PLAN_DAYS);
    expect(d.items.some((it) => totalUnits(d, it.sku) === 0)).toBe(true);
    expect(d.items.some((it) => orderQtyOf(d, it) === 0)).toBe(true);
  });
});

describe('answer keys re-derived from the sheet, the way the intended formulas compute them', () => {
  it('month-end inventory: SUMIFS grid, SUMPRODUCT counts, shares and the sorted FILTER', () => {
    for (const seed of MANY_SEEDS.slice(0, 60)) {
      for (const d of allStates<InventoryData>(monthEndInventory, seed)) {
        const blocks = monthEndInventory.blocks(d);
        const inv = tableRows(blocks, 'Inventory');
        const whs = cellsAt(blocks, 'K1')[0] as string[];
        const cats = cellsAt(blocks, 'J2').map((r) => r[0] as string);
        const cats2 = cellsAt(blocks, 'J11').map((r) => r[0] as string);
        const eq = (a: Cell, b: string) => String(a).toLowerCase() === b.toLowerCase();
        const low = (r: Record<string, Cell>) => (r['On hand'] as number) <= (r['Reorder point'] as number);
        const total = inv.reduce((a, r) => a + (r.Value as number), 0);
        const want = [
          cats.map((c) => whs.map((w) => inv.filter((r) => eq(r.Category, c) && eq(r.Warehouse, w)).reduce((a, r) => a + (r.Value as number), 0))),
          [whs.map((w) => inv.filter((r) => eq(r.Warehouse, w) && low(r)).length)],
          cats2.map((c) => [inv.filter((r) => eq(r.Category, c)).reduce((a, r) => a + (r.Value as number), 0) / total]),
        ];
        const listed = excelSort(inv.filter(low).map((r) => r.SKU as string));
        const bySku = new Map(inv.map((r) => [r.SKU as string, r]));
        const list = listed.map((sku) => ['SKU', 'Item', 'Category', 'Warehouse'].map((h) => bySku.get(sku)![h]));
        const got = monthEndInventory.steps.map((s) => s.expected(d));
        for (let i = 0; i < 3; i++) {
          got[i].forEach((row, r) => row.forEach((v, c) => expect(Math.abs((v as number) - (want[i][r][c] as number))).toBeLessThan(1e-9 * Math.max(1, Math.abs(want[i][r][c] as number)))));
        }
        expect(got[3]).toEqual(list);
      }
    }
  });

  it('reorder plan: average, ROUNDUP reorder point, MAX order qty and the sorted FILTER', () => {
    for (const seed of MANY_SEEDS.slice(0, 60)) {
      for (const d of allStates<PlanData>(reorderPlan, seed)) {
        const blocks = reorderPlan.blocks(d);
        const demand = tableRows(blocks, 'Demand');
        const items = tableRows(blocks, 'Items');
        const days = cellsAt(blocks, 'N1')[0][0] as number;
        // Excel's order of operations: SUMIFS ÷ N1, then × (lead + safety), then ROUNDUP.
        const avg = items.map((it) => demand.filter((l) => l.SKU === it.SKU).reduce((a, l) => a + (l.Units as number), 0) / days);
        const rp = items.map((it, i) => Math.ceil(avg[i] * ((it['Lead days'] as number) + (it['Safety days'] as number))));
        const qty = items.map((it, i) => Math.max(0, rp[i] - (it['On hand'] as number)));
        const list = excelSort(items.filter((_, i) => qty[i] > 0).map((it) => it.SKU as string)).map((s) => [s]);
        const [e1, e2, e3, e4] = reorderPlan.steps.map((s) => s.expected(d));
        expect(e1).toEqual(avg.map((v) => [v]));
        expect(e2).toEqual(rp.map((v) => [v]));
        expect(e3).toEqual(qty.map((v) => [v]));
        expect(e4).toEqual(list);
      }
    }
  });
});

describe('rules accept realistic correct formulas', () => {
  const good: Record<string, string[]> = {
    'mission-month-end-inventory#1': ['=SUMIFS(Inventory[Value], Inventory[Category], $J2, Inventory[Warehouse], K$1)'],
    'mission-month-end-inventory#2': [
      '=SUM((Inventory[Warehouse]=K$1)*(Inventory[On hand]<=Inventory[Reorder point]))',
      '=SUMPRODUCT(--(Inventory[Warehouse]=K1),--(Inventory[On hand]<=Inventory[Reorder point]))',
    ],
    'mission-month-end-inventory#3': ['=SUMIF(Inventory[Category],$J11,Inventory[Value])/SUM(Inventory[Value])'],
    'mission-month-end-inventory#4': [
      '=_xlfn._xlws.SORT(_xlfn._xlws.FILTER(Inventory[[SKU]:[Warehouse]],Inventory[On hand]<=Inventory[Reorder point]))',
      '=SORT(FILTER(Inventory[[SKU]:[Warehouse]],Inventory[On hand]<=Inventory[Reorder point]),1,1)',
    ],
    'mission-reorder-plan#1': ['=SUMIF(Demand[SKU],[@SKU],Demand[Units])/N$1', '=SUMIFS(Demand[Units],Demand[SKU],Items[@SKU])/$N$1'],
    'mission-reorder-plan#2': [
      '=CEILING([@[Avg daily demand]]*([@[Lead days]]+[@[Safety days]]),1)',
      '=ROUNDUP(SUMIFS(Demand[Units],Demand[SKU],[@SKU])/$N$1*([@[Lead days]]+[@[Safety days]]),0)',
      '=_xlfn.CEILING.MATH([@[Avg daily demand]]*([@[Lead days]]+[@[Safety days]]))',
    ],
    'mission-reorder-plan#3': ['=IF([@[On hand]]<[@[Reorder point]],[@[Reorder point]]-[@[On hand]],0)', '=MAX([@[Reorder point]]-[@[On hand]],0)'],
    'mission-reorder-plan#4': ['=_xlfn._xlws.SORT(_xlfn._xlws.FILTER(Items[SKU],Items[Order qty]>0),1,1)'],
  };
  it('no rule fails on the solution or common equivalents', () => {
    const failures: string[] = [];
    for (const m of OPS_MISSIONS) {
      const d = m.make(new Rng(1));
      m.steps.forEach((s, i) => {
        const id = `${m.id}#${i + 1}`;
        for (const f of [s.solution(d), ...(good[id] ?? [])]) {
          for (const item of gradeRules(s.rules, [f])) if (item.status === 'fail') failures.push(`${id} ${f}: ${item.label}`);
        }
      });
    }
    expect(failures).toEqual([]);
  });
});

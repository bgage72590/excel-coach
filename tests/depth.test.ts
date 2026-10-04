import { describe, expect, it } from 'vitest';
import { cellMatches, isMatcher } from '../src/engine/compare';
import { gradeRules } from '../src/engine/grade';
import { CATEGORIES } from '../src/engine/data';
import { Rng } from '../src/engine/rng';
import {
  BANDS,
  CORNER,
  DEPTH,
  TOTAL_LABEL,
  ZONES,
  byrowPeak,
  cardRate,
  lookupTwoWay,
  pivotBySum,
  pivotbyVendorCategory,
  tierRate,
  xlookupTiered,
} from '../src/exercises/depth';
import { SEEDS, exerciseSuite } from './helpers/suite';

exerciseSuite(DEPTH);

describe('tierRate (approximate match, next smaller)', () => {
  const tiers = [
    { min: 0, rate: 0.02 },
    { min: 25000, rate: 0.03 },
    { min: 50000, rate: 0.045 },
  ];

  it('returns the tier of the largest threshold at or below the sale', () => {
    expect(tierRate(tiers, 1)).toBe(0.02);
    expect(tierRate(tiers, 24999)).toBe(0.02);
    expect(tierRate(tiers, 30000)).toBe(0.03);
    expect(tierRate(tiers, 999999)).toBe(0.045);
  });

  it('gives a sale exactly on a threshold that tier', () => {
    expect(tierRate(tiers, 25000)).toBe(0.03);
    expect(tierRate(tiers, 50000)).toBe(0.045);
    expect(tierRate(tiers, 0)).toBe(0.02);
  });

  it('does not depend on the order of the tiers', () => {
    expect(tierRate([...tiers].reverse(), 30000)).toBe(0.03);
  });
});

describe('xlookup-tiered data', () => {
  it.each(SEEDS)('seed %i: tiers ascend from $0, every sale is covered, one sale sits on a threshold', (seed) => {
    const d = xlookupTiered.make(new Rng(seed));
    expect(d.tiers[0].min).toBe(0);
    for (let i = 1; i < d.tiers.length; i++) {
      expect(d.tiers[i].min).toBeGreaterThan(d.tiers[i - 1].min);
      expect(d.tiers[i].rate).toBeGreaterThan(d.tiers[i - 1].rate);
    }
    for (const r of d.reps) expect(r.sales).toBeGreaterThanOrEqual(0);
    expect(d.reps.some((r) => d.tiers.slice(1).some((t) => t.min === r.sales))).toBe(true);
    expect(new Set(d.reps.map((r) => r.rep)).size).toBe(d.reps.length);
  });

  it.each(SEEDS)('seed %i: the new-top-tier variant puts a rep in that tier', (seed) => {
    const base = xlookupTiered.make(new Rng(seed));
    const v = xlookupTiered.variants.find((x) => x.label.includes('new top tier'))!;
    const d = v.apply(base, new Rng(seed + 3));
    expect(d.tiers.length).toBe(base.tiers.length + 1);
    const top = d.tiers[d.tiers.length - 1];
    expect(xlookupTiered.expected(d).some((row) => row[0] === top.rate)).toBe(true);
  });
});

describe('lookup-two-way', () => {
  it('finds the cell where the zone row and band column cross', () => {
    const card = { zones: ['Zone 2', 'Zone 1'], rates: [[20, 30, 40, 50, 60], [10, 11, 12, 13, 14]] };
    expect(cardRate(card, 'Zone 1', BANDS[0])).toBe(10);
    expect(cardRate(card, 'Zone 2', BANDS[4])).toBe(60);
  });

  it.each(SEEDS)('seed %i: rates climb with weight, and re-sorting zones keeps each shipment’s rate', (seed) => {
    const d = lookupTwoWay.make(new Rng(seed));
    expect(d.card.zones).toEqual([...ZONES]);
    for (const row of d.card.rates) for (let b = 1; b < row.length; b++) expect(row[b]).toBeGreaterThan(row[b - 1]);
    const resort = lookupTwoWay.variants.find((x) => x.label.includes('re-sorted'))!;
    const moved = resort.apply(d, new Rng(seed + 1));
    expect(moved.card.zones).not.toEqual(d.card.zones);
    expect(lookupTwoWay.expected(moved)).toEqual(lookupTwoWay.expected(d));
  });

  it('labels are text Excel will not convert to numbers or dates', () => {
    for (const label of [...ZONES, ...BANDS]) expect(Number.isNaN(Number(label))).toBe(true);
  });
});

describe('byrow-peak', () => {
  it.each(SEEDS)('seed %i: one whole-number peak per SKU', (seed) => {
    const d = byrowPeak.make(new Rng(seed));
    const expected = byrowPeak.expected(d);
    expect(expected.length).toBe(d.rows.length);
    d.rows.forEach((r, i) => {
      expect(r.units.length).toBe(7);
      expect(expected[i][0]).toBe(Math.max(...r.units));
      expect(Number.isInteger(expected[i][0])).toBe(true);
    });
    expect(new Set(d.rows.map((r) => r.sku)).size).toBe(d.rows.length);
  });
});

describe('pivotBySum (PIVOTBY default layout)', () => {
  const rows = [
    { row: 'b', col: 'Y', value: 1 },
    { row: 'a', col: 'Y', value: 2 },
    { row: 'a', col: 'X', value: 4 },
    { row: 'b', col: 'X', value: 8 },
    { row: 'a', col: 'X', value: 16 },
  ];

  it('puts a corner, sorted column headers and a Total column on top', () => {
    const out = pivotBySum(rows);
    expect(out[0].slice(1, 3)).toEqual(['X', 'Y']);
    expect(out[0][0]).toBe(CORNER);
    expect(out[0][3]).toBe(TOTAL_LABEL);
  });

  it('sorts row labels, sums each cell and adds row and column totals', () => {
    const out = pivotBySum(rows);
    expect(out.slice(1)).toEqual([
      ['a', 20, 2, 22],
      ['b', 8, 1, 9],
      [TOTAL_LABEL, 28, 3, 31],
    ]);
  });

  it('matchers accept the labels Excel shows', () => {
    expect(cellMatches('', CORNER)).toBe(true);
    expect(cellMatches('Total', TOTAL_LABEL)).toBe(true);
    expect(cellMatches('Grand Total', TOTAL_LABEL)).toBe(true);
    expect(cellMatches(12, CORNER)).toBe(false);
  });
});

describe('pivotby-vendor-category data', () => {
  it.each(SEEDS)('seed %i: every vendor and category pair has spend, so no cell is blank', (seed) => {
    const base = pivotbyVendorCategory.make(new Rng(seed));
    const datasets = [base, ...pivotbyVendorCategory.variants.map((v, i) => v.apply(base, new Rng(seed + i)))];
    for (const d of datasets) {
      const vendors = new Set(d.rows.map((r) => r.vendor));
      for (const v of vendors) for (const c of CATEGORIES) expect(d.rows.some((r) => r.vendor === v && r.category === c)).toBe(true);
      const out = pivotbyVendorCategory.expected(d);
      expect(out.length).toBe(vendors.size + 2);
      expect(out[0].length).toBe(CATEGORIES.length + 2);
      // Row totals and column totals both add up to the grand total.
      const body = out.slice(1, -1);
      const grand = out[out.length - 1][out[0].length - 1] as number;
      const rowSum = body.reduce((a, r) => a + (r[r.length - 1] as number), 0);
      const colSum = out[out.length - 1].slice(1, -1).reduce<number>((a, c) => a + (c as number), 0);
      expect(rowSum).toBeCloseTo(grand, 6);
      expect(colSum).toBeCloseTo(grand, 6);
      for (const r of body) for (const c of r.slice(1)) expect(isMatcher(c) || typeof c === 'number').toBe(true);
    }
  });
});

describe('rules accept realistic correct formulas and reject typed shortcuts', () => {
  const failing = (ex: { rules?: Parameters<typeof gradeRules>[0] }, f: string) =>
    gradeRules(ex.rules, [f]).filter((i) => i.status === 'fail').map((i) => i.label);

  it.each([
    '=XLOOKUP(E2,Tiers[Min sales],Tiers[Rate],,-1)',
    '=_xlfn.XLOOKUP(E2,Tiers[Min sales],Tiers[Rate],0,-1)',
    '=XLOOKUP(E2, Tiers[Min sales], Tiers[Rate], "", -1, 1)',
    '=VLOOKUP(E2,Tiers,2,TRUE)',
    '=VLOOKUP($E2,Tiers[#Data],2)',
    '=INDEX(Tiers[Rate],MATCH(E2,Tiers[Min sales],1))',
    '=INDEX(Tiers[Rate],XMATCH(E2,Tiers[Min sales],-1))',
    '=LOOKUP(E2,Tiers[Min sales],Tiers[Rate])',
  ])('xlookup-tiered passes %s', (f) => expect(failing(xlookupTiered, f)).toEqual([]));

  it('xlookup-tiered flags a fixed range, which Excel would quietly grow with the Table', () => {
    expect(failing(xlookupTiered, '=VLOOKUP(E2,$A$2:$B$7,2,TRUE)')).toEqual(['Refers to the Tiers Table by name']);
  });

  it('xlookup-tiered flags typed thresholds', () => {
    expect(failing(xlookupTiered, '=IF(E2>=50000,0.05,IF(E2>=25000,0.035,0.02))').length).toBeGreaterThan(0);
  });

  it.each([
    '=INDEX($B$2:$F$7,MATCH(I2,$A$2:$A$7,0),MATCH(J2,$B$1:$F$1,0))',
    '=INDEX($A$1:$F$7,MATCH($I2,$A:$A,0),MATCH($J2,$1:$1,0))',
    '=XLOOKUP(I2,$A$2:$A$7,XLOOKUP(J2,$B$1:$F$1,$B$2:$F$7))',
    '=INDEX($B$2:$F$7,XMATCH(I2,$A$2:$A$7),XMATCH(J2,$B$1:$F$1))',
  ])('lookup-two-way passes %s', (f) => expect(failing(lookupTwoWay, f)).toEqual([]));

  it('lookup-two-way flags a typed zone or band position', () => {
    expect(failing(lookupTwoWay, '=INDEX($B$2:$F$7,MATCH(I2,$A$2:$A$7,0),3)').length).toBeGreaterThan(0);
    expect(failing(lookupTwoWay, '=XLOOKUP("Zone 1",$A$2:$A$7,$B$2:$B$7)').length).toBeGreaterThan(0);
  });

  it.each([
    '=BYROW(DailyUnits[[Mon]:[Sun]],LAMBDA(r,MAX(r)))',
    '=_xlfn.BYROW(DailyUnits[[Mon]:[Sun]],_xlfn.LAMBDA(_xlpm.r,MAX(_xlpm.r)))',
    '=BYROW(DailyUnits[[Mon]:[Sun]],MAX)',
  ])('byrow-peak passes %s', (f) => expect(failing(byrowPeak, f)).toEqual([]));

  it('byrow-peak flags a fixed range, which Excel would quietly grow with the Table', () => {
    expect(failing(byrowPeak, '=BYROW($B$2:$H$21,LAMBDA(r,MAX(r)))')).toEqual(['Refers to the DailyUnits Table’s columns by name']);
  });

  it.each(['=PIVOTBY(Spend[Vendor],Spend[Category],Spend[Amount],SUM)', '=_xlfn.PIVOTBY(Spend[Vendor],Spend[Category],Spend[Amount],_xleta.SUM)'])(
    'pivotby-vendor-category passes %s',
    (f) => expect(failing(pivotbyVendorCategory, f)).toEqual([]),
  );
});

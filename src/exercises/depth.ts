import { CATEGORIES, EXTRA_VENDORS, REPS, VENDORS, excelTextCompare, sum } from '../engine/data';
import { round, type Rng } from '../engine/rng';
import type { CellMatcher, ColumnSpec, ExpectedGrid, Grid } from '../engine/types';
import { FMT, cells, dataBlock, defineExercise, rangeWrite, tableWrite } from './common';
import { SPEND_COLS, spendGrid, spendLine } from './sumifs';

/** Whole dollars, for sales figures and thresholds. */
const DOLLARS = '$#,##0';

// ---------- Tiered commission rate, no nested IFs ----------

export interface Tier {
  min: number;
  rate: number;
}

interface RepSales {
  rep: string;
  sales: number;
}

interface TierData {
  tiers: Tier[];
  reps: RepSales[];
}

const TIER_COLS: ColumnSpec[] = [
  { header: 'Min sales', format: DOLLARS },
  { header: 'Rate', format: FMT.pct },
];

const TIER_COUNT = 6;

/** Ascending thresholds starting at $0, rates climbing in half-point steps. */
function makeTiers(rng: Rng): Tier[] {
  const tiers: Tier[] = [{ min: 0, rate: rng.pick([0.02, 0.025, 0.03]) }];
  for (let i = 1; i < TIER_COUNT; i++) {
    const prev = tiers[i - 1];
    tiers.push({ min: prev.min + rng.int(3, 7) * 5000, rate: round(prev.rate + rng.int(1, 3) * 0.005, 4) });
  }
  return tiers;
}

/**
 * The rate of the tier with the largest Min sales at or below `sales`: what
 * XLOOKUP(…, -1), VLOOKUP(…, TRUE) and MATCH(…, 1) return. A sale exactly on a
 * threshold earns that tier.
 */
export function tierRate(tiers: readonly Tier[], sales: number): number {
  let best: Tier | undefined;
  for (const t of tiers) if (t.min <= sales && (!best || t.min > best.min)) best = t;
  if (!best) throw new Error(`Sales ${sales} fall below the first tier`);
  return best.rate;
}

/** Sales spread across every tier, with one rep landing exactly on a threshold. */
function makeSales(rng: Rng, tiers: readonly Tier[], reps: readonly string[]): RepSales[] {
  const top = tiers.length - 1;
  const out = reps.map((rep) => {
    const t = rng.int(0, top);
    const lo = Math.max(tiers[t].min + 1, 2000);
    const hi = t < top ? tiers[t + 1].min - 1 : tiers[t].min + 40000;
    return { rep, sales: rng.int(lo, hi) };
  });
  const k = rng.int(0, out.length - 1);
  out[k] = { ...out[k], sales: tiers[rng.int(1, top)].min };
  return out;
}

const tierGrid = (tiers: Tier[]) => tiers.map((t) => [t.min, t.rate]);
const repGrid = (reps: RepSales[]) => reps.map((r) => [r.rep, r.sales]);

export const xlookupTiered = defineExercise<TierData>({
  id: 'xlookup-tiered',
  module: 'lookups',
  title: 'Tiered commission rate, no nested IFs',
  replaces: 'Nested IFs rewritten every time the commission plan changes',
  minutes: 5,
  task: () =>
    'Commission is tiered: each rep earns the rate of the highest tier whose `Min sales` is at or below their quarter sales. In `F2:F9`, return each rep’s rate from the `Tiers` Table, using one approximate-match lookup in `F2` filled down. A rep whose sales land exactly on a threshold earns that tier’s rate. Refer to the Table’s columns by name, not cell addresses, so the formula reads like the rate card and keeps up when the plan changes.',
  concept: {
    summary:
      'An approximate-match lookup finds the largest value that is at or below the one you’re looking for. That turns a threshold table into a rate card, so one lookup replaces a chain of nested IFs, and changing the plan means editing the table, not the formulas.',
    syntax: '=XLOOKUP(lookup_value, threshold_column, return_column, , -1)',
    example: '=XLOOKUP(E2, Tiers[Min sales], Tiers[Rate], , -1)',
    tip: 'VLOOKUP(E2, Tiers, 2, TRUE) does the same job, but only while Min sales stays sorted smallest to largest. XLOOKUP with match mode -1 doesn’t depend on the sort.',
  },
  hints: [
    'Each rep belongs to the highest tier whose Min sales is at or below their sales. That’s an approximate match, not an exact one.',
    'XLOOKUP’s fifth argument, match_mode, controls this: -1 means an exact match or else the next smaller value. Leave the fourth argument (if_not_found) empty.',
    '=XLOOKUP(E2, Tiers[Min sales], Tiers[Rate], , -1), then fill down to F9.',
  ],
  solution: () => '=XLOOKUP(E2,Tiers[Min sales],Tiers[Rate],,-1)',
  make: (rng) => {
    const tiers = makeTiers(rng);
    return { tiers, reps: makeSales(rng, tiers, rng.shuffle(REPS)) };
  },
  layout: (d) => ({
    blocks: [
      dataBlock('Tiers', 'A1', TIER_COLS, tierGrid(d.tiers)),
      cells('D1', [['Rep', 'Quarter sales', 'Commission rate']], 'header'),
      cells('D2', repGrid(d.reps), 'input', DOLLARS),
    ],
    answer: { kind: 'cells', range: `F2:F${d.reps.length + 1}`, format: FMT.pct, consistency: 'all' },
  }),
  expected: (d) => d.reps.map((r) => [tierRate(d.tiers, r.sales)]),
  inputs: (d) => [tableWrite('Tiers', TIER_COLS, tierGrid(d.tiers)), rangeWrite(`D2:E${d.reps.length + 1}`, repGrid(d.reps))],
  variants: [
    {
      label: 'the tier thresholds and rates change',
      explain: 'Read the thresholds and rates from the Tiers Table instead of typing them into the formula.',
      apply: (d, rng) => {
        let tiers = makeTiers(rng);
        while (JSON.stringify(tiers) === JSON.stringify(d.tiers)) tiers = makeTiers(rng);
        return { ...d, tiers };
      },
    },
    {
      label: 'the sales change',
      apply: (d, rng) => ({ ...d, reps: makeSales(rng, d.tiers, d.reps.map((r) => r.rep)) }),
    },
    {
      label: 'a new top tier is added',
      explain: 'Point the lookup at the whole Tiers columns so a new tier is included.',
      apply: (d, rng) => {
        const top = d.tiers[d.tiers.length - 1];
        const added: Tier = { min: top.min + rng.int(3, 6) * 5000, rate: round(top.rate + 0.01, 4) };
        const k = rng.int(0, d.reps.length - 1);
        const reps = d.reps.map((r, i) => (i === k ? { ...r, sales: added.min + rng.int(500, 20000) } : r));
        return { tiers: [...d.tiers, added], reps };
      },
    },
  ],
  rules: {
    require: [
      {
        pattern: /XLOOKUP\(|VLOOKUP\(|LOOKUP\(|INDEX\(/i,
        label: 'Uses an approximate-match lookup',
        advice: 'Look up the sales in Tiers[Min sales] with XLOOKUP and match_mode -1, which returns an exact match or else the next smaller value.',
      },
      {
        // A fixed range that exactly covers the Table grows with it too (Excel rewrites $A$2:$B$7 to
        // $A$2:$B$8), so the new-tier variant can't catch it. The skill is naming the columns.
        pattern: /\bTiers\b/i,
        label: 'Refers to the Tiers Table by name',
        advice: 'Use Tiers[Min sales] and Tiers[Rate] instead of cell addresses. Select the column in the Table while writing the formula and Excel types the name for you.',
      },
    ],
    allowNumbers: [-1, 0, 1, 2],
  },
});

// ---------- Look up a rate by row and column ----------

export const ZONES = ['Zone 1', 'Zone 2', 'Zone 3', 'Zone 4', 'Zone 5', 'Zone 6'] as const;
export const BANDS = ['Up to 50 kg', '51 to 150 kg', '151 to 500 kg', '501 to 1,000 kg', 'Over 1,000 kg'] as const;
const BAND_MULT = [1, 1.7, 3.1, 5.6, 9.4];

interface RateCard {
  /** Zone label for each row of `rates`, in sheet order. */
  zones: string[];
  /** rates[row][band]: flat rate per shipment. */
  rates: number[][];
}

interface ZoneShip {
  id: string;
  zone: string;
  band: string;
}

interface TwoWayData {
  card: RateCard;
  ships: ZoneShip[];
}

function makeCard(rng: Rng): RateCard {
  const rates = ZONES.map((_, z) => {
    const base = rng.float(12, 16, 2) + z * rng.float(2.5, 4.5, 2);
    return BAND_MULT.map((m) => round(base * m * rng.float(0.95, 1.05, 3), 2));
  });
  return { zones: [...ZONES], rates };
}

function makeZoneShips(rng: Rng): ZoneShip[] {
  return Array.from({ length: 20 }, (_, i) => ({ id: `SH-${40100 + i}`, zone: rng.pick(ZONES), band: rng.pick(BANDS) }));
}

/** The rate where the shipment's zone row and band column cross. */
export function cardRate(card: RateCard, zone: string, band: string): number {
  const row = card.zones.indexOf(zone);
  const col = (BANDS as readonly string[]).indexOf(band);
  if (row < 0 || col < 0) throw new Error(`No rate for ${zone} / ${band}`);
  return card.rates[row][col];
}

const cardGrid = (c: RateCard): Grid => c.zones.map((z, i) => [z, ...c.rates[i]]);
const zoneShipGrid = (rows: ZoneShip[]): Grid => rows.map((s) => [s.id, s.zone, s.band]);

export const lookupTwoWay = defineExercise<TwoWayData>({
  id: 'lookup-two-way',
  module: 'lookups',
  title: 'Look up a rate by row and column',
  replaces: 'Reading freight rates off a printed rate card, one shipment at a time',
  minutes: 6,
  task: () =>
    'The rate card in `A1:F7` lists a flat freight rate for each zone (down the left) and weight band (across the top). In `K2:K21`, return each shipment’s rate by matching its zone in column `I` to the card’s rows and its weight band in column `J` to the card’s columns. Write one formula in `K2` and fill it down.',
  concept: {
    summary:
      'A two-way lookup finds a row and a column, then returns the cell where they cross. MATCH gives a label’s position in a list, and INDEX returns the cell at a given row and column of a range.',
    syntax: '=INDEX(rates, MATCH(row_label, row_labels, 0), MATCH(column_label, column_labels, 0))',
    example: '=INDEX($B$2:$F$7, MATCH(I2, $A$2:$A$7, 0), MATCH(J2, $B$1:$F$1, 0))',
    tip: 'XLOOKUP can do it by nesting: the inner XLOOKUP returns the whole band column, the outer one picks the zone’s row. =XLOOKUP(I2, $A$2:$A$7, XLOOKUP(J2, $B$1:$F$1, $B$2:$F$7))',
  },
  hints: [
    'There are two questions: which row is the zone on, and which column is the weight band in? MATCH(I2, $A$2:$A$7, 0) answers the first.',
    'MATCH(J2, $B$1:$F$1, 0) finds the column. Give both positions to INDEX over the rates in $B$2:$F$7.',
    '=INDEX($B$2:$F$7, MATCH(I2, $A$2:$A$7, 0), MATCH(J2, $B$1:$F$1, 0)). Press {absKey} on each range to lock it, then fill down.',
  ],
  solution: () => '=INDEX($B$2:$F$7,MATCH(I2,$A$2:$A$7,0),MATCH(J2,$B$1:$F$1,0))',
  make: (rng) => ({ card: makeCard(rng), ships: makeZoneShips(rng) }),
  layout: (d) => ({
    blocks: [
      cells('A1', [['Zone', ...BANDS]], 'header'),
      cells('A2', cardGrid(d.card), 'input', FMT.currency),
      // Column G stays empty so the card reads as its own block.
      cells('H1', [['Ship ID', 'Zone', 'Weight band', 'Rate']], 'header'),
      cells('H2', zoneShipGrid(d.ships), 'input'),
    ],
    answer: { kind: 'cells', range: 'K2:K21', format: FMT.currency, consistency: 'all' },
  }),
  expected: (d) => d.ships.map((s) => [cardRate(d.card, s.zone, s.band)]),
  inputs: (d) => [rangeWrite('A2:F7', cardGrid(d.card)), rangeWrite('H2:J21', zoneShipGrid(d.ships))],
  variants: [
    {
      label: 'the rates on the card change',
      apply: (d, rng) => ({ ...d, card: { ...d.card, rates: d.card.rates.map((row) => row.map((r) => round(r * rng.float(0.85, 1.2, 3), 2))) } }),
    },
    {
      label: 'the zones on the card are re-sorted',
      explain: 'Match the zone label with MATCH or XLOOKUP instead of counting rows.',
      apply: (d, rng) => {
        const idx = d.card.zones.map((_, i) => i);
        let order = rng.shuffle(idx);
        while (order.join() === idx.join()) order = rng.shuffle(idx);
        return { ...d, card: { zones: order.map((i) => d.card.zones[i]), rates: order.map((i) => d.card.rates[i]) } };
      },
    },
    { label: 'the shipments change', apply: (d, rng) => ({ ...d, ships: makeZoneShips(rng) }) },
  ],
  rules: {
    require: [
      {
        pattern: /INDEX\(|XLOOKUP\(/i,
        label: 'Looks up by row and column',
        advice: 'Use INDEX with two MATCHes, or one XLOOKUP inside another.',
      },
    ],
    forbidText: { values: [...ZONES, ...BANDS], advice: 'Match against the zone and band in columns I and J instead.' },
    allowNumbers: [0],
  },
});

// ---------- Each SKU’s peak day in one formula ----------

const DAYS = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'] as const;

const DAILY_COLS: ColumnSpec[] = [{ header: 'SKU' }, ...DAYS.map((day) => ({ header: day, format: FMT.int }))];

interface DailyRow {
  sku: string;
  units: number[];
}

interface DailyData {
  rows: DailyRow[];
}

function dailyRow(rng: Rng, sku: string): DailyRow {
  const level = rng.int(20, 300);
  return { sku, units: DAYS.map(() => Math.max(0, Math.round(level * rng.float(0.3, 1.6, 2)))) };
}

function dailyRows(rng: Rng, count: number): DailyRow[] {
  const numbers = rng.sample(Array.from({ length: 400 }, (_, i) => 2000 + i * 7), count);
  return numbers.map((n) => dailyRow(rng, `SKU-${n}`));
}

const dailyGrid = (rows: DailyRow[]): Grid => rows.map((r) => [r.sku, ...r.units]);

export const byrowPeak = defineExercise<DailyData>({
  id: 'byrow-peak',
  module: 'dynamic',
  title: 'Each SKU’s peak day in one formula',
  replaces: 'A helper column of MAX formulas you extend by hand',
  minutes: 5,
  m365: true,
  task: () =>
    'The `DailyUnits` Table has one row per SKU and one column per day of the week. With one formula in `J2`, return each SKU’s peak: the most units it shipped on any single day. The result should spill one number per SKU, in the Table’s row order. Refer to the Table’s day columns by name, not cell addresses, so the formula reads clearly and keeps up as SKUs are added.',
  concept: {
    summary:
      'BYROW hands each row of a range to a small LAMBDA and spills one result per row. MAX over the whole grid gives a single number; BYROW gives one per row.',
    syntax: '=BYROW(array, LAMBDA(row, calculation))',
    example: '=BYROW(DailyUnits[[Mon]:[Sun]], LAMBDA(r, MAX(r)))',
    tip: 'DailyUnits[[Mon]:[Sun]] means the Mon through Sun columns of the Table; selecting those cells while you type the formula writes it for you. Recent Microsoft 365 builds also accept the function name alone: =BYROW(DailyUnits[[Mon]:[Sun]], MAX).',
  },
  hints: [
    'MAX(DailyUnits[[Mon]:[Sun]]) returns one number for the whole grid. You need one per row, which is what BYROW is for.',
    'BYROW’s second argument is a LAMBDA that receives one row at a time: LAMBDA(r, MAX(r)).',
    '=BYROW(DailyUnits[[Mon]:[Sun]], LAMBDA(r, MAX(r)))',
  ],
  solution: () => '=BYROW(DailyUnits[[Mon]:[Sun]],LAMBDA(r,MAX(r)))',
  make: (rng) => ({ rows: dailyRows(rng, 20) }),
  layout: (d) => ({
    blocks: [dataBlock('DailyUnits', 'A1', DAILY_COLS, dailyGrid(d.rows)), cells('J1', [['Peak daily units']], 'header')],
    answer: { kind: 'spill', anchor: 'J2', format: FMT.int },
  }),
  expected: (d) => d.rows.map((r) => [Math.max(...r.units)]),
  inputs: (d) => [tableWrite('DailyUnits', DAILY_COLS, dailyGrid(d.rows))],
  variants: [
    {
      label: 'the daily units change',
      apply: (d, rng) => ({ rows: d.rows.map((r) => dailyRow(rng, r.sku)) }),
    },
    {
      label: 'two SKUs are added',
      explain: 'Point BYROW at the Table’s day columns so new SKUs get a result.',
      apply: (d, rng) => {
        const used = new Set(d.rows.map((r) => r.sku));
        const added: DailyRow[] = [];
        while (added.length < 2) {
          const sku = `SKU-${rng.int(5000, 5999)}`;
          if (!used.has(sku)) {
            used.add(sku);
            added.push(dailyRow(rng, sku));
          }
        }
        return { rows: [...d.rows, ...added] };
      },
    },
  ],
  rules: {
    require: [
      { pattern: /BYROW\(/i, label: 'Uses BYROW', advice: 'BYROW(array, LAMBDA(r, MAX(r))) returns one result per row.' },
      {
        // A fixed range that exactly covers the Table grows with it, so the added-SKUs variant can't catch it.
        pattern: /DailyUnits\[/i,
        label: 'Refers to the DailyUnits Table’s columns by name',
        advice: 'Use DailyUnits[[Mon]:[Sun]] instead of cell addresses. Select the day columns in the Table while writing the formula and Excel types the name for you.',
      },
    ],
  },
});

// ---------- A cross-tab in one formula ----------

type SpendRow = ReturnType<typeof spendLine>;

interface CrossTabData {
  rows: SpendRow[];
}

/**
 * PIVOTBY's default labels. The total label reads "Total" in English Excel; the corner above the
 * row labels is blank when field headers aren't shown. Both are matched loosely.
 */
export const TOTAL_LABEL: CellMatcher = { match: /total/i, describe: 'a Total label' };
export const CORNER: CellMatcher = { match: /^[^0-9]*$/, describe: 'the blank corner cell' };

const sortedUnique = (values: readonly string[]) => [...new Set(values)].sort(excelTextCompare);

/** One line for every vendor and category pair, so no cell of the cross-tab is ever empty. */
function everyPair(rng: Rng, vendors: readonly string[]): SpendRow[] {
  return vendors.flatMap((v) => CATEGORIES.map((c) => ({ ...spendLine(rng, [v]), category: c })));
}

function crossTabRows(rng: Rng): SpendRow[] {
  const vendors = rng.sample(VENDORS, 5);
  return rng.shuffle([...everyPair(rng, vendors), ...Array.from({ length: 45 }, () => spendLine(rng, vendors))]);
}

/** What =PIVOTBY(row, col, values, SUM) spills with every optional argument left out. */
export function pivotBySum(rows: readonly { row: string; col: string; value: number }[]): ExpectedGrid {
  const rowKeys = sortedUnique(rows.map((r) => r.row));
  const colKeys = sortedUnique(rows.map((r) => r.col));
  const total = (keep: (r: (typeof rows)[number]) => boolean) => sum(rows.filter(keep).map((r) => r.value));
  const out: ExpectedGrid = [[CORNER, ...colKeys, TOTAL_LABEL]];
  for (const rk of rowKeys) {
    out.push([rk, ...colKeys.map((ck) => total((r) => r.row === rk && r.col === ck)), total((r) => r.row === rk)]);
  }
  out.push([TOTAL_LABEL, ...colKeys.map((ck) => total((r) => r.col === ck)), total(() => true)]);
  return out;
}

export const pivotbyVendorCategory = defineExercise<CrossTabData>({
  id: 'pivotby-vendor-category',
  module: 'dynamic',
  title: 'A cross-tab in one formula',
  replaces: 'Building and refreshing a PivotTable for a one-off cross-tab',
  minutes: 4,
  m365: true,
  task: () =>
    'With one formula in `F2`, build a cross-tab of the `Spend` Table: vendors down the side, categories across the top, and the total Amount where they meet. Keep PIVOTBY’s defaults: both lists sorted A to Z, a Total row at the bottom and a Total column on the right. Point it at the Table’s columns, not fixed ranges, so a new vendor gets its own row.',
  concept: {
    summary:
      'PIVOTBY builds a cross-tab as one formula: one field down the side, one across the top, and a function applied where they meet. By default it sorts both lists A to Z and adds a Total row and a Total column.',
    syntax: '=PIVOTBY(row_fields, col_fields, values, function)',
    example: '=PIVOTBY(Spend[Vendor], Spend[Category], Spend[Amount], SUM)',
    tip: 'Pass the function by name, without parentheses. Swap SUM for AVERAGE or COUNT and every cell changes meaning.',
  },
  hints: [
    'PIVOTBY takes four things in order: what goes down the side, what goes across the top, the numbers, and the function.',
    'Vendors go down the side (Spend[Vendor]) and categories across the top (Spend[Category]). The numbers are Spend[Amount].',
    '=PIVOTBY(Spend[Vendor], Spend[Category], Spend[Amount], SUM)',
  ],
  solution: () => '=PIVOTBY(Spend[Vendor],Spend[Category],Spend[Amount],SUM)',
  make: (rng) => ({ rows: crossTabRows(rng) }),
  layout: (d) => ({
    blocks: [dataBlock('Spend', 'A1', SPEND_COLS, spendGrid(d.rows)), cells('F1', [['Spend by vendor and category']], 'label')],
    answer: { kind: 'spill', anchor: 'F2', format: FMT.currency },
  }),
  expected: (d) => pivotBySum(d.rows.map((r) => ({ row: r.vendor, col: r.category, value: r.amount }))),
  inputs: (d) => [tableWrite('Spend', SPEND_COLS, spendGrid(d.rows))],
  variants: [
    {
      label: 'the amounts change',
      apply: (d, rng) => ({ rows: d.rows.map((r) => ({ ...r, amount: round(r.amount * rng.float(0.5, 1.8, 3), 2) })) }),
    },
    {
      label: 'a new vendor appears',
      explain: 'Point PIVOTBY at whole Table columns so a new vendor gets its own row.',
      apply: (d, rng) => ({ rows: [...d.rows, ...everyPair(rng, [rng.pick(EXTRA_VENDORS)])] }),
    },
  ],
  rules: {
    require: [
      {
        pattern: /PIVOTBY\(/i,
        label: 'Uses PIVOTBY',
        advice: 'PIVOTBY(Spend[Vendor], Spend[Category], Spend[Amount], SUM) builds the whole cross-tab.',
      },
    ],
  },
});

export const DEPTH = [xlookupTiered, lookupTwoWay, byrowPeak, pivotbyVendorCategory];

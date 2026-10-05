import { CATEGORIES, EXTRA_VENDORS, REPS, VENDORS, excelTextCompare, sum } from '../engine/data';
import { round, type Rng } from '../engine/rng';
import type { CellMatcher, ColumnSpec, ExpectedGrid, Grid } from '../engine/types';
import { numberToCol } from '../engine/address';
import { FMT, cells, dataBlock, defineExercise, rangeWrite, tableWrite } from './common';
import { TABLE_TYPING_TIP, cellList, checkStep, fillStep, money, part, raw, rowsWhere, typeStep } from './guides';
import { SPEND_COLS, spendGrid, spendLine } from './sumifs';

/** Whole dollars, for sales figures and thresholds. */
const DOLLARS = '$#,##0';

const wholeUsd = new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD', maximumFractionDigits: 0 });
/** $12,345, the way a DOLLARS cell shows. */
const dollars = (n: number) => wholeUsd.format(n);
/** 3.5%, the way an FMT.pct cell shows. */
const pct = (n: number) => `${(n * 100).toFixed(1)}%`;
/** 1,234, the way an FMT.int cell shows. */
const count = (n: number) => n.toLocaleString('en-US');

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
  guide: (d) => {
    const rep = d.reps[0];
    // The tier XLOOKUP(…, -1) lands on: the largest Min sales at or below the rep's sales.
    const at = d.tiers.reduce((best, t, i) => (t.min <= rep.sales && t.min >= d.tiers[best].min ? i : best), 0);
    const tier = d.tiers[at];
    const onLine = d.reps.find((r) => r !== rep && r.sales > 0 && d.tiers.some((t) => t.min === r.sales));
    const last = d.reps.length + 1;
    return [
      {
        do: `Meet the data. The rate card in \`A1:B${d.tiers.length + 1}\` is a Table named **Tiers**. Each row is a tier: the lowest sales that earn it (Min sales) and the rate it pays.`,
        why: 'Reading down, the thresholds climb and so do the rates. A rep earns the rate of the last tier they reach: the highest Min sales at or below their quarter sales.',
        show: [
          { label: 'Min sales column', at: 'Tiers[Min sales]', note: 'That’s `Tiers[Min sales]` (column `A`): the thresholds, smallest first.' },
          { label: 'Rate column', at: 'Tiers[Rate]', note: 'That’s `Tiers[Rate]` (column `B`): the rate each tier pays.' },
          { label: 'Quarter sales', at: `E2:E${last}`, note: 'Each rep’s sales for the quarter. Your formula looks each one up on the rate card.' },
        ],
      },
      {
        do: `See what \`F2\` should show. ${rep.rep} sold ${dollars(rep.sales)}.`,
        why:
          rep.sales === tier.min
            ? `${dollars(rep.sales)} is exactly a threshold, so ${rep.rep} earns that tier’s rate: ${pct(tier.rate)}.`
            : `No threshold equals ${dollars(rep.sales)}, so an exact match would find nothing. Read down Min sales and stop at the last one that isn’t above it: ${dollars(tier.min)}. That tier pays ${pct(tier.rate)}.`,
        show: [
          {
            label: `Select ${rep.rep}’s tier`,
            at: `A${at + 2}:B${at + 2}`,
            note: `${dollars(tier.min)} is the highest Min sales at or below ${dollars(rep.sales)}. Its rate, ${pct(tier.rate)}, is the number for \`F2\`.`,
          },
        ],
      },
      typeStep({
        cell: 'F2',
        formula: [
          part('=XLOOKUP(', 'Looks a value up in one column and returns the value on the same row of another.'),
          part('E2', `What to look up: this rep’s quarter sales in \`E2\` (${dollars(rep.sales)}).`, 'E2'),
          raw(', '),
          part('Tiers[Min sales]', 'Where to look: the thresholds.', 'Tiers[Min sales]'),
          raw(', '),
          part('Tiers[Rate]', 'What to return: the rate on the row it lands on.', 'Tiers[Rate]'),
          part(', ,', 'Two commas with nothing between them. The empty slot is the fourth argument, what to show when nothing is found. Leave it empty: every rep reaches at least the $0 tier.'),
          raw(' '),
          part('-1', `The match mode. \`-1\` means an exact match, or else the next smaller value. So ${dollars(rep.sales)} lands on ${dollars(tier.min)}.`),
          raw(')'),
        ],
        why: TABLE_TYPING_TIP,
      }),
      fillStep({
        from: 'F2',
        range: `F2:F${last}`,
        direction: 'down',
        why: `Excel moves \`E2\` down to each rep’s sales; the Tiers columns stay put.${
          onLine ? ` ${onLine.rep} sold exactly ${dollars(onLine.sales)}, a threshold, and gets that tier’s rate, because \`-1\` takes an exact match first.` : ''
        }`,
      }),
      checkStep('The coach changes the thresholds and rates, changes the sales and adds a new top tier behind the scenes, then puts everything back. Your formula reads the Tiers Table by column name, so it keeps up with each change.'),
    ];
  },
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
  guide: (d) => {
    const ship = d.ships[0];
    const zoneAt = d.card.zones.indexOf(ship.zone) + 1;
    const bandAt = (BANDS as readonly string[]).indexOf(ship.band) + 1;
    const col = numberToCol(1 + bandAt);
    const row = 1 + zoneAt;
    const rate = cardRate(d.card, ship.zone, ship.band);
    return [
      {
        do: 'Meet the layout. The rate card fills `A1:F7`: zones run down column `A`, weight bands run across row `1`, and each rate sits where a zone’s row meets a band’s column.',
        why: 'Each shipment in `H2:J21` names a zone (column `I`) and a weight band (column `J`). Its rate is the cell where that zone’s row and that band’s column cross on the card.',
        show: [
          { label: 'Zones', at: 'A2:A7', note: 'The zones: one row of the card each.' },
          { label: 'Weight bands', at: 'B1:F1', note: 'The weight bands: one column of the card each.' },
          { label: 'Rates', at: 'B2:F7', note: 'The rates. Your formula picks one of these for each shipment.' },
        ],
      },
      {
        do: `See what \`K2\` should show. Shipment ${ship.id} is ${ship.zone}, ${ship.band}.`,
        why: `${ship.zone} is number ${zoneAt} in the zone list, and ${ship.band} is number ${bandAt} in the band list. So the rate is in row ${zoneAt}, column ${bandAt} of the rates.`,
        show: [
          { label: `${ship.zone} row`, at: `A${row}:F${row}`, note: `${ship.zone} is number ${zoneAt} down the zone list.` },
          { label: `${ship.band} column`, at: `${col}1:${col}7`, note: `${ship.band} is number ${bandAt} across the band list.` },
          { label: 'Where they cross', at: `${col}${row}`, note: `\`${col}${row}\` holds ${money(rate)}. That’s the number for \`K2\`.` },
        ],
      },
      typeStep({
        cell: 'K2',
        formula: [
          part('=INDEX(', 'Returns the cell at a given row number and column number inside a block of cells.'),
          part('$B$2:$F$7', 'The block: the rates on the card. The `$` signs lock it, so it stays put when you fill down.', 'B2:F7'),
          raw(', '),
          part('MATCH(', 'Which row? MATCH gives the position of a value in a list: 1 for the first item, 2 for the second, and so on.'),
          part('I2', `The value to find: this shipment’s zone in \`I2\` (${ship.zone}).`, 'I2'),
          raw(', '),
          part('$A$2:$A$7', `The list to find it in: the zones down the card, locked with \`$\`. ${ship.zone} is number ${zoneAt}.`, 'A2:A7'),
          raw(', '),
          part('0', 'Exact match only.'),
          raw('), '),
          part('MATCH(', 'Which column? A second MATCH, this time across the top.'),
          part('J2', `The value to find: this shipment’s weight band in \`J2\` (${ship.band}).`, 'J2'),
          raw(', '),
          part('$B$1:$F$1', `The list to find it in: the bands across row \`1\`, locked with \`$\`. ${ship.band} is number ${bandAt}.`, 'B1:F1'),
          raw(', '),
          part('0', 'Exact match again.'),
          part('))', `Closes the second MATCH, then INDEX. INDEX now has row ${zoneAt} and column ${bandAt} of the rates: ${money(rate)}.`),
        ],
        why: 'Adding the `$`: after typing a range, press {absKey} until it reads `$B$2:$F$7`, and do the same for the two lists. Leave `I2` and `J2` without `$`, so they move down a row as you fill.',
      }),
      fillStep({
        from: 'K2',
        range: 'K2:K21',
        direction: 'down',
        why: 'Fill Down copies `K2`’s formula into the cells below. `I2` and `J2` move down to each shipment’s zone and band, and the three locked ranges stay on the card.',
      }),
      checkStep('The coach changes the rates, re-sorts the zones on the card and swaps the shipments behind the scenes, then puts everything back. MATCH finds each zone by its name wherever it sits, so your rates stay right.'),
    ];
  },
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
  guide: (d) => {
    const first = d.rows[0];
    const peak = Math.max(...first.units);
    const peakDay = first.units.indexOf(peak);
    const last = d.rows.length + 1;
    const days = `B2:H${last}`;
    const gridMax = Math.max(...d.rows.flatMap((r) => r.units));
    return [
      {
        do: 'Meet the data. The blue block is a Table named **DailyUnits**: one row per SKU, and one column per day from Mon (column `B`) to Sun (column `H`).',
        why: 'In a formula, `DailyUnits[[Mon]:[Sun]]` means the Table’s Mon through Sun columns together: every number in the grid, without the SKU column.',
        show: [{ label: 'Mon to Sun columns', at: days, note: 'That’s `DailyUnits[[Mon]:[Sun]]`: the units each SKU shipped on each day.' }],
      },
      {
        do: `See what \`J2\` should show: the largest of ${first.sku}’s seven daily numbers.`,
        why: `\`MAX\` over the whole grid would give a single number, ${count(gridMax)}, for all the SKUs together. You want one peak per row, and that’s what BYROW is for.`,
        show: [
          { label: `Select ${first.sku}’s week`, at: 'B2:H2', note: `${first.sku}’s seven days, Mon to Sun. The biggest is ${count(peak)}.` },
          {
            label: 'Select its peak day',
            at: `${numberToCol(2 + peakDay)}2`,
            note: `${DAYS[peakDay]}: ${count(peak)} units. That’s the number for \`J2\`.`,
          },
        ],
      },
      typeStep({
        cell: 'J2',
        whole: true,
        formula: [
          part('=BYROW(', 'Runs a calculation on each row in turn and gives back one result per row.'),
          part('DailyUnits[[Mon]:[Sun]]', 'The rows to work through: the Table’s Mon to Sun columns.', days),
          raw(', '),
          part('LAMBDA(', 'The calculation to run on each row. LAMBDA writes a small formula with a name for its input.'),
          part('r', `The input’s name. Each time round, \`r\` stands for one row: first ${first.sku}’s seven numbers in \`B2:H2\`, then the next row, and so on.`, 'B2:H2'),
          raw(', '),
          part('MAX(r)', 'What to work out for each row: its largest number.'),
          part('))', 'Closes LAMBDA, then BYROW.'),
        ],
        why: `Typing tip: type \`=BYROW(\`, then drag across \`${days}\`, and Excel writes \`DailyUnits[[Mon]:[Sun]]\` for you. One formula is enough: the peaks spill, which means Excel writes them into \`J2\` and the cells below, one per SKU (\`J2:J${last}\`). Keep those cells empty: a #SPILL! error means something is in the way.`,
      }),
      checkStep('The coach changes the daily units and adds two SKUs to the Table behind the scenes, then puts everything back. BYROW reads the Table’s columns, so the new SKUs get a peak too.'),
    ];
  },
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
  guide: (d) => {
    const vendors = sortedUnique(d.rows.map((r) => r.vendor));
    const cats = sortedUnique(d.rows.map((r) => r.category));
    const [vendor, cat] = [vendors[0], cats[0]];
    const keep = (r: SpendRow) => r.vendor === vendor && r.category === cat;
    const total = sum(d.rows.filter(keep).map((r) => r.amount));
    // Row labels in F, one column per category from G, then the Total column; header row 2, Total row last.
    const area = `F2:${numberToCol(7 + cats.length)}${3 + vendors.length}`;
    return [
      {
        do: 'Meet the data. The blue block in columns `A` to `D` is a Table named **Spend**. Each row is one purchase from a vendor, in a category.',
        why: 'A cross-tab puts one field down the side and another across the top, with a total where each row meets each column. Here: vendors down the side, categories across the top, and Amount added up in the middle.',
        show: [
          { label: 'Vendor column', at: 'Spend[Vendor]', note: `That’s \`Spend[Vendor]\` (column \`B\`): ${vendors.length} vendors, one row each in the cross-tab.` },
          { label: 'Category column', at: 'Spend[Category]', note: `That’s \`Spend[Category]\` (column \`C\`): ${cats.length} categories, one column each.` },
          { label: 'Amount column', at: 'Spend[Amount]', note: 'That’s `Spend[Amount]` (column `D`): the numbers to add up.' },
        ],
      },
      {
        do: `See the shape of the result: ${vendors.length} vendors down the side and ${cats.length} categories across the top, both A to Z, plus a Total row and a Total column.`,
        why: `It will fill \`${area}\`. The category names run across row \`2\` from \`G2\`, and the vendor names run down column \`F\` from \`F3\`. So \`G3\` is ${vendor} in ${cat}.`,
        show: d.rows.some(keep)
          ? [
              {
                label: `Select ${vendor} in ${cat}`,
                at: cellList('D', rowsWhere(d.rows, keep)),
                note: `The status bar’s **Sum** is ${money(total)}. That’s where ${vendor}’s row meets the ${cat} column: \`G3\`, the first number in the cross-tab.`,
              },
            ]
          : undefined,
      },
      typeStep({
        cell: 'F2',
        whole: true,
        formula: [
          part('=PIVOTBY(', 'Builds a cross-tab: one field down the side, one across the top, and a summary where they meet.'),
          part('Spend[Vendor]', 'Down the side: one row per vendor.', 'Spend[Vendor]'),
          raw(', '),
          part('Spend[Category]', 'Across the top: one column per category.', 'Spend[Category]'),
          raw(', '),
          part('Spend[Amount]', 'The numbers to summarize.', 'Spend[Amount]'),
          raw(', '),
          part('SUM', 'How to combine the amounts where a row meets a column: add them. Type the name alone, with no bracket after it. If you pick SUM from Excel’s list and it adds `(`, delete the `(`.'),
          raw(')'),
        ],
        why: `${TABLE_TYPING_TIP} The cross-tab spills, which means Excel writes it into \`F2\` and the cells below and to the right, \`${area}\` today. Keep that area empty: a #SPILL! error means something is in the way.`,
      }),
      checkStep('The coach changes the amounts and adds a new vendor behind the scenes, then puts everything back. PIVOTBY reads whole Table columns, so the new vendor gets its own row.'),
    ];
  },
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

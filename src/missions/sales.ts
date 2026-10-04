import { EXTRA_REPS, REPS, eomonth, excelTextCompare, fromSerial, serial, sum } from '../engine/data';
import { round, type Rng } from '../engine/rng';
import type { AnswerArea, ColumnSpec, Grid, Inspection, Variant } from '../engine/types';
import { FMT, cells, column, dataBlock, rangeWrite, tableWrite } from '../exercises/common';
import { defineMission, type Mission } from './types';

const MONTH_NAMES = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];

function monthName(s: number): string {
  const { year, month } = fromSerial(s);
  return `${MONTH_NAMES[month - 1]} ${year}`;
}

/** Whole dollars, for deal amounts, quotas and pipeline totals. */
const DOLLARS = '$#,##0';
/** A ratio read as a multiple: 1.50x. */
const TIMES = '0.00"x"';

export interface RepQuota {
  rep: string;
  quota: number;
}

/** A shuffle that is guaranteed to change the order (items are compared by `key`). */
function reshuffle<T>(rng: Rng, items: readonly T[], key: (x: T) => string = String): T[] {
  const order = (xs: readonly T[]) => xs.map(key).join('|');
  let next = rng.shuffle(items);
  while (order(next) === order(items)) next = rng.shuffle(items);
  return next;
}

const ALL_REPS = [...REPS, ...EXTRA_REPS];
const REP_TEXT = { values: ALL_REPS, advice: 'Match against the rep on the same row instead of typing a name.' };

// =====================================================================
// Monthly commission statement
// =====================================================================

export const COMMISSION_MONTHS = [serial(2026, 7, 1), serial(2026, 8, 1), serial(2026, 9, 1)];

/** Tier thresholds in whole percent of quota. The plan's rates change by seed; the thresholds don't. */
export const TIER_MINS_PCT = [0, 50, 70, 100, 120];

/**
 * Attainment levels the data is built around: every tier threshold and every flag level the
 * variants use. A rep's attainment sits exactly on one of them (the edge cases the plan spells
 * out) or at least a full point away, so nothing reads 70% on screen while sitting a hair below it.
 */
export const LEVELS_PCT = [50, 60, 70, 75, 80, 100, 120];

/** Plan defaults written to H2 and H3; the variants pick from the alternatives. */
export const PLAN_ACCELERATOR = 1.5;
export const PLAN_FLAG_BELOW = 0.7;
const ACCELERATORS = [1.25, 1.75, 2];
export const FLAG_LEVELS = [0.75, 0.8, 0.6];

export const AT_RISK = 'At risk';

export type DealStatus = 'Won' | 'Cancelled';

export interface Deal {
  id: string;
  rep: string;
  /** Close date. A cancelled deal carries the date it was cancelled. */
  date: number;
  amount: number;
  status: DealStatus;
}

export interface PlanTier {
  /** Threshold as a fraction of quota (0.7 = 70%). */
  min: number;
  rate: number;
}

export interface CommissionData {
  reps: RepQuota[];
  deals: Deal[];
  tiers: PlanTier[];
  /** First day of the statement month (H1). */
  month: number;
  /** Multiple of the rate paid on bookings above quota (H2). */
  accelerator: number;
  /** Attainment below this is flagged (H3). */
  flagBelow: number;
}

const DEAL_COLS: ColumnSpec[] = [
  { header: 'Deal' },
  { header: 'Rep' },
  { header: 'Close date', format: FMT.date },
  { header: 'Amount', format: DOLLARS },
  { header: 'Status' },
];
const TIER_COLS: ColumnSpec[] = [
  { header: 'Min attainment', format: '0%' },
  { header: 'Rate', format: '0%' },
];
/** Quota is the last column, so the Bookings column the learner adds picks up its dollar format. */
const REP_COLS: ColumnSpec[] = [{ header: 'Rep' }, { header: 'Quota', format: DOLLARS }];

const dealGrid = (deals: Deal[]): Grid => deals.map((x) => [x.id, x.rep, x.date, x.amount, x.status]);
const tierGrid = (tiers: PlanTier[]): Grid => tiers.map((t) => [t.min, t.rate]);
const repGrid = (reps: RepQuota[]): Grid => reps.map((r) => [r.rep, r.quota]);
const dealId = (n: number) => `D-${24101 + n}`;

const inMonth = (date: number, month: number) => date >= month && date <= eomonth(month);

/** Monthly quotas in steps of $5,000, so 50%, 70% and 120% of quota are whole dollars. */
const QUOTAS = Array.from({ length: 13 }, (_, k) => (12 + k) * 5000);
const newQuota = (rng: Rng) => rng.pick(QUOTAS);

/** Rates in whole percent, climbing two or three points a tier, so a one-point change keeps them in order. */
function makePlanTiers(rng: Rng): PlanTier[] {
  let rate = rng.pick([0.03, 0.04]);
  return TIER_MINS_PCT.map((pct, i) => {
    if (i > 0) rate = round(rate + rng.int(2, 3) / 100, 4);
    return { min: pct / 100, rate };
  });
}

/**
 * True when bookings ÷ quota is at least a point away from every level in LEVELS_PCT, or with
 * `orOnLevel`, exactly on one. Done in whole numbers: 100 × bookings against level × quota.
 */
export function clearOfLevels(bookings: number, quota: number, orOnLevel = false): boolean {
  return LEVELS_PCT.every((p) => {
    const gap = Math.abs(100 * bookings - p * quota);
    return gap >= quota || (orOnLevel && gap === 0);
  });
}

type Role = 'edge70' | 'edge' | 'over' | 'under' | 'any';

const ROLE_RANGE: Record<'over' | 'under' | 'any', [number, number]> = {
  over: [1.05, 1.6],
  under: [0.3, 0.62],
  any: [0.35, 1.5],
};

/**
 * Each month, one rep lands exactly on 70% of quota (the tier edge and the flag edge at once),
 * one exactly on another threshold, one over quota so the accelerator pays, and one well under 70%.
 */
function monthRoles(rng: Rng, count: number): Role[] {
  const roles: Role[] = Array.from({ length: count }, () => 'any');
  const [edge70, edge, over, under] = rng.sample([...roles.keys()], 4);
  roles[edge70] = 'edge70';
  roles[edge] = 'edge';
  roles[over] = 'over';
  roles[under] = 'under';
  return roles;
}

/** Whole-dollar bookings (in $50 steps) for a rep's role, clear of every level unless it's an edge. */
function targetBookings(rng: Rng, quota: number, role: Role): number {
  if (role === 'edge70') return (70 * quota) / 100;
  if (role === 'edge') return (rng.pick([50, 100, 120]) * quota) / 100;
  const [lo, hi] = ROLE_RANGE[role];
  const at = (share: number) => Math.round((share * quota) / 50) * 50;
  for (let i = 0; i < 200; i++) {
    const b = at(rng.float(lo, hi, 3));
    if (clearOfLevels(b, quota)) return b;
  }
  // Deterministic fallback: walk the range in half-point steps.
  for (let share = lo; share <= hi; share += 0.005) if (clearOfLevels(at(share), quota)) return at(share);
  return at(lo);
}

/** Splits a total into `parts` amounts in $50 steps, each at least $1,000, that add up to it exactly. */
function splitAmount(rng: Rng, total: number, parts: number): number[] {
  for (let i = 0; i < 50; i++) {
    const weights = Array.from({ length: parts }, () => rng.float(0.4, 1, 3));
    const w = sum(weights);
    const out = weights.slice(0, -1).map((x) => Math.round((total * x) / w / 50) * 50);
    out.push(total - sum(out));
    if (out.every((v) => v >= 1000)) return out;
  }
  const even = Math.floor(total / parts / 50) * 50;
  return [...Array.from({ length: parts - 1 }, () => even), total - even * (parts - 1)];
}

/**
 * Prices one rep's month in place: Won deals add up to the role's bookings, and each cancellation
 * is 6–30% of them, so a clawback never wipes out the month.
 */
function priceRepMonth(rng: Rng, deals: Deal[], rep: RepQuota, month: number, role: Role) {
  const mine = deals.filter((x) => x.rep === rep.rep && inMonth(x.date, month));
  const won = mine.filter((x) => x.status === 'Won');
  const total = targetBookings(rng, rep.quota, role);
  splitAmount(rng, total, won.length).forEach((amount, i) => {
    won[i].amount = amount;
  });
  for (const c of mine.filter((x) => x.status === 'Cancelled')) c.amount = Math.max(1000, Math.round((total * rng.float(0.06, 0.3, 3)) / 50) * 50);
}

/** Sets every amount. Reps, dates and statuses stay put. */
function priceDeals(rng: Rng, reps: RepQuota[], deals: Deal[]): Deal[] {
  const out = deals.map((x) => ({ ...x }));
  for (const m of COMMISSION_MONTHS) {
    const roles = monthRoles(rng, reps.length);
    reps.forEach((r, i) => priceRepMonth(rng, out, r, m, roles[i]));
  }
  return out;
}

const wonDeal = (rng: Rng, rep: string, month: number): Deal => ({ id: '', rep, date: rng.int(month, eomonth(month)), amount: 0, status: 'Won' });

/**
 * Two to four wins per rep per month and three cancellations. Each month has a win on its first
 * day and one on its last day, so a window built with > or < instead of >= and <= gives a wrong total.
 */
function dealSkeleton(rng: Rng, reps: RepQuota[]): Deal[] {
  const deals: Deal[] = [];
  for (const m of COMMISSION_MONTHS) {
    const month = reps.flatMap((r) => Array.from({ length: rng.int(2, 4) }, () => wonDeal(rng, r.rep, m)));
    const [first, last] = rng.sample(month, 2);
    first.date = m;
    last.date = eomonth(m);
    const cancelled = rng.sample(reps, 3).map((r): Deal => ({ id: '', rep: r.rep, date: rng.int(m, eomonth(m)), amount: 0, status: 'Cancelled' }));
    deals.push(...month, ...cancelled);
  }
  return deals.sort((a, b) => a.date - b.date).map((x, i) => ({ ...x, id: dealId(i) }));
}

/** What SUMIFS returns: the rep's deals with `status` that closed in the month in H1. */
export function repTotal(d: CommissionData, rep: string, status: DealStatus): number {
  return sum(d.deals.filter((x) => x.rep === rep && x.status === status && inMonth(x.date, d.month)).map((x) => x.amount));
}

/**
 * The rate of the highest tier whose threshold is at or below attainment: what XLOOKUP(…, -1)
 * returns. Compared in whole numbers (100 × bookings against threshold × quota). Excel divides
 * first, but with whole-dollar bookings and quotas the quotient is either exactly a threshold
 * (the same double as the typed 0.7) or far more than rounding error away from it, so it lands
 * on the same side.
 */
export function tierRateFor(tiers: readonly PlanTier[], bookings: number, quota: number): number {
  let best: PlanTier | undefined;
  for (const t of tiers) if (100 * bookings >= Math.round(t.min * 100) * quota && (!best || t.min > best.min)) best = t;
  if (!best) throw new Error(`Attainment ${bookings}/${quota} falls below the first tier`);
  return best.rate;
}

/** Rate on bookings up to quota, rate × accelerator above it, minus the rate on cancellations. */
export function commissionFor(rate: number, bookings: number, quota: number, cancelled: number, accelerator: number): number {
  return rate * (Math.min(bookings, quota) + accelerator * Math.max(0, bookings - quota) - cancelled);
}

export interface StatementLine {
  bookings: number;
  attainment: number;
  rate: number;
  cancelled: number;
  commission: number;
  flag: string;
}

export function statement(d: CommissionData): StatementLine[] {
  return d.reps.map((r) => {
    const bookings = repTotal(d, r.rep, 'Won');
    const cancelled = repTotal(d, r.rep, 'Cancelled');
    const rate = tierRateFor(d.tiers, bookings, r.quota);
    return {
      bookings,
      attainment: bookings / r.quota,
      rate,
      cancelled,
      commission: commissionFor(rate, bookings, r.quota, cancelled, d.accelerator),
      flag: 100 * bookings < Math.round(d.flagBelow * 100) * r.quota ? AT_RISK : '',
    };
  });
}

const repsColumn = (name: string, format?: string): AnswerArea => ({ kind: 'tableColumn', table: 'Reps', column: name, format });
const hasRepsColumn = (name: string): Inspection[] => [{ kind: 'tableColumn', table: 'Reps', column: name, label: `The Reps Table has a ${name} column` }];

const commissionMonthChanges: Variant<CommissionData> = {
  label: 'the month in H1 changes',
  explain: 'Build the date window from $H$1 and EOMONTH instead of typing dates.',
  apply: (d, rng) => ({ ...d, month: rng.pick(COMMISSION_MONTHS.filter((m) => m !== d.month)) }),
};

const dealAmountsChange: Variant<CommissionData> = {
  label: 'deal amounts change',
  apply: (d, rng) => ({ ...d, deals: priceDeals(rng, d.reps, d.deals) }),
};

const dealsResorted: Variant<CommissionData> = {
  label: 'the Deals Table is re-sorted',
  explain: 'Total whole Table columns with SUMIFS so the order of the rows doesn’t matter.',
  apply: (d, rng) => ({ ...d, deals: reshuffle(rng, d.deals, (x) => x.id) }),
};

/** A new rep with a quota and a few wins in every month. Their row has to fill by itself. */
function addRep(d: CommissionData, rng: Rng): CommissionData {
  const name = rng.pick(EXTRA_REPS.filter((n) => !d.reps.some((r) => r.rep === n)));
  const rep = { rep: name, quota: newQuota(rng) };
  const next = Math.max(...d.deals.map((x) => Number(x.id.slice(2)) - 24101)) + 1;
  const added = COMMISSION_MONTHS.flatMap((m) => Array.from({ length: rng.int(2, 4) }, () => wonDeal(rng, name, m)))
    .sort((a, b) => a.date - b.date)
    .map((x, i) => ({ ...x, id: dealId(next + i) }));
  for (const m of COMMISSION_MONTHS) priceRepMonth(rng, added, rep, m, 'any');
  return { ...d, reps: [...d.reps, rep], deals: [...d.deals, ...added] };
}

const repAdded: Variant<CommissionData> = {
  label: 'a new rep joins the Reps Table',
  explain: 'A calculated column fills new rows by itself. Formulas typed cell by cell don’t.',
  apply: addRep,
};

const quotasChange: Variant<CommissionData> = {
  label: 'quotas change',
  explain: 'Divide by [@Quota] on the same row instead of typing a quota.',
  apply: (d, rng) => ({
    ...d,
    reps: d.reps.map((r, i) => {
      const bookings = repTotal(d, r.rep, 'Won');
      // Any quota that keeps attainment readable. The first rep's has to move, so the variant is never a no-op.
      const [quota] = rng.shuffle(QUOTAS).filter((q) => clearOfLevels(bookings, q, true) && (i > 0 || q !== r.quota));
      return quota ? { ...r, quota } : r;
    }),
  }),
};

/**
 * Moves one tier's rate a point up or down, choosing a tier at least one rep earns this month.
 * Tiers are two or more points apart, so the rates stay in order.
 */
const tierRateChanges: Variant<CommissionData> = {
  label: 'a tier’s rate changes',
  explain: 'Return the rate from Tiers[Rate] instead of typing rates into the formula.',
  apply: (d, rng) => {
    const used = [...new Set(d.reps.map((r) => tierRateFor(d.tiers, repTotal(d, r.rep, 'Won'), r.quota)))];
    const target = rng.pick(used);
    const step = rng.chance(0.5) ? 0.01 : -0.01;
    return { ...d, tiers: d.tiers.map((t) => (t.rate === target ? { ...t, rate: round(t.rate + step, 4) } : t)) };
  },
};

const acceleratorChanges: Variant<CommissionData> = {
  label: 'the accelerator in H2 changes',
  explain: 'Multiply by $H$2 instead of typing the accelerator.',
  apply: (d, rng) => ({ ...d, accelerator: rng.pick(ACCELERATORS) }),
};

const flagLevelChanges: Variant<CommissionData> = {
  label: 'the flag level in H3 changes',
  explain: 'Compare with $H$3 instead of typing the level.',
  apply: (d, rng) => {
    // Prefer a level that changes at least one flag, so a typed-in level can't pass by luck.
    const current = statement(d).map((s) => s.flag).join('|');
    const levels = rng.shuffle(FLAG_LEVELS);
    const pick = levels.find((l) => statement({ ...d, flagBelow: l }).map((s) => s.flag).join('|') !== current) ?? levels[0];
    return { ...d, flagBelow: pick };
  },
};

export const commissionStatement = defineMission<CommissionData>({
  id: 'mission-commission-statement',
  title: 'Monthly commission statement',
  role: 'sales',
  summary: 'The monthly commission statement: bookings, attainment, tiered rates, accelerators and clawbacks',
  minutes: 25,
  skills: ['tables-calc-column', 'sumifs-month', 'xlookup-tiered', 'let-reorder'],
  brief: (d) => ({
    from: 'Tessa Moreno, Sales operations lead',
    subject: `Commission statement for ${monthName(d.month)}`,
    body: [
      'Hi,',
      '',
      `Payroll needs the ${monthName(d.month)} commission statement by Wednesday. The \`Deals\` Table is the CRM export for July through September, \`Reps\` has each rep’s monthly quota, and \`Tiers\` is this year’s rate card. Add the statement as new columns of \`Reps\`, starting in \`L1\`, and build it off the month in \`H1\` so I can rerun any month.`,
      '',
      '1. `Bookings`: each rep’s Won deals that closed in the month.',
      '2. `Attainment`: bookings as a share of quota.',
      '3. `Rate`: the rep’s rate from `Tiers`. Landing exactly on a threshold earns that tier.',
      '4. `Commission`: the rate on bookings up to quota, the rate × the accelerator in `H2` on anything above quota, minus the clawback.',
      '5. `Flag`: `At risk` for anyone below the level in `H3`.',
      '',
      'Cancelled deals carry the date they were cancelled. They don’t change attainment, but we claw back their commission at the rep’s rate for the month.',
      '',
      'Thanks,',
      'Tessa',
    ].join('\n'),
  }),
  make: (rng) => {
    const reps = rng
      .sample(REPS, 7)
      .sort(excelTextCompare)
      .map((rep) => ({ rep, quota: newQuota(rng) }));
    const deals = priceDeals(rng, reps, dealSkeleton(rng, reps));
    return { reps, deals, tiers: makePlanTiers(rng), month: rng.pick(COMMISSION_MONTHS), accelerator: PLAN_ACCELERATOR, flagBelow: PLAN_FLAG_BELOW };
  },
  // Reps sits at J1 with nothing in L:Q, so the five columns the learner adds (L:P) and the gap
  // after them stay clear. Column F and I and row 11 keep the other Tables apart.
  blocks: (d) => [
    dataBlock('Deals', 'A1', DEAL_COLS, dealGrid(d.deals)),
    cells('G1', [['Month'], ['Accelerator'], ['Flag below']], 'label'),
    cells('H1', [[d.month]], 'input', FMT.month),
    cells('H2', [[d.accelerator]], 'input', TIMES),
    cells('H3', [[d.flagBelow]], 'input', '0%'),
    dataBlock('Tiers', 'G5', TIER_COLS, tierGrid(d.tiers)),
    dataBlock('Reps', 'J1', REP_COLS, repGrid(d.reps)),
  ],
  inputs: (d) => [
    tableWrite('Deals', DEAL_COLS, dealGrid(d.deals)),
    tableWrite('Tiers', TIER_COLS, tierGrid(d.tiers)),
    tableWrite('Reps', REP_COLS, repGrid(d.reps)),
    rangeWrite('H1:H3', [[d.month], [d.accelerator], [d.flagBelow]]),
  ],
  steps: [
    // ---------- 1. SUMIFS with a status and a date window ----------
    {
      title: 'Bookings for the month',
      task: () =>
        'Add a column named `Bookings` to the `Reps` Table by typing the name in `L1`. In it, total each rep’s Won deals from the `Deals` Table that closed within the month in `H1` (`H1` holds the first day of that month). Cancelled deals don’t count toward bookings.',
      hints: [
        'Type Bookings in `L1` and press {enter}; the Table grows to take in the new column, and the formula you write in `L2` fills the whole column.',
        'Use SUMIFS with four conditions: Deals[Rep] matches [@Rep], Deals[Status] is "Won", and Deals[Close date] is on or after $H$1 and on or before EOMONTH($H$1,0). Press {absKey} to lock `H1`.',
        '`=SUMIFS(Deals[Amount], Deals[Rep], [@Rep], Deals[Status], "Won", Deals[Close date], ">="&$H$1, Deals[Close date], "<="&EOMONTH($H$1,0))`',
      ],
      solution: () => '=SUMIFS(Deals[Amount],Deals[Rep],[@Rep],Deals[Status],"Won",Deals[Close date],">="&$H$1,Deals[Close date],"<="&EOMONTH($H$1,0))',
      answer: () => repsColumn('Bookings', DOLLARS),
      expected: (d) => statement(d).map((s) => [s.bookings]),
      variants: [commissionMonthChanges, dealAmountsChange, dealsResorted, repAdded],
      rules: { forbidText: REP_TEXT, allowNumbers: [0, 1] },
      inspections: () => hasRepsColumn('Bookings'),
    },

    // ---------- 2. Ratio on the same row ----------
    {
      title: 'Quota attainment',
      task: () =>
        'Add an `Attainment` column in `M1` that divides each rep’s bookings by their quota. The value is what’s checked, so 85% of quota is 0.85. The new column picks up the dollar format of the one beside it, so format it as a percentage to make it read that way.',
      hints: [
        'Type Attainment in `M1`. In `M2`, type = and click the cells on the same row, so Excel writes row references like [@Bookings].',
        '`=[@Bookings]/[@Quota]`, then Home › Percent Style (the % button).',
      ],
      solution: () => '=[@Bookings]/[@Quota]',
      answer: () => repsColumn('Attainment', FMT.pct),
      expected: (d) => statement(d).map((s) => [s.attainment]),
      variants: [dealAmountsChange, quotasChange, repAdded],
      rules: { allowNumbers: [] },
      inspections: () => hasRepsColumn('Attainment'),
    },

    // ---------- 3. Approximate-match lookup on a rate card ----------
    {
      title: 'Commission rate from the tiers',
      task: () =>
        'Add a `Rate` column in `N1` with each rep’s commission rate from the `Tiers` Table: the rate of the highest tier whose `Min attainment` is at or below the rep’s attainment. A rep exactly on a threshold earns that tier’s rate. Use one approximate-match lookup on the Table’s columns, not nested IFs.',
      hints: [
        'This is an approximate match: find the largest Min attainment that is at or below [@Attainment].',
        'XLOOKUP’s fifth argument, match_mode, does this: -1 means an exact match or else the next smaller value. Leave the fourth argument (if_not_found) empty.',
        '`=XLOOKUP([@Attainment], Tiers[Min attainment], Tiers[Rate], , -1)`',
      ],
      solution: () => '=XLOOKUP([@Attainment],Tiers[Min attainment],Tiers[Rate],,-1)',
      answer: () => repsColumn('Rate', '0%'),
      expected: (d) => statement(d).map((s) => [s.rate]),
      variants: [tierRateChanges, dealAmountsChange, quotasChange, repAdded],
      rules: {
        require: [
          {
            pattern: /XLOOKUP\(|VLOOKUP\(|LOOKUP\(|INDEX\(/i,
            label: 'Uses an approximate-match lookup',
            advice: 'Look up [@Attainment] in Tiers[Min attainment] with XLOOKUP and match_mode -1, which returns an exact match or else the next smaller value.',
          },
        ],
        allowNumbers: [-1, 0, 1, 2],
      },
      inspections: () => hasRepsColumn('Rate'),
    },

    // ---------- 4. Accelerator and clawback ----------
    {
      title: 'Commission with accelerator and clawbacks',
      task: () =>
        'Add a `Commission` column in `O1`. Bookings up to quota earn the rep’s rate, and bookings above quota earn the rate × the accelerator in `H2`. Then subtract the clawback: the rep’s Cancelled deals for the month in `H1`, times the same rate. Format the column as currency.',
      hints: [
        'Split bookings at quota: MIN([@Bookings],[@Quota]) is the part up to quota, and MAX(0,[@Bookings]-[@Quota]) is the part above it, which is 0 for anyone under quota.',
        'The clawback is the same SUMIFS as Bookings with "Cancelled" in place of "Won". LET keeps the formula readable: name that total cancelled, then use the name in the final calculation.',
        '`=LET(cancelled, SUMIFS(Deals[Amount], Deals[Rep], [@Rep], Deals[Status], "Cancelled", Deals[Close date], ">="&$H$1, Deals[Close date], "<="&EOMONTH($H$1,0)), [@Rate]*(MIN([@Bookings],[@Quota]) + $H$2*MAX(0,[@Bookings]-[@Quota]) - cancelled))`, then Home › Accounting Number Format (the $ button).',
      ],
      solution: () =>
        '=LET(cancelled,SUMIFS(Deals[Amount],Deals[Rep],[@Rep],Deals[Status],"Cancelled",Deals[Close date],">="&$H$1,Deals[Close date],"<="&EOMONTH($H$1,0)),[@Rate]*(MIN([@Bookings],[@Quota])+$H$2*MAX(0,[@Bookings]-[@Quota])-cancelled))',
      answer: () => repsColumn('Commission', FMT.currency),
      expected: (d) => statement(d).map((s) => [s.commission]),
      variants: [
        { ...commissionMonthChanges, explain: 'Build the clawback’s date window from $H$1 too, so it follows the month like Bookings does.' },
        dealAmountsChange,
        tierRateChanges,
        acceleratorChanges,
        repAdded,
      ],
      rules: { forbidText: REP_TEXT, allowNumbers: [0, 1] },
      inspections: () => hasRepsColumn('Commission'),
    },

    // ---------- 5. Flag ----------
    {
      title: 'Flag reps at risk',
      task: () =>
        'Add a `Flag` column in `P1` that shows `At risk` when a rep’s attainment is below the level in `H3`, and an empty string ("") otherwise. A rep exactly at the level isn’t at risk. Point to `H3` instead of typing the level.',
      hints: [
        '“Below” is a strict test: use < so a rep exactly at the level isn’t flagged. Press {absKey} to lock `H3`.',
        '`=IF([@Attainment]<$H$3, "At risk", "")`',
      ],
      solution: () => '=IF([@Attainment]<$H$3,"At risk","")',
      answer: () => repsColumn('Flag'),
      expected: (d) => statement(d).map((s) => [s.flag]),
      variants: [
        { ...commissionMonthChanges, explain: 'Build the flag from [@Attainment], so it follows the bookings when the month changes.' },
        flagLevelChanges,
        quotasChange,
        repAdded,
      ],
      rules: { allowNumbers: [] },
      inspections: () => hasRepsColumn('Flag'),
    },
  ],
});

// =====================================================================
// Pipeline review by stage
// =====================================================================

export const OPEN_STAGES = ['Prospecting', 'Discovery', 'Proposal', 'Negotiation'] as const;
export const CLOSED_STAGES = ['Closed Won', 'Closed Lost'] as const;
/** How an open pipeline usually spreads: lots of early deals, few in negotiation. */
const STAGE_WEIGHTS = [0.3, 0.3, 0.25, 0.15];

/** Win probability per open stage, in OPEN_STAGES order. Any two sets differ in at least three stages. */
export const PROBABILITY_SETS: readonly (readonly number[])[] = [
  [0.1, 0.25, 0.5, 0.75],
  [0.05, 0.2, 0.4, 0.7],
  [0.15, 0.3, 0.6, 0.8],
  [0.1, 0.2, 0.45, 0.65],
];

/**
 * Monday export dates. Never TODAY(): the sheet has to give the same answer next week. Every one
 * of these has an open deal closing on it (not slipped yet) and one closing the day before (slipped).
 */
export const AS_OF_DATES = [serial(2026, 9, 14), serial(2026, 9, 21), serial(2026, 9, 28)];
const OPEN_FROM = serial(2026, 9, 1);
const OPEN_TO = serial(2026, 12, 18);
/** Closed deals all closed before the earliest as-of date, so "still open" decides whether they count. */
const CLOSED_FROM = serial(2026, 8, 3);
const CLOSED_TO = serial(2026, 9, 11);

/** Reps run down K9:K14, one grid row each; every answer below the stage summary shares those rows. */
const PIPELINE_REPS = 6;
const LAST_REP_ROW = 8 + PIPELINE_REPS;
const GRID = `L9:O${LAST_REP_ROW}`;
const COVERAGE = `Q9:Q${LAST_REP_ROW}`;
const SLIPPED = `R9:R${LAST_REP_ROW}`;

export interface Opp {
  id: string;
  rep: string;
  stage: string;
  date: number;
  amount: number;
}

export interface StageProbability {
  stage: string;
  probability: number;
}

export interface PipelineData {
  opps: Opp[];
  stages: StageProbability[];
  /** As-of date of the export (I1). */
  asOf: number;
  /** Stage order down K2:K5. */
  stageRows: string[];
  /** Stage order across L8:O8. */
  gridStages: string[];
  /** Rep order down K9:K14, with each rep's quarterly quota in column P. */
  reps: RepQuota[];
}

/** Amount is the last column, so the Weighted column the learner adds picks up its dollar format. */
const OPP_COLS: ColumnSpec[] = [
  { header: 'Opportunity' },
  { header: 'Rep' },
  { header: 'Stage' },
  { header: 'Close date', format: FMT.date },
  { header: 'Amount', format: DOLLARS },
];
const STAGE_COLS: ColumnSpec[] = [{ header: 'Stage' }, { header: 'Probability', format: '0%' }];

const oppGrid = (opps: Opp[]): Grid => opps.map((o) => [o.id, o.rep, o.stage, o.date, o.amount]);
const stageGrid = (stages: StageProbability[]): Grid => stages.map((s) => [s.stage, s.probability]);
const oppId = (n: number) => `OPP-${3101 + n}`;
const oppNumber = (id: string) => Number(id.slice(4)) - 3101;

export const isOpenStage = (stage: string) => (OPEN_STAGES as readonly string[]).includes(stage);

const probabilitySet = (set: readonly number[]): StageProbability[] => OPEN_STAGES.map((stage, i) => ({ stage, probability: set[i] }));

function pickStage(rng: Rng): string {
  let r = rng.next();
  for (let i = 0; i < OPEN_STAGES.length; i++) {
    r -= STAGE_WEIGHTS[i];
    if (r < 0) return OPEN_STAGES[i];
  }
  return OPEN_STAGES[OPEN_STAGES.length - 1];
}

const oppAmount = (rng: Rng) => rng.int(10, 160) * 500;
const openOpp = (rng: Rng, rep: string): Opp => ({ id: '', rep, stage: pickStage(rng), date: rng.int(OPEN_FROM, OPEN_TO), amount: oppAmount(rng) });
const closedOpp = (rng: Rng, rep: string): Opp => ({ id: '', rep, stage: rng.chance(0.55) ? 'Closed Won' : 'Closed Lost', date: rng.int(CLOSED_FROM, CLOSED_TO), amount: oppAmount(rng) });

/** Every open stage has at least one deal, so no row of the stage summary is trivially zero. */
function coverEveryStage(rng: Rng, opps: Opp[]): Opp[] {
  const out = opps.map((o) => ({ ...o }));
  for (const stage of OPEN_STAGES) {
    if (out.some((o) => o.stage === stage)) continue;
    const counts = OPEN_STAGES.map((s) => out.filter((o) => o.stage === s).length);
    const busiest = OPEN_STAGES[counts.indexOf(Math.max(...counts))];
    rng.pick(out.filter((o) => o.stage === busiest)).stage = stage;
  }
  return out;
}

export const probabilityOf = (d: PipelineData, stage: string) => d.stages.find((s) => s.stage === stage)?.probability ?? 0;

const amounts = (opps: Opp[]) => sum(opps.map((o) => o.amount));

export function stageSummary(d: PipelineData): Grid {
  return d.stageRows.map((stage) => {
    const mine = d.opps.filter((o) => o.stage === stage);
    return [amounts(mine), mine.length];
  });
}

/** Amount × the stage's probability; closed stages aren't in the Stages Table, so they weigh 0. */
export const weightedAmounts = (d: PipelineData) => d.opps.map((o) => o.amount * probabilityOf(d, o.stage));

export const repStageTotal = (d: PipelineData, rep: string, stage: string) => amounts(d.opps.filter((o) => o.rep === rep && o.stage === stage));

export const openTotal = (d: PipelineData, rep: string) => amounts(d.opps.filter((o) => o.rep === rep && isOpenStage(o.stage)));

/** Still open, and the close date is before the as-of date. Closing on the as-of date hasn't slipped yet. */
export const slippedCount = (d: PipelineData, rep: string) => d.opps.filter((o) => o.rep === rep && isOpenStage(o.stage) && o.date < d.asOf).length;

const pipelineAmountsChange: Variant<PipelineData> = {
  label: 'deal amounts change',
  apply: (d, rng) => {
    const opps = d.opps.map((o) => ({ ...o, amount: oppAmount(rng) }));
    if (opps.every((o, i) => o.amount === d.opps[i].amount)) opps[0].amount += 500;
    return { ...d, opps };
  },
};

const oppsResorted: Variant<PipelineData> = {
  label: 'the Opps Table is re-sorted',
  explain: 'Total whole Table columns with SUMIFS so the order of the rows doesn’t matter.',
  apply: (d, rng) => ({ ...d, opps: reshuffle(rng, d.opps, (o) => o.id) }),
};

const newOpps: Variant<PipelineData> = {
  label: 'five new opportunities are added',
  explain: 'Use whole Table columns like Opps[Amount] so new rows are counted.',
  apply: (d, rng) => {
    const next = Math.max(...d.opps.map((o) => oppNumber(o.id))) + 1;
    const added = Array.from({ length: 5 }, (_, i) => ({ ...openOpp(rng, rng.pick(d.reps).rep), id: oppId(next + i), date: rng.int(d.asOf, OPEN_TO) }));
    // One of them was entered late and has already slipped.
    added[0].date = d.asOf - rng.int(1, 10);
    return { ...d, opps: [...d.opps, ...added] };
  },
};

/** A week of selling: deals advance a stage, win or lose. One slipped deal always closes. */
const dealsMove: Variant<PipelineData> = {
  label: 'deals move to new stages',
  apply: (d, rng) => {
    const opps = d.opps.map((o) => ({ ...o }));
    const open = opps.filter((o) => isOpenStage(o.stage));
    const late = open.filter((o) => o.date < d.asOf);
    const closing = late.length ? rng.pick(late) : rng.pick(open);
    closing.stage = rng.chance(0.5) ? 'Closed Won' : 'Closed Lost';
    for (const o of rng.sample(open.filter((x) => x !== closing), 6)) {
      const r = rng.next();
      const i = OPEN_STAGES.indexOf(o.stage as (typeof OPEN_STAGES)[number]);
      o.stage = r < 0.5 ? (OPEN_STAGES[i + 1] ?? 'Closed Won') : r < 0.75 ? 'Closed Won' : 'Closed Lost';
    }
    return { ...d, opps };
  },
};

const probabilitiesChange: Variant<PipelineData> = {
  label: 'the stage probabilities change',
  explain: 'Look the probability up in the Stages Table instead of typing it.',
  apply: (d, rng) => {
    const current = d.stages.map((s) => s.probability).join('|');
    const set = rng.pick(PROBABILITY_SETS.filter((s) => s.join('|') !== current));
    return { ...d, stages: probabilitySet(set) };
  },
};

const asOfChanges: Variant<PipelineData> = {
  label: 'the as-of date in I1 changes',
  explain: 'Compare with $I$1 instead of typing a date or using TODAY().',
  apply: (d, rng) => ({ ...d, asOf: rng.pick(AS_OF_DATES.filter((x) => x !== d.asOf)) }),
};

/** Reps push most slipped deals out to new dates, and a couple of deals that looked safe slip. */
const closeDatesChange: Variant<PipelineData> = {
  label: 'close dates are updated',
  apply: (d, rng) => {
    const opps = d.opps.map((o) => ({ ...o }));
    const open = opps.filter((o) => isOpenStage(o.stage));
    const late = rng.shuffle(open.filter((o) => o.date < d.asOf));
    const ahead = rng.shuffle(open.filter((o) => o.date >= d.asOf));
    for (const o of late.slice(1)) o.date = d.asOf + rng.int(7, 60);
    for (const o of ahead.slice(0, 2)) o.date = d.asOf - rng.int(1, 14);
    return { ...d, opps };
  },
};

const pipelineQuotasChange: Variant<PipelineData> = {
  label: 'quotas change',
  explain: 'Divide by the quota in column P on the same row instead of typing it.',
  apply: (d, rng) => {
    const reps = d.reps.map((r) => ({ ...r, quota: rng.int(6, 14) * 25000 }));
    if (reps.every((r, i) => r.quota === d.reps[i].quota)) reps[0].quota += 25000;
    return { ...d, reps };
  },
};

const stageRowsReordered: Variant<PipelineData> = {
  label: 'the stages in column K are reordered',
  explain: 'Point the criteria at the stage in column K instead of typing it.',
  apply: (d, rng) => ({ ...d, stageRows: reshuffle(rng, d.stageRows) }),
};

const gridStagesReordered: Variant<PipelineData> = {
  label: 'the stages in row 8 are reordered',
  explain: 'Lock the row with L$8 so every cell reads its stage from row 8.',
  apply: (d, rng) => ({ ...d, gridStages: reshuffle(rng, d.gridStages) }),
};

const repsReordered: Variant<PipelineData> = {
  label: 'the reps in column K are reordered',
  explain: 'Lock the column with $K9 so every cell reads its rep from column K.',
  apply: (d, rng) => ({ ...d, reps: reshuffle(rng, d.reps, (r) => r.rep) }),
};

export const pipelineByStage = defineMission<PipelineData>({
  id: 'mission-pipeline-by-stage',
  title: 'Pipeline review by stage',
  role: 'sales',
  summary: 'The weekly pipeline review: open and weighted pipeline, a rep-by-stage grid, coverage and slipped deals',
  minutes: 22,
  skills: ['sumifs-warehouse', 'sumifs-grid', 'tables-calc-column', 'xlookup-not-found'],
  brief: (d) => ({
    from: 'Nate Brooks, Regional sales director',
    subject: 'Pipeline review for Monday’s forecast call',
    body: [
      'Hi,',
      '',
      `Monday’s forecast call is a pipeline review, and I want one sheet that answers the usual questions. The \`Opps\` Table is this morning’s CRM export (${d.opps.length} opportunities), \`Stages\` has the win probability we use for each open stage, and \`I1\` holds the export’s as-of date. Use \`I1\` rather than TODAY(), so the numbers still match the export when we look at it next week.`,
      '',
      '1. Open pipeline and the number of deals in each stage, in `L2:M5`.',
      '2. A `Weighted` column on `Opps`: each deal’s amount × its stage’s probability.',
      `3. Pipeline by rep and stage, in \`${GRID}\`.`,
      `4. Each rep’s coverage against quota, in \`${COVERAGE}\`.`,
      `5. Slipped deals per rep, in \`${SLIPPED}\`: still open, with a close date before \`I1\`.`,
      '',
      'Thanks,',
      'Nate',
    ].join('\n'),
  }),
  make: (rng) => {
    const reps = rng
      .sample(REPS, PIPELINE_REPS)
      .sort(excelTextCompare)
      .map((rep) => ({ rep, quota: rng.int(6, 14) * 25000 }));
    const opps = reps.flatMap((r) => [
      ...Array.from({ length: rng.int(6, 9) }, () => openOpp(rng, r.rep)),
      ...Array.from({ length: rng.int(1, 3) }, () => closedOpp(rng, r.rep)),
    ]);
    for (const day of AS_OF_DATES) for (const date of [day, day - 1]) opps.push({ ...openOpp(rng, rng.pick(reps).rep), date });
    // CRM exports list opportunities in the order they were created, which mixes reps and stages.
    const numbered = coverEveryStage(rng, rng.shuffle(opps)).map((o, i) => ({ ...o, id: oppId(i) }));
    return {
      opps: numbered,
      stages: probabilitySet(rng.pick(PROBABILITY_SETS)),
      asOf: rng.pick(AS_OF_DATES),
      stageRows: [...OPEN_STAGES],
      gridStages: [...OPEN_STAGES],
      reps,
    };
  },
  // Opps keeps column F free for the Weighted column and G as its gap. Stages sits below the
  // as-of date with column J and row 8 clear; the summaries start in column K.
  blocks: (d) => [
    dataBlock('Opps', 'A1', OPP_COLS, oppGrid(d.opps)),
    cells('H1', [['As of']], 'label'),
    cells('I1', [[d.asOf]], 'input', FMT.date),
    dataBlock('Stages', 'H3', STAGE_COLS, stageGrid(d.stages)),
    cells('K1', [['Stage', 'Open $', 'Count']], 'header'),
    cells('K2', column(d.stageRows), 'label'),
    cells('K8', [['Rep']], 'header'),
    cells('L8', [d.gridStages], 'header'),
    cells('P8', [['Quota', 'Coverage', 'Slipped']], 'header'),
    cells('K9', column(d.reps.map((r) => r.rep)), 'label'),
    cells('P9', column(d.reps.map((r) => r.quota)), 'input', DOLLARS),
  ],
  inputs: (d) => [
    tableWrite('Opps', OPP_COLS, oppGrid(d.opps)),
    tableWrite('Stages', STAGE_COLS, stageGrid(d.stages)),
    rangeWrite('I1', [[d.asOf]]),
    rangeWrite('K2:K5', column(d.stageRows)),
    rangeWrite('L8:O8', [d.gridStages]),
    rangeWrite(`K9:K${LAST_REP_ROW}`, column(d.reps.map((r) => r.rep))),
    rangeWrite(`P9:P${LAST_REP_ROW}`, column(d.reps.map((r) => r.quota))),
  ],
  steps: [
    // ---------- 1. SUMIFS and COUNTIFS by stage ----------
    {
      title: 'Open pipeline by stage',
      task: () =>
        'In `L2:M5`, summarize the open pipeline for each stage in column `K`: the total Amount in column `L` and the number of opportunities in column `M`, from the `Opps` Table. Write one formula in `L2` and one in `M2`, then fill each down to row `5`.',
      hints: [
        'SUMIFS adds up the amounts that match a condition, and COUNTIFS counts the rows. Both test Opps[Stage] against the stage in column `K`.',
        '`=SUMIFS(Opps[Amount], Opps[Stage], $K2)` in `L2` and `=COUNTIFS(Opps[Stage], $K2)` in `M2`, each filled down to row `5`.',
      ],
      solution: () => 'L2: =SUMIFS(Opps[Amount],Opps[Stage],$K2)    M2: =COUNTIFS(Opps[Stage],$K2)',
      // One number format covers both columns: whole numbers with separators.
      answer: () => ({ kind: 'cells', range: 'L2:M5', format: FMT.int, consistency: 'columns' }),
      expected: stageSummary,
      variants: [pipelineAmountsChange, newOpps, dealsMove, stageRowsReordered],
      rules: { forbidText: { values: [...OPEN_STAGES], advice: 'Point to the stage in column K instead.' }, allowNumbers: [] },
    },

    // ---------- 2. Calculated column with an if-not-found lookup ----------
    {
      title: 'Weighted pipeline',
      task: () =>
        'Add a column named `Weighted` to the `Opps` Table by typing the name in `F1`. Fill it with each opportunity’s Amount × its stage’s probability from the `Stages` Table. Closed Won and Closed Lost aren’t in `Stages`, so closed deals weigh 0.',
      hints: [
        'Type Weighted in `F1` and press {enter}. In `F2`, look up [@Stage] in Stages[Stage] and return Stages[Probability].',
        'Closed stages aren’t in the Stages Table, so a plain lookup returns #N/A for them. XLOOKUP’s fourth argument, if_not_found, can return 0 instead.',
        '`=[@Amount]*XLOOKUP([@Stage], Stages[Stage], Stages[Probability], 0)`',
      ],
      solution: () => '=[@Amount]*XLOOKUP([@Stage],Stages[Stage],Stages[Probability],0)',
      answer: () => ({ kind: 'tableColumn', table: 'Opps', column: 'Weighted', format: DOLLARS }),
      expected: (d) => weightedAmounts(d).map((w) => [w]),
      variants: [
        probabilitiesChange,
        pipelineAmountsChange,
        dealsMove,
        { ...newOpps, explain: 'A calculated column fills new rows by itself. Type the formula once in F2 and let the Table fill the column.' },
      ],
      rules: {
        require: [
          {
            pattern: /XLOOKUP\(|VLOOKUP\(|LOOKUP\(|INDEX\(/i,
            label: 'Looks up the stage probability',
            advice: 'XLOOKUP([@Stage], Stages[Stage], Stages[Probability], 0) returns the probability, or 0 for a closed stage.',
          },
        ],
        forbidText: { values: [...OPEN_STAGES], advice: 'Look the stage up in the Stages Table instead of typing stage names.' },
        allowNumbers: [0, 1, 2],
      },
      inspections: () => [{ kind: 'tableColumn', table: 'Opps', column: 'Weighted', label: 'The Opps Table has a Weighted column' }],
    },

    // ---------- 3. Two-way SUMIFS grid ----------
    {
      title: 'Pipeline by rep and stage',
      task: () =>
        `Fill \`${GRID}\` with each rep’s pipeline (column \`K\`) in each stage (row \`8\`). Write one formula in \`L9\` that still works when you fill it right and down.`,
      hints: [
        'Use SUMIFS with two conditions: Opps[Rep] matched against the rep in column `K`, and Opps[Stage] matched against the stage in row `8`.',
        'Lock the column on the rep ($K9) and the row on the stage (L$8) so both stay put as you fill. Press {absKey} to cycle the $ signs.',
        'Table column names shift when you fill right, the same way relative references do: in `M9`, Opps[Amount] would become Opps[Weighted]. Double the brackets, as in Opps[[Amount]:[Amount]], to lock a Table column.',
        `\`=SUMIFS(Opps[[Amount]:[Amount]], Opps[[Rep]:[Rep]], $K9, Opps[[Stage]:[Stage]], L$8)\`, then fill right to column \`O\` and down to row \`${LAST_REP_ROW}\`.`,
      ],
      solution: () => '=SUMIFS(Opps[[Amount]:[Amount]],Opps[[Rep]:[Rep]],$K9,Opps[[Stage]:[Stage]],L$8)',
      answer: () => ({ kind: 'cells', range: GRID, format: DOLLARS, consistency: 'all' }),
      expected: (d) => d.reps.map((r) => d.gridStages.map((s) => repStageTotal(d, r.rep, s))),
      variants: [pipelineAmountsChange, oppsResorted, gridStagesReordered, repsReordered, newOpps],
      rules: {
        require: [
          {
            // Dragging right shifts plain Table columns: Opps[Amount] becomes Opps[Weighted] in M9,
            // and Opps[Stage] becomes Opps[Close date]. Locked ([[Amount]:[Amount]]) and qualified forms pass.
            pattern: /^(?![\s\S]*Opps\[(?!(?:Amount|Rep|Stage)\]|[[#@]))/i,
            label: 'Table columns stay fixed across the grid',
            advice: 'Filling right shifts Table column names: one column over, Opps[Amount] becomes Opps[Weighted]. Lock each one as Opps[[Amount]:[Amount]], or copy L9 and paste it over the grid instead of dragging.',
          },
        ],
        forbidText: { values: [...OPEN_STAGES, ...ALL_REPS], advice: 'Point to the rep in column K and the stage in row 8 instead.' },
        allowNumbers: [],
      },
    },

    // ---------- 4. Ratio against quota ----------
    {
      title: 'Coverage against quota',
      task: () =>
        `In \`${COVERAGE}\`, show each rep’s pipeline coverage: their open pipeline (the four stage totals on the same row) divided by their quota in column \`P\`. A rep with $600,000 open against a $200,000 quota has coverage of 3, which the sheet shows as 3.00x. Write one formula in \`Q9\` and fill it down.`,
      hints: [
        'Add up the rep’s row of the grid with SUM($L9:$O9), then divide by $P9, the quota on the same row. The $ keeps the columns fixed, and the row moves down with each rep as you fill.',
        `\`=SUM($L9:$O9)/$P9\`, filled down to row \`${LAST_REP_ROW}\`.`,
      ],
      solution: () => '=SUM($L9:$O9)/$P9',
      answer: () => ({ kind: 'cells', range: COVERAGE, format: TIMES, consistency: 'all' }),
      expected: (d) => d.reps.map((r) => [openTotal(d, r.rep) / r.quota]),
      variants: [pipelineQuotasChange, pipelineAmountsChange, { ...repsReordered, explain: 'Point at the same row, $L9:$O9 and $P9, so the ratio moves with the rep.' }],
      rules: { forbidText: REP_TEXT, allowNumbers: [] },
    },

    // ---------- 5. COUNTIFS with a date cutoff and "not equal" ----------
    {
      title: 'Slipped deals',
      task: () =>
        `In \`${SLIPPED}\`, count each rep’s slipped deals: opportunities that are still open (any stage other than Closed Won or Closed Lost) with a Close date before the as-of date in \`I1\`. A deal closing on the as-of date hasn’t slipped yet. Write one formula in \`R9\` and fill it down.`,
      hints: [
        'COUNTIFS takes as many conditions as you need: the rep, a Close date before $I$1, and a stage that is neither closed stage.',
        'Build the date test by joining an operator to the cell: "<"&$I$1. For the stage, "<>Closed Won" means not Closed Won.',
        '`=COUNTIFS(Opps[Rep], $K9, Opps[Close date], "<"&$I$1, Opps[Stage], "<>Closed Won", Opps[Stage], "<>Closed Lost")`',
      ],
      solution: () => '=COUNTIFS(Opps[Rep],$K9,Opps[Close date],"<"&$I$1,Opps[Stage],"<>Closed Won",Opps[Stage],"<>Closed Lost")',
      answer: () => ({ kind: 'cells', range: SLIPPED, format: FMT.int, consistency: 'all' }),
      expected: (d) => d.reps.map((r) => [slippedCount(d, r.rep)]),
      variants: [asOfChanges, dealsMove, closeDatesChange, newOpps],
      // 0 allows ISNUMBER(MATCH(Opps[Stage], Stages[Stage], 0)) as the "still open" test.
      rules: { forbidText: REP_TEXT, allowNumbers: [0] },
    },
  ],
});

/** Sales missions in the order they're offered. */
export const SALES_MISSIONS: Mission<any>[] = [commissionStatement, pipelineByStage];

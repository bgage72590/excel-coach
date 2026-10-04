import { DEPARTMENTS, GL_ACCOUNTS, VENDORS, eomonth, fromSerial, monthStart, serial, sum } from '../engine/data';
import type { Rng } from '../engine/rng';
import type { AnswerArea, ColumnSpec, ExpectedGrid, Grid, Inspection, Variant } from '../engine/types';
import { FMT, cells, column, dataBlock, rangeWrite, tableWrite } from '../exercises/common';
import { defineMission, type Mission } from './types';

const MONTH_NAMES = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];

function monthName(s: number): string {
  const { year, month } = fromSerial(s);
  return `${MONTH_NAMES[month - 1]} ${year}`;
}

// =====================================================================
// Budget vs actual for the month
// =====================================================================

interface AccountSpec {
  /** Monthly budget range in dollars. */
  lo: number;
  hi: number;
  /** Steady accounts (payroll, rent, contracts) only move a few percent; volatile ones swing with volume. */
  steady: boolean;
  depts: readonly string[];
}

const ACCOUNT_SPECS: Record<string, AccountSpec> = {
  'Cost of goods sold': { lo: 140_000, hi: 260_000, steady: false, depts: ['Operations', 'Warehouse'] },
  'Salaries and wages': { lo: 150_000, hi: 240_000, steady: true, depts: DEPARTMENTS },
  'Freight in': { lo: 22_000, hi: 45_000, steady: false, depts: ['Procurement', 'Warehouse'] },
  Rent: { lo: 28_000, hi: 55_000, steady: true, depts: ['Operations', 'Warehouse'] },
  'Freight out': { lo: 15_000, hi: 35_000, steady: false, depts: ['Warehouse', 'Sales'] },
  'Payroll taxes': { lo: 12_000, hi: 22_000, steady: true, depts: ['Finance', 'HR'] },
  'Software subscriptions': { lo: 6_000, hi: 18_000, steady: true, depts: ['IT', 'Finance', 'Sales', 'Customer success'] },
  Insurance: { lo: 5_000, hi: 12_000, steady: true, depts: ['Finance', 'Operations'] },
  'Repairs and maintenance': { lo: 5_000, hi: 14_000, steady: false, depts: ['Warehouse', 'Operations'] },
  Utilities: { lo: 4_000, hi: 9_000, steady: false, depts: ['Operations', 'Warehouse'] },
  Travel: { lo: 3_000, hi: 10_000, steady: false, depts: ['Sales', 'Customer success', 'Procurement'] },
  'Office supplies': { lo: 1_500, hi: 4_000, steady: false, depts: DEPARTMENTS },
  'Inventory adjustments': { lo: 2_000, hi: 6_000, steady: false, depts: ['Warehouse'] },
};

/**
 * Three accounts always (two big volatile ones and payroll), two steady ones and three small
 * volatile ones, so every month can show each kind of flag outcome.
 */
const ALWAYS_BUDGETED = ['Cost of goods sold', 'Salaries and wages', 'Freight in'];
const ACCOUNT_POOLS: readonly { names: readonly string[]; take: number }[] = [
  { names: ['Rent', 'Freight out', 'Payroll taxes', 'Software subscriptions', 'Insurance'], take: 2 },
  { names: ['Repairs and maintenance', 'Utilities', 'Travel', 'Office supplies', 'Inventory adjustments'], take: 3 },
];

export const BVA_MONTHS = [serial(2026, 7, 1), serial(2026, 8, 1), serial(2026, 9, 1)];

/**
 * Review limits the data is built around. The first is company policy (the sheet's starting
 * values); the "limits change" variant picks another. Every account-month variance is generated
 * well clear of every one of these, so float noise in Excel's sums can never flip a flag.
 */
export const REVIEW_LIMITS: readonly { pct: number; amount: number }[] = [
  { pct: 0.1, amount: 5000 },
  { pct: 0.05, amount: 2500 },
  { pct: 0.15, amount: 10000 },
  { pct: 0.08, amount: 7500 },
];
export const PCT_MARGIN = 0.006;
export const DOLLAR_MARGIN = 150;

export interface Account {
  name: string;
  budget: number;
}

export interface GlLine {
  date: number;
  account: string;
  dept: string;
  amount: number;
}

export interface BvaData {
  accounts: Account[];
  lines: GlLine[];
  /** First day of the month being reported (G1). */
  month: number;
  /** Review limit as a fraction (G2). */
  pct: number;
  /** Review limit in dollars (G3). */
  limit: number;
}

const GL_COLS: ColumnSpec[] = [
  { header: 'Date', format: FMT.date },
  { header: 'Account' },
  { header: 'Department' },
  { header: 'Amount', format: FMT.currency },
];
const BUDGET_COLS: ColumnSpec[] = [{ header: 'Account' }, { header: 'Budget', format: '$#,##0' }];

const glGrid = (lines: GlLine[]): Grid => lines.map((l) => [l.date, l.account, l.dept, l.amount]);
const budgetGrid = (accounts: Account[]): Grid => accounts.map((a) => [a.name, a.budget]);
const glNumber = (name: string) => GL_ACCOUNTS.find((a) => a.name === name)?.no ?? 0;

function pickAccounts(rng: Rng): Account[] {
  const names = [...ALWAYS_BUDGETED, ...ACCOUNT_POOLS.flatMap((pool) => rng.sample(pool.names, pool.take))].sort((a, b) => glNumber(a) - glNumber(b));
  return names.map((name) => {
    const spec = ACCOUNT_SPECS[name];
    return { name, budget: rng.int(spec.lo / 100, spec.hi / 100) * 100 };
  });
}

/** True when a variance (in cents) sits clear of every review limit, in percent and in dollars. */
export function clearOfLimits(budget: number, varianceCents: number): boolean {
  const dollars = Math.abs(varianceCents) / 100;
  const pct = dollars / budget;
  return pct >= 0.003 && REVIEW_LIMITS.every((l) => Math.abs(pct - l.pct) >= PCT_MARGIN && Math.abs(dollars - l.amount) >= DOLLAR_MARGIN);
}

export type FlagOutcome = 'over' | 'under' | 'pctOnly' | 'dollarOnly' | 'within';

export function flagOutcome(budget: number, variance: number, limit: { pct: number; amount: number }): FlagOutcome {
  const pctOver = Math.abs(variance / budget) > limit.pct;
  const dollarOver = Math.abs(variance) > limit.amount;
  if (pctOver && dollarOver) return variance > 0 ? 'over' : 'under';
  if (pctOver) return 'pctOnly';
  if (dollarOver) return 'dollarOnly';
  return 'within';
}

const SMALL_SWING: [number, number] = [0.004, 0.094];
const BIG_SWING: [number, number] = [0.106, 0.32];

/** The range of variance sizes (as a fraction of budget) that gives `role` under company policy. */
function swingFor(role: FlagOutcome | 'any', budget: number, steady: boolean): [number, number] {
  const policy = REVIEW_LIMITS[0];
  const overDollars = (policy.amount + DOLLAR_MARGIN) / budget;
  const underDollars = (policy.amount - DOLLAR_MARGIN) / budget;
  switch (role) {
    case 'over':
    case 'under':
      return [Math.max(BIG_SWING[0], overDollars + 0.001), BIG_SWING[1]];
    case 'pctOnly':
      return [BIG_SWING[0], Math.min(BIG_SWING[1], underDollars - 0.0005)];
    case 'dollarOnly':
      return [Math.max(SMALL_SWING[0], overDollars + 0.0005), SMALL_SWING[1]];
    case 'within':
      return [SMALL_SWING[0], Math.min(SMALL_SWING[1], underDollars - 0.0005)];
    default:
      return steady ? SMALL_SWING : [SMALL_SWING[0], BIG_SWING[1]];
  }
}

/**
 * One account's variance for one month, in cents: clear of every review limit, and landing in
 * `role` under company policy when a role is given.
 */
function drawVariance(rng: Rng, budget: number, steady: boolean, role: FlagOutcome | 'any'): number {
  const sign = role === 'over' ? 1 : role === 'under' ? -1 : rng.chance(0.6) ? 1 : -1;
  const fits = (cents: number) => clearOfLimits(budget, cents) && (role === 'any' || flagOutcome(budget, cents / 100, REVIEW_LIMITS[0]) === role);
  const [lo, hi] = swingFor(role, budget, steady);
  for (let i = 0; i < 80; i++) {
    const size =
      role === 'any' && !steady
        ? rng.chance(0.45)
          ? rng.float(SMALL_SWING[0], SMALL_SWING[1], 4)
          : rng.float(BIG_SWING[0], BIG_SWING[1], 4)
        : rng.float(lo, hi, 4);
    const cents = sign * Math.round(budget * 100 * size);
    if (fits(cents)) return cents;
  }
  // Deterministic fallback: walk the range in 0.05% steps.
  for (let size = lo; size <= hi + 1e-9; size += 0.0005) {
    const cents = sign * Math.round(budget * 100 * size);
    if (fits(cents)) return cents;
  }
  return sign * Math.round(budget * 100 * lo);
}

/**
 * Variances for one month. Payroll lands over the dollar limit but under the percent limit; two
 * volatile accounts big enough to matter land outside policy, one over and one under; one small
 * volatile account swings by a big percent that's still under the dollar limit. The rest are free.
 */
function drawMonth(rng: Rng, accounts: Account[]): number[] {
  const roles = new Map<number, FlagOutcome>();
  const steady = (i: number) => ACCOUNT_SPECS[accounts[i].name].steady;
  const idx = accounts.map((_, i) => i);
  const payroll = accounts.findIndex((a) => a.name === 'Salaries and wages');
  if (payroll >= 0) roles.set(payroll, 'dollarOnly');
  const reviewable = rng.shuffle(idx.filter((i) => !roles.has(i) && !steady(i) && accounts[i].budget >= 20_000));
  if (reviewable[0] !== undefined) roles.set(reviewable[0], 'over');
  if (reviewable[1] !== undefined) roles.set(reviewable[1], 'under');
  const small = rng.shuffle(idx.filter((i) => !roles.has(i) && !steady(i) && accounts[i].budget <= 30_000));
  if (small[0] !== undefined) roles.set(small[0], 'pctOnly');
  return accounts.map((a, i) => drawVariance(rng, a.budget, steady(i), roles.get(i) ?? 'any'));
}

/** Splits a total (in cents) into `parts` positive amounts that add up to it exactly. */
function splitCents(rng: Rng, total: number, parts: number): number[] {
  const weights = Array.from({ length: parts }, () => rng.float(0.4, 1, 3));
  const w = sum(weights);
  const out = weights.slice(0, -1).map((x) => Math.round((total * x) / w));
  out.push(total - sum(out));
  return out;
}

/**
 * Dates, accounts and departments for every GL line. Each month has a line posted on its first
 * day and one on its last day, so a window built with > or < instead of >= and <= gives a wrong total.
 */
function glSkeleton(rng: Rng, accounts: Account[]): GlLine[] {
  const lines: GlLine[] = [];
  for (const m of BVA_MONTHS) {
    const month: GlLine[] = [];
    for (const a of accounts) {
      const n = rng.int(2, 4);
      for (let i = 0; i < n; i++) month.push({ date: rng.int(m, eomonth(m)), account: a.name, dept: rng.pick(ACCOUNT_SPECS[a.name].depts), amount: 0 });
    }
    const [first, last] = rng.sample(month, 2);
    first.date = m;
    last.date = eomonth(m);
    lines.push(...month);
  }
  return lines.sort((x, y) => x.date - y.date);
}

/** Sets every line's amount so each account-month total hits a target variance. Dates and departments stay put. */
function priceLines(rng: Rng, accounts: Account[], lines: GlLine[]): GlLine[] {
  const out = lines.map((l) => ({ ...l }));
  for (const m of BVA_MONTHS) {
    const variances = drawMonth(rng, accounts);
    accounts.forEach((a, i) => {
      const idx = out.flatMap((l, k) => (l.account === a.name && monthStart(l.date) === m ? [k] : []));
      const parts = splitCents(rng, a.budget * 100 + variances[i], idx.length);
      idx.forEach((k, j) => {
        out[k].amount = parts[j] / 100;
      });
    });
  }
  return out;
}

/** What SUMIFS returns for each budget account in the chosen month. */
export function actuals(d: BvaData): number[] {
  const end = eomonth(d.month);
  const lines = d.lines.filter((l) => l.date >= d.month && l.date <= end);
  return d.accounts.map((a) => sum(lines.filter((l) => l.account === a.name).map((l) => l.amount)));
}

export function reviewFlags(d: BvaData): string[] {
  const act = actuals(d);
  return d.accounts.map((a, i) => {
    const variance = act[i] - a.budget;
    return Math.abs(variance / a.budget) > d.pct && Math.abs(variance) > d.limit ? 'Review' : '';
  });
}

const budgetColumn = (name: string, format?: string): AnswerArea => ({ kind: 'tableColumn', table: 'Budget', column: name, format });
const hasColumn = (name: string): Inspection[] => [{ kind: 'tableColumn', table: 'Budget', column: name, label: `The Budget Table has a ${name} column` }];

const monthChanges: Variant<BvaData> = {
  label: 'the month in G1 changes',
  explain: 'Build the date window from $G$1 and EOMONTH instead of typing dates.',
  apply: (d, rng) => ({ ...d, month: rng.pick(BVA_MONTHS.filter((m) => m !== d.month)) }),
};

/** The same month change, explained for the steps that build on the Actual column. */
const monthChangesLater: Variant<BvaData> = {
  ...monthChanges,
  explain: 'Build it from the columns on the same row, so it follows the actuals when the month changes.',
};

const glAmountsChange: Variant<BvaData> = {
  label: 'the GL amounts change',
  apply: (d, rng) => ({ ...d, lines: priceLines(rng, d.accounts, d.lines) }),
};

const lateEntries: Variant<BvaData> = {
  label: 'late journal entries are posted to the GL',
  explain: 'Use whole Table columns like GL[Amount] so new rows are counted.',
  apply: (d, rng) => ({
    ...d,
    lines: [
      ...d.lines,
      ...Array.from({ length: 4 }, () => {
        const a = rng.pick(d.accounts);
        return { date: rng.int(d.month, eomonth(d.month)), account: a.name, dept: rng.pick(ACCOUNT_SPECS[a.name].depts), amount: rng.float(150, 3500, 2) };
      }),
    ],
  }),
};

const budgetsChange: Variant<BvaData> = {
  label: 'the budgets change',
  explain: 'Point at the Budget column on the same row instead of typing a number.',
  apply: (d, rng) => ({
    ...d,
    accounts: d.accounts.map((a) => {
      const factor = rng.chance(0.5) ? rng.float(0.8, 0.95, 3) : rng.float(1.05, 1.25, 3);
      return { ...a, budget: Math.round((a.budget * factor) / 100) * 100 };
    }),
  }),
};

const limitsChange: Variant<BvaData> = {
  label: 'the review limits in G2 and G3 change',
  explain: 'Compare against $G$2 and $G$3 instead of typing the limits into the formula.',
  apply: (d, rng) => {
    // Prefer limits that change at least one flag, so typed-in limits can't pass by luck.
    const current = reviewFlags(d).join('|');
    const others = rng.shuffle(REVIEW_LIMITS.filter((l) => l.pct !== d.pct || l.amount !== d.limit));
    const pick = others.find((l) => reviewFlags({ ...d, pct: l.pct, limit: l.amount }).join('|') !== current) ?? others[0];
    return { ...d, pct: pick.pct, limit: pick.amount };
  },
};

const ACCOUNT_NAMES = Object.keys(ACCOUNT_SPECS);

export const budgetVsActual = defineMission<BvaData>({
  id: 'mission-budget-vs-actual',
  title: 'Budget vs actual for the month',
  role: 'finance',
  summary: 'The month-end budget vs actual report and its review flags',
  minutes: 20,
  skills: ['tables-calc-column', 'sumifs-month', 'lambda-varpct'],
  brief: (d) => ({
    from: 'Priya Raman, Controller',
    subject: `Budget vs actual for ${monthName(d.month)}`,
    body: [
      'Hi,',
      `Before Thursday’s review I need budget vs actual for ${monthName(d.month)}. The \`GL\` Table is the detail export for July through September, and the \`Budget\` Table has each account’s monthly budget.`,
      'Add the actuals for the month in `G1`, the variance in dollars and as a percent, and a flag on anything outside policy: more than 10% and more than $5,000 off budget, over or under. The limits sit in `G2` and `G3` in case audit tightens them.',
      'Build it off `G1` so I can switch months without calling you.',
      'Priya',
    ].join('\n\n'),
  }),
  make: (rng) => {
    const accounts = pickAccounts(rng);
    const lines = priceLines(rng, accounts, glSkeleton(rng, accounts));
    return { accounts, lines, month: rng.pick(BVA_MONTHS), pct: REVIEW_LIMITS[0].pct, limit: REVIEW_LIMITS[0].amount };
  },
  // Budget sits at F6 with room on its right for the four columns the learner adds (H:K), and
  // nothing in column L or row 15, so the grown Table never runs into anything.
  blocks: (d) => [
    dataBlock('GL', 'A1', GL_COLS, glGrid(d.lines)),
    cells('F1', [['Month'], ['Review limit %'], ['Review limit $']], 'label'),
    cells('G1', [[d.month]], 'input', FMT.month),
    cells('G2', [[d.pct]], 'input', '0%'),
    cells('G3', [[d.limit]], 'input', '$#,##0'),
    dataBlock('Budget', 'F6', BUDGET_COLS, budgetGrid(d.accounts)),
  ],
  inputs: (d) => [
    tableWrite('GL', GL_COLS, glGrid(d.lines)),
    tableWrite('Budget', BUDGET_COLS, budgetGrid(d.accounts)),
    rangeWrite('G1:G3', [[d.month], [d.pct], [d.limit]]),
  ],
  steps: [
    {
      title: 'Actuals for the month',
      task: () =>
        'Add a column named `Actual` to the `Budget` Table by typing the name in `H6`. In it, total each account’s `GL` amounts dated within the month in `G1` (`G1` holds the first day of that month). Every total should update when someone changes `G1`.',
      hints: [
        'Type Actual in H6 and press {enter}; the Table grows to take in the new column, and the formula you write in H7 fills the whole column.',
        'Use SUMIFS with three conditions: GL[Account] matches [@Account], GL[Date] is on or after $G$1, and GL[Date] is on or before EOMONTH($G$1,0). Press {absKey} to lock G1.',
        '`=SUMIFS(GL[Amount], GL[Account], [@Account], GL[Date], ">="&$G$1, GL[Date], "<="&EOMONTH($G$1,0))`',
      ],
      solution: () => '=SUMIFS(GL[Amount],GL[Account],[@Account],GL[Date],">="&$G$1,GL[Date],"<="&EOMONTH($G$1,0))',
      answer: () => budgetColumn('Actual', FMT.currency),
      expected: (d) => actuals(d).map((a) => [a]),
      variants: [monthChanges, glAmountsChange, lateEntries],
      rules: {
        forbidText: { values: ACCOUNT_NAMES, advice: 'Match against [@Account] instead of typing an account name.' },
        allowNumbers: [0, 1],
      },
      inspections: () => hasColumn('Actual'),
    },
    {
      title: 'Variance in dollars',
      task: () =>
        'Add a `Variance $` column in `I6`: each account’s actual minus its budget. Spending over budget shows as a positive number, under budget as a negative one.',
      hints: [
        'Type Variance $ in I6. In I7, type = and click the cells on the same row, so Excel writes row references like [@Actual].',
        '`=[@Actual]-[@Budget]`',
      ],
      solution: () => '=[@Actual]-[@Budget]',
      answer: () => budgetColumn('Variance $', FMT.currency),
      expected: (d) => {
        const act = actuals(d);
        return d.accounts.map((a, i) => [act[i] - a.budget]);
      },
      variants: [glAmountsChange, budgetsChange],
      rules: { allowNumbers: [] },
      inspections: () => hasColumn('Variance $'),
    },
    {
      title: 'Variance as a percent',
      task: () =>
        'Add a `Variance %` column in `J6` that divides each account’s dollar variance by its budget. The value is what’s checked, so 12.5% over budget is 0.125; format the column as a percentage so it reads that way.',
      hints: [
        'Divide this row’s Variance $ by this row’s Budget. Column names with spaces or symbols need an extra pair of brackets: [@[Variance $]].',
        '`=[@[Variance $]]/[@Budget]`, then Home › Percent Style (the % button) to show it as a percentage.',
      ],
      solution: () => '=[@[Variance $]]/[@Budget]',
      answer: () => budgetColumn('Variance %', FMT.pct),
      expected: (d) => {
        const act = actuals(d);
        return d.accounts.map((a, i) => [(act[i] - a.budget) / a.budget]);
      },
      variants: [glAmountsChange, budgetsChange, monthChangesLater],
      rules: { allowNumbers: [1] },
      inspections: () => hasColumn('Variance %'),
    },
    {
      title: 'Flag accounts for review',
      task: () =>
        'Add a `Flag` column in `K6` that shows `Review` when an account is outside policy: its variance % is more than the limit in `G2` and its variance $ is more than the limit in `G3`, comparing absolute values so over and under budget both count. Otherwise show an empty string (""). Point to `G2` and `G3` instead of typing the limits.',
      hints: [
        'Both tests must be true, so put them inside AND. ABS turns an under-budget variance positive, so one test covers both directions.',
        'The two tests are ABS([@[Variance %]])>$G$2 and ABS([@[Variance $]])>$G$3. Press {absKey} to lock each cell.',
        '`=IF(AND(ABS([@[Variance %]])>$G$2, ABS([@[Variance $]])>$G$3), "Review", "")`',
      ],
      solution: () => '=IF(AND(ABS([@[Variance %]])>$G$2,ABS([@[Variance $]])>$G$3),"Review","")',
      answer: () => budgetColumn('Flag'),
      expected: (d) => reviewFlags(d).map((f) => [f]),
      variants: [monthChangesLater, glAmountsChange, limitsChange],
      rules: { allowNumbers: [0] },
      inspections: () => hasColumn('Flag'),
    },
  ],
});

// =====================================================================
// Vendor on-time scorecard
// =====================================================================

export interface Receipt {
  vendor: string;
  promised: number;
  delivered: number;
  qty: number;
  defects: number;
}

interface VendorProfile {
  vendor: string;
  /** Chance a delivery arrives on or before the promised date. */
  onTime: number;
  /** Typical share of defective units. */
  defect: number;
}

export interface ScoreData {
  profiles: VendorProfile[];
  rows: Receipt[];
  /** Vendor order in G2:G7. */
  vendors: string[];
}

const RECEIPT_COLS: ColumnSpec[] = [
  { header: 'Vendor' },
  { header: 'Promised', format: FMT.date },
  { header: 'Delivered', format: FMT.date },
  { header: 'Qty', format: FMT.int },
  { header: 'Defects', format: FMT.int },
];

const receiptGrid = (rows: Receipt[]): Grid => rows.map((r) => [r.vendor, r.promised, r.delivered, r.qty, r.defects]);

const FIRST_PROMISE = serial(2026, 8, 3);
const LAST_PROMISE = serial(2026, 9, 30);

export function vendorTally(rows: Receipt[], vendor: string) {
  const mine = rows.filter((r) => r.vendor === vendor);
  return {
    count: mine.length,
    onTime: mine.filter((r) => r.delivered <= r.promised).length,
    qty: sum(mine.map((r) => r.qty)),
    defects: sum(mine.map((r) => r.defects)),
  };
}

export function onTimeRate(rows: Receipt[], vendor: string): number {
  const t = vendorTally(rows, vendor);
  return t.onTime / t.count;
}

/** No two vendors share an on-time rate (compared as exact fractions), so a sort has one right order. */
export function ratesAreDistinct(rows: Receipt[]): boolean {
  const t = [...new Set(rows.map((r) => r.vendor))].map((v) => vendorTally(rows, v));
  for (let i = 0; i < t.length; i++) {
    for (let j = i + 1; j < t.length; j++) if (t[i].onTime * t[j].count === t[j].onTime * t[i].count) return false;
  }
  return true;
}

const defectsFor = (rng: Rng, qty: number, rate: number) => rng.int(0, Math.max(1, Math.round(qty * rate * 2)));

function receipt(rng: Rng, p: VendorProfile, from = FIRST_PROMISE, to = LAST_PROMISE): Receipt {
  const qty = rng.int(4, 60) * 10;
  return { vendor: p.vendor, promised: rng.int(from, to), delivered: 0, qty, defects: defectsFor(rng, qty, p.defect) };
}

/** Days from promised to delivered: 0 or early when on time, 1–7 days late otherwise. */
const deliveryOffset = (rng: Rng, onTime: boolean) => (onTime ? (rng.chance(0.4) ? 0 : -rng.int(1, 3)) : rng.int(1, 7));

/** Every vendor with on-time deliveries gets at least one on the promised day, so < and <= give different answers. */
function withOnTheDay(rows: Receipt[]): Receipt[] {
  const out = rows.map((r) => ({ ...r }));
  for (const v of new Set(out.map((r) => r.vendor))) {
    const onTime = out.filter((r) => r.vendor === v && r.delivered <= r.promised);
    if (onTime.length && !onTime.some((r) => r.delivered === r.promised)) onTime[0].delivered = onTime[0].promised;
  }
  return out;
}

/** Sets every Delivered date, redrawing until no two vendors tie on on-time rate. */
function setDelivered(rng: Rng, rows: Receipt[], profiles: VendorProfile[]): Receipt[] {
  const chance = new Map(profiles.map((p) => [p.vendor, p.onTime]));
  let out = rows;
  for (let i = 0; i < 500; i++) {
    out = rows.map((r) => {
      const onTime = rng.chance(chance.get(r.vendor) ?? 0.8);
      return { ...r, delivered: r.promised + deliveryOffset(rng, onTime) };
    });
    if (ratesAreDistinct(out)) break;
  }
  return withOnTheDay(out);
}

const scoreDatesChange: Variant<ScoreData> = {
  label: 'delivery dates change',
  explain: 'Compare Deliveries[Delivered] with Deliveries[Promised] row by row instead of typing counts or rates.',
  apply: (d, rng) => ({ ...d, rows: setDelivered(rng, d.rows, d.profiles) }),
};

const moreDeliveries: Variant<ScoreData> = {
  label: 'new deliveries are added',
  explain: 'Use whole Table columns like Deliveries[Vendor] so new rows are counted.',
  apply: (d, rng) => {
    let rows = d.rows;
    for (let i = 0; i < 500; i++) {
      const extra = Array.from({ length: 5 }, () => {
        const p = rng.pick(d.profiles);
        const r = receipt(rng, p, LAST_PROMISE - 6, LAST_PROMISE);
        return { ...r, delivered: r.promised + deliveryOffset(rng, rng.chance(p.onTime)) };
      });
      rows = [...d.rows, ...extra];
      if (ratesAreDistinct(rows)) break;
    }
    return { ...d, rows };
  },
};

const vendorsReordered: Variant<ScoreData> = {
  label: 'the vendor list in column G is reordered',
  explain: 'Point the criteria at G2 instead of typing a vendor name.',
  apply: (d, rng) => {
    let vendors = rng.shuffle(d.vendors);
    while (vendors.join() === d.vendors.join()) vendors = rng.shuffle(d.vendors);
    return { ...d, vendors };
  },
};

const qtyChange: Variant<ScoreData> = {
  label: 'quantities and defects change',
  explain: 'Total the vendor’s Defects and Qty columns with SUMIFS so the rate follows the data.',
  apply: (d, rng) => {
    const rate = new Map(d.profiles.map((p) => [p.vendor, p.defect]));
    return {
      ...d,
      rows: d.rows.map((r) => {
        const qty = rng.int(4, 60) * 10;
        return { ...r, qty, defects: defectsFor(rng, qty, (rate.get(r.vendor) ?? 0.01) * rng.float(0.5, 2, 2)) };
      }),
    };
  },
};

const VENDOR_TEXT = { values: [...VENDORS], advice: 'Point to the vendor in column G instead.' };

export function ranking(d: ScoreData): ExpectedGrid {
  return d.vendors
    .map((vendor) => ({ vendor, rate: onTimeRate(d.rows, vendor) }))
    .sort((a, b) => b.rate - a.rate)
    .map((x) => [x.vendor, x.rate]);
}

export const vendorScorecard = defineMission<ScoreData>({
  id: 'mission-vendor-scorecard',
  title: 'Vendor on-time scorecard',
  role: 'ops',
  summary: 'A supplier scorecard: volume, on-time rate, defect rate and a ranking',
  minutes: 18,
  skills: ['sumifs-warehouse', 'filter-late', 'unique-vendors'],
  brief: () => ({
    from: 'Marcus Bell, Operations manager',
    subject: 'Vendor scorecard for next week’s reviews',
    body: [
      'Hi,',
      'We sit down with our six suppliers next week and I want numbers, not impressions. The `Deliveries` Table has every delivery promised for August and September.',
      'For each vendor in `G2:G7` I need the number of deliveries, the share that arrived on or before the promised date, and the defect rate (defective units over units received). Then give me one ranked list, best on-time vendor first, that I can drop into the deck.',
      'Marcus',
    ].join('\n\n'),
  }),
  make: (rng) => {
    const profiles = VENDORS.map((vendor) => ({ vendor, onTime: rng.float(0.5, 0.95, 2), defect: rng.float(0.002, 0.03, 4) }));
    const rows = profiles.flatMap((p) => Array.from({ length: rng.int(8, 13) }, () => receipt(rng, p))).sort((a, b) => a.promised - b.promised);
    return { profiles, rows: setDelivered(rng, rows, profiles), vendors: rng.shuffle(VENDORS) };
  },
  // On-time % sits right beside Vendor so the ranking is one SORTBY over G2:H7.
  blocks: (d) => [
    dataBlock('Deliveries', 'A1', RECEIPT_COLS, receiptGrid(d.rows)),
    cells('G1', [['Vendor', 'On-time %', 'Deliveries', 'Defect rate']], 'header'),
    cells('G2', column(d.vendors), 'input'),
    cells('L1', [['Vendor', 'On-time %']], 'header'),
  ],
  inputs: (d) => [tableWrite('Deliveries', RECEIPT_COLS, receiptGrid(d.rows)), rangeWrite('G2:G7', column(d.vendors))],
  steps: [
    {
      title: 'Deliveries per vendor',
      task: () =>
        'In `I2:I7`, count how many deliveries each vendor in `G2:G7` made, using the `Deliveries` Table. Write one formula in `I2` and fill it down.',
      hints: [
        'COUNTIFS counts the rows that meet every condition you give it, as pairs of (column to test, value to match).',
        'Test Deliveries[Vendor] against G2, not a typed vendor name: `=COUNTIFS(Deliveries[Vendor], G2)`, then fill down to I7.',
      ],
      solution: () => '=COUNTIFS(Deliveries[Vendor],G2)',
      answer: () => ({ kind: 'cells', range: 'I2:I7', format: FMT.int, consistency: 'all' }),
      expected: (d) => d.vendors.map((v) => [vendorTally(d.rows, v).count]),
      variants: [moreDeliveries, vendorsReordered],
      rules: {
        require: [{ pattern: /COUNTIFS?\(/i, label: 'Counts with COUNTIFS', advice: 'COUNTIFS(Deliveries[Vendor], G2) counts the vendor’s rows.' }],
        forbidText: VENDOR_TEXT,
        allowNumbers: [],
      },
    },
    {
      title: 'On-time rate',
      task: () =>
        'In `H2:H7`, show the share of each vendor’s deliveries that arrived on time, meaning Delivered is on or before Promised (arriving on the promised date counts). COUNTIFS can’t compare two columns, so count the on-time rows with SUMPRODUCT rather than adding a helper column, then divide by the count in column `I`. Leave the result as a fraction (0.75, not 75); the cells are already formatted as percentages. Write one formula in `H2` and fill it down.',
      hints: [
        '`Deliveries[Delivered]<=Deliveries[Promised]` compares the two columns row by row, giving TRUE or FALSE for each delivery.',
        'Multiply it by `(Deliveries[Vendor]=G2)` so only this vendor’s on-time rows become 1, and let SUMPRODUCT add them up.',
        '`=SUMPRODUCT((Deliveries[Vendor]=G2)*(Deliveries[Delivered]<=Deliveries[Promised]))/I2`',
      ],
      solution: () => '=SUMPRODUCT((Deliveries[Vendor]=G2)*(Deliveries[Delivered]<=Deliveries[Promised]))/I2',
      answer: () => ({ kind: 'cells', range: 'H2:H7', format: FMT.pct, consistency: 'all' }),
      expected: (d) => d.vendors.map((v) => [onTimeRate(d.rows, v)]),
      variants: [scoreDatesChange, moreDeliveries, vendorsReordered],
      rules: {
        require: [
          {
            pattern: /SUMPRODUCT\(/i,
            label: 'Counts on-time rows with SUMPRODUCT',
            advice: 'SUMPRODUCT((Deliveries[Vendor]=G2)*(Deliveries[Delivered]<=Deliveries[Promised])) counts them without a helper column.',
          },
        ],
        forbidText: VENDOR_TEXT,
        allowNumbers: [0, 1],
      },
    },
    {
      title: 'Defect rate',
      task: () =>
        'In `J2:J7`, show each vendor’s defect rate: its total Defects divided by its total Qty. Divide the two totals rather than averaging each delivery’s rate. Write one formula in `J2` and fill it down.',
      hints: [
        'You need two totals for the vendor in G2: the sum of Defects and the sum of Qty. SUMIFS gives each one.',
        '`=SUMIFS(Deliveries[Defects], Deliveries[Vendor], G2)/SUMIFS(Deliveries[Qty], Deliveries[Vendor], G2)`',
      ],
      solution: () => '=SUMIFS(Deliveries[Defects],Deliveries[Vendor],G2)/SUMIFS(Deliveries[Qty],Deliveries[Vendor],G2)',
      answer: () => ({ kind: 'cells', range: 'J2:J7', format: '0.00%', consistency: 'all' }),
      expected: (d) =>
        d.vendors.map((v) => {
          const t = vendorTally(d.rows, v);
          return [t.defects / t.qty];
        }),
      variants: [qtyChange, moreDeliveries],
      rules: { forbidText: VENDOR_TEXT, allowNumbers: [] },
    },
    {
      title: 'Rank vendors by on-time rate',
      task: () =>
        'With one formula in `L2`, list the vendors from best to worst on-time rate in two columns: the vendor name and its on-time % from column `H`, highest rate first. No two vendors share a rate, so there is one right order.',
      hints: [
        'SORTBY returns a range sorted by the values in another range: SORTBY(what to return, sort by, order).',
        'Return G2:H7 and sort it by H2:H7. An order of -1 puts the highest rate first.',
        '`=SORTBY(G2:H7, H2:H7, -1)`',
      ],
      solution: () => '=SORTBY(G2:H7,H2:H7,-1)',
      answer: () => ({ kind: 'spill', anchor: 'L2', formats: [undefined, FMT.pct] }),
      expected: ranking,
      variants: [scoreDatesChange, moreDeliveries],
      rules: {
        require: [{ pattern: /SORT(BY)?\(/i, label: 'Sorts with SORTBY', advice: 'SORTBY(G2:H7, H2:H7, -1) sorts the scorecard by on-time rate, highest first.' }],
      },
    },
  ],
});

export const FINANCE_MISSIONS: Mission<any>[] = [budgetVsActual, vendorScorecard];

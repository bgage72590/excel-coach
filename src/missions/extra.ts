import { CARRIERS, VENDORS, WAREHOUSES, eomonth, excelTextCompare, fromSerial, serial, sum } from '../engine/data';
import { round, type Rng } from '../engine/rng';
import type { AnswerArea, ColumnSpec, Grid, Inspection, Variant } from '../engine/types';
import { FMT, cells, column, dataBlock, rangeWrite, tableWrite } from '../exercises/common';
import { defineMission, type Mission } from './types';

const MONTH_NAMES = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];

/** "September 30, 2026", for briefs. */
function longDate(s: number): string {
  const { year, month, day } = fromSerial(s);
  return `${MONTH_NAMES[month - 1]} ${day}, ${year}`;
}

/** "Sep 21", for step notes. */
function shortDate(s: number): string {
  const { month, day } = fromSerial(s);
  return `${MONTH_NAMES[month - 1].slice(0, 3)} ${day}`;
}

/** A shuffle that is guaranteed to change the order. */
function reshuffle(rng: Rng, items: readonly string[]): string[] {
  let next = rng.shuffle(items);
  while (next.join('|') === items.join('|')) next = rng.shuffle(items);
  return next;
}

/** Whole dollars, for limits. */
const DOLLARS = '$#,##0';

// =====================================================================================
// Mission 1: AP aging report
// =====================================================================================

export interface ApInvoice {
  vendor: string;
  invoice: string;
  date: number;
  due: number;
  amount: number;
  paid: boolean;
}

export interface AgingData {
  invoices: ApInvoice[];
  /** Report date (K1). */
  asOf: number;
  /** Over-60 balance that puts a vendor on the call list (K2). */
  limit: number;
  /** Vendor order in J13:J18. */
  vendors: string[];
}

export interface Bucket {
  min: number;
  label: string;
}

/**
 * The Buckets Table. The labels carry "days" so Excel can never read "1–30" as a date, whether it
 * sits in a cell or in a SUMIFS criterion.
 */
export const BUCKETS: readonly Bucket[] = [
  { min: 0, label: 'Current' },
  { min: 1, label: '1–30 days' },
  { min: 31, label: '31–60 days' },
  { min: 61, label: '61–90 days' },
  { min: 91, label: 'Over 90 days' },
];

/** Days past due on either side of every bucket edge. Every sheet has an open invoice on each one. */
export const BOUNDARY_DAYS = [0, 1, 30, 31, 60, 61, 90, 91] as const;

export const AS_OF_DATES = [serial(2026, 7, 31), serial(2026, 8, 31), serial(2026, 9, 30)];

/**
 * Call-list limits the data is built around. Balances are in cents, so the only way float noise
 * could flip a "more than" test is a balance landing on the limit itself; every state keeps each
 * vendor's over-60 balance at least LIMIT_MARGIN away from the limit in force.
 */
export const OVER_60_LIMITS = [2500, 5000, 7500, 10000, 15000, 20000, 25000];
export const LIMIT_MARGIN = 1;

const AGING_COLS: ColumnSpec[] = [
  { header: 'Vendor' },
  { header: 'Invoice' },
  { header: 'Invoice date', format: FMT.date },
  { header: 'Due date', format: FMT.date },
  { header: 'Amount', format: FMT.currency },
  // Text, formatted #,##0 on purpose: the Days past due column the learner types in G1 picks up
  // this format, so it shows whole days instead of Excel formatting a date difference as a date.
  { header: 'Paid', format: FMT.int },
];
const BUCKET_COLS: ColumnSpec[] = [{ header: 'Min days', format: FMT.int }, { header: 'Bucket' }];

const invoiceGrid = (rows: ApInvoice[]): Grid => rows.map((r) => [r.vendor, r.invoice, r.date, r.due, r.amount, r.paid ? 'Y' : 'N']);
const bucketGrid = (): Grid => BUCKETS.map((b) => [b.min, b.label]);

/** What the Days past due column shows: "" for a paid invoice, otherwise days after the due date, never below 0. */
export const daysPastDue = (inv: ApInvoice, asOf: number): number | '' => (inv.paid ? '' : Math.max(0, asOf - inv.due));

/**
 * What XLOOKUP(days, Buckets[Min days], Buckets[Bucket], , -1) returns: the bucket with the largest
 * Min days at or below `days`. Exactly 30 days is still 1–30; 31 starts 31–60.
 */
export function bucketFor(days: number): string {
  let best: Bucket | undefined;
  for (const b of BUCKETS) if (b.min <= days && (!best || b.min > best.min)) best = b;
  if (!best) throw new Error(`${days} days falls below the first bucket`);
  return best.label;
}

export const bucketOf = (inv: ApInvoice, asOf: number): string => (inv.paid ? '' : bucketFor(Math.max(0, asOf - inv.due)));

/** SUMIFS(Invoices[Amount], Invoices[Vendor], vendor, Invoices[Bucket], bucket) for every vendor and bucket. */
export function agingGrid(d: AgingData): number[][] {
  return d.vendors.map((v) => BUCKETS.map((b) => sum(d.invoices.filter((i) => i.vendor === v && bucketOf(i, d.asOf) === b.label).map((i) => i.amount))));
}

/** Each bucket's open balance over the whole open balance. */
export function bucketShares(d: AgingData): number[] {
  const open = d.invoices.filter((i) => !i.paid);
  const total = sum(open.map((i) => i.amount));
  return BUCKETS.map((b) => sum(open.filter((i) => bucketOf(i, d.asOf) === b.label).map((i) => i.amount)) / total);
}

/** A vendor's open balance more than 60 days past due: its 61–90 and Over 90 buckets together. */
export function over60(d: AgingData, vendor: string): number {
  return sum(d.invoices.filter((i) => i.vendor === vendor && !i.paid && d.asOf - i.due > 60).map((i) => i.amount));
}

/** Vendors whose over-60 balance is more than the limit, in the order of column J. */
export const callList = (d: AgingData, limit = d.limit): string[] => d.vendors.filter((v) => over60(d, v) > limit);

/** The limit splits the vendors (someone to call, someone not) and no balance sits within LIMIT_MARGIN of it. */
export function limitWorks(d: AgingData, limit = d.limit): boolean {
  const totals = d.vendors.map((v) => over60(d, v));
  const over = totals.filter((t) => t > limit).length;
  return over >= 1 && over < totals.length && totals.every((t) => Math.abs(t - limit) >= LIMIT_MARGIN);
}

const workingLimits = (d: AgingData) => OVER_60_LIMITS.filter((l) => limitWorks(d, l));

/** "Apex Supply" → "AS", the prefix on that vendor's invoice numbers. */
const initials = (vendor: string) =>
  vendor
    .split(' ')
    .map((w) => w[0])
    .join('');

const invoiceAmount = (rng: Rng) => rng.float(180, 16_000, 2);

type DraftInvoice = Omit<ApInvoice, 'invoice'>;

/**
 * Two slow payers carry most of the old balances; the rest pay close to terms. On top of that,
 * one open invoice sits on each day in BOUNDARY_DAYS and two more aren't due yet.
 */
function draftInvoices(rng: Rng, asOf: number): DraftInvoice[] {
  const slow = new Set<string>(rng.sample(VENDORS, 2));
  const terms = new Map<string, number>(VENDORS.map((v) => [v, rng.pick([15, 30, 30, 45, 60])]));
  const drafts: DraftInvoice[] = [];
  const add = (vendor: string, due: number, paid: boolean) => drafts.push({ vendor, date: due - terms.get(vendor)!, due, amount: invoiceAmount(rng), paid });

  for (const vendor of VENDORS) {
    for (let n = rng.int(5, 7); n > 0; n--) {
      const due = asOf - rng.int(3, 150) + terms.get(vendor)!;
      add(vendor, due, rng.chance(due > asOf ? 0.1 : slow.has(vendor) ? 0.35 : 0.9));
    }
  }
  const regular = drafts.length;
  for (const days of BOUNDARY_DAYS) add(rng.pick(VENDORS), asOf - days, false);
  for (let n = 0; n < 2; n++) add(rng.pick(VENDORS), asOf + rng.int(1, 14), false);

  // A paid invoice past its due date, so a formula that ignores Paid shows a number there.
  if (!drafts.some((x) => x.paid && x.due < asOf)) {
    const k = drafts.slice(0, regular).findIndex((x) => x.due < asOf);
    if (k >= 0) drafts[k].paid = true;
  }
  return drafts;
}

/** The export comes sorted by invoice date, with each vendor's invoice numbers rising. */
function numberInvoices(rng: Rng, drafts: DraftInvoice[]): ApInvoice[] {
  const next = new Map<string, number>(VENDORS.map((v) => [v, rng.int(10_000, 80_000)]));
  return drafts
    .slice()
    .sort((a, b) => a.date - b.date || excelTextCompare(a.vendor, b.vendor))
    .map((x) => {
      const no = next.get(x.vendor)! + rng.int(1, 40);
      next.set(x.vendor, no);
      return { vendor: x.vendor, invoice: `${initials(x.vendor)}-${no}`, date: x.date, due: x.due, amount: x.amount, paid: x.paid };
    });
}

/** Marks three open invoices paid, keeping the call list meaningful. */
function postPayments(d: AgingData, rng: Rng): AgingData {
  const open = d.invoices.flatMap((inv, i) => (inv.paid ? [] : [i]));
  const pay = (idx: number[]): AgingData => ({ ...d, invoices: d.invoices.map((inv, i) => (idx.includes(i) ? { ...inv, paid: true } : inv)) });
  for (let attempt = 0; attempt < 40; attempt++) {
    const next = pay(rng.sample(open, 3));
    if (limitWorks(next)) return next;
  }
  // Paying invoices 60 days past due or less leaves every over-60 balance where it was.
  return pay(open.filter((i) => d.asOf - d.invoices[i].due <= 60).slice(0, 3));
}

/** Four more open invoices of mixed ages, appended the way the next export would add them. */
function addInvoices(d: AgingData, rng: Rng): AgingData {
  const used = new Set(d.invoices.map((i) => i.invoice));
  const added = Array.from({ length: 4 }, (): ApInvoice => {
    const vendor = rng.pick(d.vendors);
    let invoice = `${initials(vendor)}-${rng.int(90_000, 99_999)}`;
    while (used.has(invoice)) invoice = `${initials(vendor)}-${rng.int(90_000, 99_999)}`;
    used.add(invoice);
    const date = d.asOf - rng.int(0, 120);
    return { vendor, invoice, date, due: date + rng.pick([15, 30, 45]), amount: invoiceAmount(rng), paid: false };
  });
  return { ...d, invoices: [...d.invoices, ...added] };
}

const asOfMoves: Variant<AgingData> = {
  label: 'the As of date in K1 changes',
  explain: 'Subtract from $K$1 instead of typing the date or using TODAY().',
  apply: (d, rng) => ({ ...d, asOf: d.asOf + rng.pick([7, 14, 21, 28]) }),
};

const paymentsPosted: Variant<AgingData> = {
  label: 'more invoices are paid',
  explain: 'Test the Paid column on the same row, so an invoice drops out once it’s paid.',
  apply: postPayments,
};

const invoicesAdded: Variant<AgingData> = {
  label: 'four invoices are added to the Table',
  explain: 'A calculated column fills new rows by itself. Formulas typed cell by cell don’t.',
  apply: addInvoices,
};

const vendorsReordered: Variant<AgingData> = {
  label: 'the vendors in column J are reordered',
  explain: 'Point at the vendor in column J instead of typing its name.',
  apply: (d, rng) => ({ ...d, vendors: reshuffle(rng, d.vendors) }),
};

const amountsChange: Variant<AgingData> = {
  label: 'invoice amounts change',
  apply: (d, rng) => ({ ...d, invoices: d.invoices.map((i) => ({ ...i, amount: invoiceAmount(rng) })) }),
};

const limitChanges: Variant<AgingData> = {
  label: 'the limit in K2 changes',
  explain: 'Compare with K2 instead of typing the limit into the formula.',
  apply: (d, rng) => {
    // make() guarantees another working limit that changes the list, so a typed limit can't pass by luck.
    const current = callList(d).join('|');
    const others = rng.shuffle(workingLimits(d).filter((l) => l !== d.limit));
    const pick = others.find((l) => callList(d, l).join('|') !== current) ?? others[0];
    return { ...d, limit: pick };
  },
};

const invoicesColumn = (name: string, format?: string): AnswerArea => ({ kind: 'tableColumn', table: 'Invoices', column: name, format });
const hasInvoicesColumn = (name: string): Inspection[] => [{ kind: 'tableColumn', table: 'Invoices', column: name, label: `The Invoices Table has a ${name} column` }];

const BUCKET_LABELS = BUCKETS.map((b) => b.label);
const AGING_VENDOR_TEXT = { values: [...VENDORS], advice: 'Point to the vendor in column J instead.' };

// A fixed range that exactly covers the Table grows with it, so the added-invoices variant can't catch it.
const INVOICES_RULE = {
  pattern: /Invoices\[/i,
  label: 'Uses the Invoices Table’s columns',
  advice: 'Total Invoices[Amount] by Invoices[Vendor] and Invoices[Bucket] instead of cell ranges, so the grid follows the Table as each export adds invoices.',
};

export const apAging = defineMission<AgingData>({
  id: 'mission-ap-aging',
  title: 'AP aging report',
  role: 'finance',
  summary: 'The month-end AP aging: days past due, buckets, balances by vendor and a call list',
  minutes: 22,
  skills: ['tables-calc-column', 'eomonth-terms', 'xlookup-tiered', 'sumifs-grid', 'filter-late'],
  brief: (d) => ({
    from: 'Grace Liu, Assistant controller',
    subject: `AP aging as of ${longDate(d.asOf)}`,
    body: [
      'Hi,',
      '',
      `Before Friday’s cash meeting I need the AP aging as of ${longDate(d.asOf)}. The \`Invoices\` Table is this morning’s export: every vendor invoice from the last few months, with a Paid flag.`,
      '',
      '1. Days past due for each open invoice, as of the date in `K1`.',
      '2. Its aging bucket, using the cutoffs in the `Buckets` Table.',
      '3. Open balance by vendor and bucket in `K13:O18`.',
      '4. Each bucket’s share of total open AP in `K20:O20`.',
      '5. The vendors with more than the limit in `K2` over 60 days past due. I’ll call those first.',
      '',
      'Build it off `K1` so I can rerun it at the next month-end.',
      '',
      'Thanks,',
      'Grace',
    ].join('\n'),
  }),
  make: (rng) => {
    const asOf = rng.pick(AS_OF_DATES);
    const vendors = rng.shuffle(VENDORS);
    // Redraw until two limits give different call lists, so the limit variant always has somewhere
    // to go. The first draw nearly always qualifies.
    let d: AgingData = { invoices: [], asOf, limit: OVER_60_LIMITS[0], vendors };
    for (let attempt = 0; attempt < 100; attempt++) {
      d = { ...d, invoices: numberInvoices(rng, draftInvoices(rng, asOf)) };
      const limits = workingLimits(d);
      if (new Set(limits.map((l) => callList(d, l).join('|'))).size >= 2) return { ...d, limit: rng.pick(limits) };
    }
    return d;
  },
  // Invoices fills A:F and the learner adds G and H, so column I stays empty and everything else
  // starts in J. The Buckets Table keeps a blank row and column around it too.
  blocks: (d) => [
    dataBlock('Invoices', 'A1', AGING_COLS, invoiceGrid(d.invoices)),
    cells('J1', [['As of'], ['Over-60 limit']], 'label'),
    cells('K1', [[d.asOf]], 'input', FMT.date),
    cells('K2', [[d.limit]], 'input', DOLLARS),
    dataBlock('Buckets', 'J4', BUCKET_COLS, bucketGrid()),
    cells('J12', [['Vendor', ...BUCKET_LABELS]], 'header'),
    cells('J13', column(d.vendors), 'input'),
    cells('J20', [['% of open AP']], 'label'),
    cells('J22', [['Vendors to call']], 'header'),
  ],
  inputs: (d) => [tableWrite('Invoices', AGING_COLS, invoiceGrid(d.invoices)), rangeWrite('K1:K2', [[d.asOf], [d.limit]]), rangeWrite('J13:J18', column(d.vendors))],
  steps: [
    // ---------- 1. Date arithmetic in a calculated column ----------
    {
      title: 'Days past due',
      task: () =>
        'Add a column named `Days past due` to the `Invoices` Table by typing the name in `G1`. For each open invoice (Paid is `N`), show how many days past its due date it is as of the date in `K1`, or 0 if it isn’t due yet. Paid invoices show an empty string (""). Point to `K1` instead of typing the date, so the column updates when the date changes.',
      hints: [
        'Type Days past due in G1 and press {enter}. The Table grows to take in the new column, and one formula in G2 fills the whole column.',
        'Subtracting two dates gives the days between them: $K$1-[@[Due date]]. MAX(0, …) turns a negative result, an invoice that isn’t due yet, into 0. Press {absKey} to lock K1.',
        'Wrap it in IF so paid invoices show "": `=IF([@Paid]="Y", "", MAX(0, $K$1-[@[Due date]]))`',
      ],
      solution: () => '=IF([@Paid]="Y","",MAX(0,$K$1-[@[Due date]]))',
      answer: () => invoicesColumn('Days past due', FMT.int),
      expected: (d) => d.invoices.map((i) => [daysPastDue(i, d.asOf)]),
      variants: [asOfMoves, paymentsPosted, invoicesAdded],
      rules: { allowNumbers: [0] },
      inspections: () => hasInvoicesColumn('Days past due'),
    },

    // ---------- 2. Approximate-match lookup ----------
    {
      title: 'Aging bucket',
      task: () =>
        'Add a `Bucket` column in `H1` that labels each open invoice with its aging bucket from the `Buckets` Table: the bucket with the largest `Min days` at or below the invoice’s days past due. An invoice exactly 30 days past due is in `1–30 days`; one 31 days past due is in `31–60 days`. Paid invoices show an empty string (""). Look the labels up instead of typing them.',
      hints: [
        'This is an approximate match: find the largest Min days at or below the days past due. XLOOKUP’s fifth argument, match_mode, does that with -1.',
        'Look up [@[Days past due]] in Buckets[Min days] and return Buckets[Bucket], leaving the fourth argument (if_not_found) empty. Keep paid invoices out with the same IF as before.',
        '`=IF([@Paid]="Y", "", XLOOKUP([@[Days past due]], Buckets[Min days], Buckets[Bucket], , -1))`',
      ],
      solution: () => '=IF([@Paid]="Y","",XLOOKUP([@[Days past due]],Buckets[Min days],Buckets[Bucket],,-1))',
      answer: () => invoicesColumn('Bucket'),
      expected: (d) => d.invoices.map((i) => [bucketOf(i, d.asOf)]),
      variants: [{ ...asOfMoves, explain: 'Look up the Days past due column, so the bucket follows K1.' }, paymentsPosted, invoicesAdded],
      rules: {
        require: [
          {
            pattern: /XLOOKUP\(|VLOOKUP\(|LOOKUP\(|INDEX\(/i,
            label: 'Looks up the bucket in the Buckets Table',
            advice: 'Look up the days in Buckets[Min days] with XLOOKUP and match_mode -1, which returns an exact match or else the next smaller value.',
          },
        ],
        forbidText: { values: BUCKET_LABELS, advice: 'Return Buckets[Bucket] instead of typing the labels.' },
        allowNumbers: [0, 1, 2],
      },
      inspections: () => hasInvoicesColumn('Bucket'),
    },

    // ---------- 3. SUMIFS grid ----------
    {
      title: 'Open balance by vendor and bucket',
      task: () =>
        'Fill `K13:O18` with the open balance for each vendor (column `J`) in each aging bucket (row `12`): the total Amount of that vendor’s invoices in that bucket. Paid invoices have no bucket, so they drop out. Write one formula in `K13` that still works when you fill it right and down.',
      hints: [
        'Use SUMIFS with two conditions: Invoices[Vendor] matched against the vendor in column J, and Invoices[Bucket] matched against the bucket in row 12.',
        'Lock the column on the vendor ($J13) and the row on the bucket (K$12) so both stay put as you fill. Press {absKey} to cycle the $ signs.',
        'Filling right also slides Table column names one column, so Invoices[Amount] would become Invoices[Paid]. Writing a name twice, Invoices[[Amount]:[Amount]], locks it the way $ locks a cell.',
        '`=SUMIFS(Invoices[[Amount]:[Amount]], Invoices[[Vendor]:[Vendor]], $J13, Invoices[[Bucket]:[Bucket]], K$12)`, then fill right to column O and down to row 18.',
      ],
      // Doubled column names: the fill handle shifts Invoices[Amount] to Invoices[Paid] one cell to the right.
      solution: () => '=SUMIFS(Invoices[[Amount]:[Amount]],Invoices[[Vendor]:[Vendor]],$J13,Invoices[[Bucket]:[Bucket]],K$12)',
      answer: () => ({ kind: 'cells', range: 'K13:O18', format: FMT.currency, consistency: 'all' }),
      expected: agingGrid,
      variants: [
        { ...asOfMoves, explain: 'Total the Bucket column, so the grid follows the As of date in K1.' },
        { ...paymentsPosted, explain: 'Match on the Bucket column, which is empty for paid invoices, so payments drop out.' },
        vendorsReordered,
        { ...invoicesAdded, explain: 'Use whole Table columns like Invoices[Amount] so new invoices are counted.' },
      ],
      rules: {
        require: [INVOICES_RULE],
        forbidText: { values: [...VENDORS, ...BUCKET_LABELS], advice: 'Point to the vendor in column J and the bucket in row 12 instead.' },
        allowNumbers: [],
      },
    },

    // ---------- 4. Share of total ----------
    {
      title: 'Share of open AP by bucket',
      task: () =>
        'In `K20:O20`, show each bucket’s share of total open AP: the bucket’s column of the grid divided by the whole grid. Leave the result as a fraction (0.25, not 25); the cells are already formatted as percentages. Write one formula in `K20` and fill it right. The five shares add up to 100%.',
      hints: [
        'The top of the fraction is the bucket’s column of the grid above: SUM(K13:K18).',
        'The bottom is the whole grid, locked so it doesn’t slide as you fill right: SUM($K$13:$O$18).',
        '`=SUM(K13:K18)/SUM($K$13:$O$18)`, then fill right to O20.',
      ],
      solution: () => '=SUM(K13:K18)/SUM($K$13:$O$18)',
      answer: () => ({ kind: 'cells', range: 'K20:O20', format: FMT.pct, consistency: 'all' }),
      expected: (d) => [bucketShares(d)],
      variants: [{ ...asOfMoves, explain: 'Build the shares from the grid or the Bucket column, so they follow the As of date.' }, paymentsPosted, amountsChange],
      rules: { allowNumbers: [] },
    },

    // ---------- 5. FILTER on a computed test ----------
    {
      title: 'Vendors to call',
      task: () =>
        'With one formula in `J23`, list the vendors whose balance more than 60 days past due (the `61–90 days` and `Over 90 days` buckets together) is more than the limit in `K2`. Keep them in the order of column `J`. The header is already in `J22`, so return the names only.',
      hints: [
        'FILTER returns the entries of a list that pass a test. The list here is the vendors in J13:J18.',
        'Each vendor’s over-60 balance is its 61–90 days cell plus its Over 90 days cell, so N13:N18+O13:O18 gives all six at once. Compare that with K2.',
        '`=FILTER(J13:J18, N13:N18+O13:O18>K2)`',
      ],
      solution: () => '=FILTER(J13:J18,N13:N18+O13:O18>K2)',
      answer: () => ({ kind: 'spill', anchor: 'J23' }),
      expected: (d) => callList(d).map((v) => [v]),
      variants: [limitChanges, { ...paymentsPosted, explain: 'Build the test from the grid or the Invoices Table, so payments change the list.' }, vendorsReordered],
      rules: {
        require: [{ pattern: /FILTER\(/i, label: 'Uses FILTER', advice: 'FILTER(J13:J18, include) returns every vendor that passes the test in one go.' }],
        forbidText: AGING_VENDOR_TEXT,
        allowNumbers: [0, 1],
      },
    },
  ],
});

// =====================================================================================
// Freight data shared by missions 2 and 3
// =====================================================================================

const DESTINATIONS = ['Denver', 'Phoenix', 'Chicago', 'Memphis', 'Charlotte', 'Seattle', 'Kansas City', 'Salt Lake City'] as const;

/** Every lane name the data can use, for the typed-text rule. */
export const ALL_LANES = WAREHOUSES.flatMap((from) => DESTINATIONS.map((to) => `${from} → ${to}`));

export interface Lane {
  name: string;
  /** Base rate in dollars per lb, before the carrier's pricing. */
  rate: number;
}

function makeLanes(rng: Rng, count: number): Lane[] {
  return rng.sample(ALL_LANES, count).map((name) => ({ name, rate: rng.float(0.12, 0.42, 3) }));
}

/** LTL shipments, from a couple of pallets to a partial truckload, in 50 lb steps. */
const shipmentWeight = (rng: Rng) => rng.int(3, 120) * 50;

/** Weight × lane rate × the carrier's pricing, with some noise for fuel and accessorials. */
const freightCost = (rng: Rng, weight: number, lane: Lane, factor: number) => round(weight * lane.rate * factor * rng.float(0.9, 1.1, 3), 2);

// =====================================================================================
// Mission 2: Carrier cost review
// =====================================================================================

export interface Shipment {
  carrier: string;
  lane: string;
  ship: number;
  promised: number;
  delivered: number;
  weight: number;
  cost: number;
}

interface CarrierProfile {
  carrier: string;
  /** Pricing against the lane rate: 1.1 is 10% dearer. */
  factor: number;
  /** Chance a shipment arrives on or before the promised date. */
  onTime: number;
}

export interface CarrierData {
  profiles: CarrierProfile[];
  lanes: Lane[];
  rows: Shipment[];
  /** Carrier order in I2:I5 and I11:I14. */
  carriers: string[];
  /** Month starts in J10:L10. */
  months: number[];
}

const SHIPMENT_COLS: ColumnSpec[] = [
  { header: 'Carrier' },
  { header: 'Lane' },
  { header: 'Ship date', format: FMT.date },
  { header: 'Promised date', format: FMT.date },
  { header: 'Delivered date', format: FMT.date },
  { header: 'Weight lb', format: FMT.int },
  { header: 'Cost', format: FMT.currency },
];

const shipmentGrid = (rows: Shipment[]): Grid => rows.map((r) => [r.carrier, r.lane, r.ship, r.promised, r.delivered, r.weight, r.cost]);

/** Four months of history; the trend grid shows three of them at a time. */
export const REVIEW_MONTHS = [serial(2026, 6, 1), serial(2026, 7, 1), serial(2026, 8, 1), serial(2026, 9, 1)];
const LAST_THREE = REVIEW_MONTHS.slice(1);
const FIRST_THREE = REVIEW_MONTHS.slice(0, 3);

/** The priciest lane leads the next one by at least 3%, so there is one right answer. */
export const TOP_LANE_LEAD = 1.03;

const PER_LB = '$0.000';

/** Days from promised to delivered: on the day or early when on time, 1–4 days late otherwise. */
const deliveryOffset = (rng: Rng, onTime: boolean) => (onTime ? (rng.chance(0.4) ? 0 : -rng.int(1, 2)) : rng.int(1, 4));

function shipment(rng: Rng, d: Pick<CarrierData, 'profiles' | 'lanes'>, ship: number): Shipment {
  const p = rng.pick(d.profiles);
  const lane = rng.pick(d.lanes);
  const weight = shipmentWeight(rng);
  const promised = ship + rng.int(2, 5);
  return { carrier: p.carrier, lane: lane.name, ship, promised, delivered: promised + deliveryOffset(rng, rng.chance(p.onTime)), weight, cost: freightCost(rng, weight, lane, p.factor) };
}

/** Moves a shipment to another ship date, keeping its transit time and lateness. */
function shipOn(r: Shipment, ship: number): Shipment {
  const shift = ship - r.ship;
  return { ...r, ship, promised: r.promised + shift, delivered: r.delivered + shift };
}

/** Every carrier with on-time shipments gets one delivered on the promised day, so < and <= give different rates. */
function withOnTheDay(rows: Shipment[]): Shipment[] {
  const out = rows.map((r) => ({ ...r }));
  for (const c of new Set(out.map((r) => r.carrier))) {
    const onTime = out.filter((r) => r.carrier === c && r.delivered <= r.promised);
    if (onTime.length && !onTime.some((r) => r.delivered === r.promised)) onTime[0].delivered = onTime[0].promised;
  }
  return out;
}

export function carrierTally(rows: Shipment[], carrier: string) {
  const mine = rows.filter((r) => r.carrier === carrier);
  return {
    count: mine.length,
    onTime: mine.filter((r) => r.delivered <= r.promised).length,
    cost: sum(mine.map((r) => r.cost)),
    weight: sum(mine.map((r) => r.weight)),
  };
}

export function costPerLb(rows: Shipment[], carrier: string): number {
  const t = carrierTally(rows, carrier);
  return t.cost / t.weight;
}

export function onTimeShare(rows: Shipment[], carrier: string): number {
  const t = carrierTally(rows, carrier);
  return t.onTime / t.count;
}

/** Every lane in the data with its cost per lb (total cost ÷ total weight), priciest first. */
export function laneRanking(rows: Shipment[]): { lane: string; perLb: number }[] {
  return [...new Set(rows.map((r) => r.lane))]
    .map((lane) => {
      const mine = rows.filter((r) => r.lane === lane);
      return { lane, perLb: sum(mine.map((r) => r.cost)) / sum(mine.map((r) => r.weight)) };
    })
    .sort((a, b) => b.perLb - a.perLb);
}

/** A carrier's cost for shipments dated within `month` (a month start), first and last days included. */
export function monthlyCost(rows: Shipment[], carrier: string, month: number): number {
  const end = eomonth(month);
  return sum(rows.filter((r) => r.carrier === carrier && r.ship >= month && r.ship <= end).map((r) => r.cost));
}

/** Prices the top lane up until it leads the next one by TOP_LANE_LEAD. */
function clearTopLane(rows: Shipment[]): Shipment[] {
  let out = rows;
  for (let i = 0; i < 20; i++) {
    const [top, next] = laneRanking(out);
    if (!next || top.perLb >= next.perLb * TOP_LANE_LEAD) break;
    out = out.map((r) => (r.lane === top.lane ? { ...r, cost: round(r.cost * 1.04, 2) } : r));
  }
  return out;
}

const shipCostsChange: Variant<CarrierData> = {
  label: 'shipment costs change',
  apply: (d, rng) => {
    const factor = new Map(d.profiles.map((p) => [p.carrier, p.factor]));
    const lanes = new Map(d.lanes.map((l) => [l.name, l]));
    return { ...d, rows: clearTopLane(d.rows.map((r) => ({ ...r, cost: freightCost(rng, r.weight, lanes.get(r.lane)!, factor.get(r.carrier)!) }))) };
  },
};

const moreShipments: Variant<CarrierData> = {
  label: 'new shipments are added',
  explain: 'Use whole Table columns like Shipments[Cost] so new rows are counted.',
  apply: (d, rng) => {
    const last = REVIEW_MONTHS[REVIEW_MONTHS.length - 1];
    const added = Array.from({ length: 6 }, () => shipment(rng, d, rng.int(last + 14, eomonth(last))));
    return { ...d, rows: clearTopLane([...d.rows, ...added]) };
  },
};

const carriersReordered: Variant<CarrierData> = {
  label: 'the carriers in column I are reordered',
  explain: 'Point the criteria at the carrier in column I instead of typing its name.',
  apply: (d, rng) => ({ ...d, carriers: reshuffle(rng, d.carriers) }),
};

const deliveriesChange: Variant<CarrierData> = {
  label: 'delivery dates change',
  explain: 'Compare Shipments[Delivered date] with Shipments[Promised date] row by row instead of typing counts or rates.',
  apply: (d, rng) => {
    const chance = new Map(d.profiles.map((p) => [p.carrier, p.onTime]));
    return { ...d, rows: d.rows.map((r) => ({ ...r, delivered: r.promised + deliveryOffset(rng, rng.chance(chance.get(r.carrier)!)) })) };
  },
};

const laneRateRises: Variant<CarrierData> = {
  label: 'rates on another lane go up',
  explain: 'Compare every lane’s cost per lb, from UNIQUE(Shipments[Lane]), instead of typing a lane name.',
  apply: (d, rng) => {
    const [top, ...rest] = laneRanking(d.rows);
    const pick = rng.pick(rest);
    // Enough to put this lane 8–20% ahead of the old leader.
    const factor = (top.perLb / pick.perLb) * rng.float(1.08, 1.2, 3);
    return { ...d, rows: clearTopLane(d.rows.map((r) => (r.lane === pick.lane ? { ...r, cost: round(r.cost * factor, 2) } : r))) };
  },
};

const monthsShift: Variant<CarrierData> = {
  label: 'the months in row 10 change',
  explain: 'Build each date window from the month start in row 10 and EOMONTH instead of typing dates.',
  apply: (d) => ({ ...d, months: d.months[0] === LAST_THREE[0] ? FIRST_THREE : LAST_THREE }),
};

const CARRIER_TEXT = { values: [...CARRIERS], advice: 'Point to the carrier in column I instead.' };

// A fixed range that exactly covers the Table grows with it, so the new-shipments variant can't catch it.
const SHIPMENTS_RULE = {
  pattern: /Shipments\[/i,
  label: 'Uses the Shipments Table’s columns',
  advice: 'Refer to columns by name, like Shipments[Cost], instead of cell ranges, so the result follows the Table as shipments are added.',
};

export const carrierCostReview = defineMission<CarrierData>({
  id: 'mission-carrier-cost-review',
  title: 'Carrier cost review',
  role: 'ops',
  summary: 'A carrier review for contract talks: cost per pound, on-time rate, the priciest lane and monthly spend',
  minutes: 20,
  skills: ['sumifs-warehouse', 'sumifs-grid', 'sumifs-month', 'unique-vendors', 'let-reorder'],
  brief: () => ({
    from: 'Tom Alvarez, Logistics director',
    subject: 'Carrier numbers before the contract renewals',
    body: [
      'Hi,',
      '',
      'Our carrier contracts come up for renewal in November, and I want to walk in with numbers. The `Shipments` Table has every shipment we tendered from June through September.',
      '',
      '1. Cost per pound for each carrier in column `I`: total cost over total weight.',
      '2. Each carrier’s on-time rate: delivered on or before the promised date.',
      '3. The lane that costs us the most per pound, across all carriers, in `J8:K8`.',
      '4. Monthly spend by carrier for the last three months in `J11:L14`, so I can see who’s trending up.',
      '',
      'Thanks,',
      'Tom',
    ].join('\n'),
  }),
  make: (rng) => {
    const profiles = CARRIERS.map((carrier) => ({ carrier, factor: rng.float(0.85, 1.2, 2), onTime: rng.float(0.72, 0.96, 2) }));
    const lanes = makeLanes(rng, 7);
    const rows: Shipment[] = [];
    for (const m of REVIEW_MONTHS) {
      const month = Array.from({ length: rng.int(15, 18) }, () => shipment(rng, { profiles, lanes }, rng.int(m, eomonth(m))));
      // A shipment on the first and last day of every month, so a window built with > or < misses one.
      month[0] = shipOn(month[0], m);
      month[1] = shipOn(month[1], eomonth(m));
      rows.push(...month);
    }
    rows.sort((a, b) => a.ship - b.ship);
    return { profiles, lanes, rows: clearTopLane(withOnTheDay(rows)), carriers: rng.shuffle(CARRIERS), months: LAST_THREE };
  },
  // Shipments fills A:G; column H stays empty so nothing touches the Table.
  blocks: (d) => [
    dataBlock('Shipments', 'A1', SHIPMENT_COLS, shipmentGrid(d.rows)),
    cells('I1', [['Carrier', 'Cost per lb', 'On-time %']], 'header'),
    cells('I2', column(d.carriers), 'input'),
    cells('J7', [['Lane', 'Cost per lb']], 'header'),
    cells('I8', [['Most expensive']], 'label'),
    cells('I10', [['Carrier']], 'header'),
    cells('J10', [d.months], 'header', FMT.month),
    cells('I11', column(d.carriers), 'input'),
  ],
  inputs: (d) => [
    tableWrite('Shipments', SHIPMENT_COLS, shipmentGrid(d.rows)),
    rangeWrite('I2:I5', column(d.carriers)),
    rangeWrite('J10:L10', [d.months]),
    rangeWrite('I11:I14', column(d.carriers)),
  ],
  steps: [
    // ---------- 1. Ratio of two SUMIFS ----------
    {
      title: 'Cost per lb by carrier',
      task: () =>
        'In `J2:J5`, show each carrier’s cost per pound: the total Cost of its shipments divided by their total Weight lb. Divide the two totals rather than averaging each shipment’s rate. Write one formula in `J2` and fill it down.',
      hints: [
        'You need two totals for the carrier in I2: the sum of Cost and the sum of Weight lb. SUMIFS gives each one.',
        '`=SUMIFS(Shipments[Cost], Shipments[Carrier], I2)/SUMIFS(Shipments[Weight lb], Shipments[Carrier], I2)`, then fill down to J5.',
      ],
      solution: () => '=SUMIFS(Shipments[Cost],Shipments[Carrier],I2)/SUMIFS(Shipments[Weight lb],Shipments[Carrier],I2)',
      answer: () => ({ kind: 'cells', range: 'J2:J5', format: PER_LB, consistency: 'all' }),
      expected: (d) => d.carriers.map((c) => [costPerLb(d.rows, c)]),
      variants: [shipCostsChange, moreShipments, carriersReordered],
      rules: { require: [SHIPMENTS_RULE], forbidText: CARRIER_TEXT, allowNumbers: [] },
    },

    // ---------- 2. Comparing two columns row by row ----------
    {
      title: 'On-time rate by carrier',
      task: () =>
        'In `K2:K5`, show the share of each carrier’s shipments delivered on time, meaning the Delivered date is on or before the Promised date (delivery on the promised day counts). Leave the result as a fraction (0.9, not 90); the cells are already formatted as percentages. Write one formula in `K2` and fill it down.',
      hints: [
        'COUNTIFS can’t compare two columns row by row, but an array test can: Shipments[Delivered date]<=Shipments[Promised date] gives TRUE or FALSE for every shipment.',
        'Multiply it by (Shipments[Carrier]=I2) so only this carrier’s on-time shipments become 1, add them up with SUMPRODUCT, then divide by the carrier’s shipment count from COUNTIFS.',
        '`=SUMPRODUCT((Shipments[Carrier]=I2)*(Shipments[Delivered date]<=Shipments[Promised date]))/COUNTIFS(Shipments[Carrier], I2)`',
      ],
      solution: () => '=SUMPRODUCT((Shipments[Carrier]=I2)*(Shipments[Delivered date]<=Shipments[Promised date]))/COUNTIFS(Shipments[Carrier],I2)',
      answer: () => ({ kind: 'cells', range: 'K2:K5', format: FMT.pct, consistency: 'all' }),
      expected: (d) => d.carriers.map((c) => [onTimeShare(d.rows, c)]),
      variants: [deliveriesChange, moreShipments, carriersReordered],
      rules: { require: [SHIPMENTS_RULE], forbidText: CARRIER_TEXT, allowNumbers: [0, 1] },
    },

    // ---------- 3. Arg-max over a computed list ----------
    {
      title: 'Most expensive lane per lb',
      task: () =>
        'In `J8:K8`, return the lane with the highest cost per pound across all carriers, and that cost per pound: the lane’s total Cost divided by its total Weight lb. Work it out from the `Shipments` Table without helper columns. One formula in `J8` can return both cells, or you can write one formula in each.',
      hints: [
        'UNIQUE(Shipments[Lane]) lists each lane once. SUMIFS accepts that whole list as its criteria and returns one total per lane, so dividing two SUMIFS gives every lane’s cost per lb at once.',
        'Name the pieces with LET, then pick the top lane: XLOOKUP(MAX(perLb), perLb, lanes) returns its name and MAX(perLb) its cost per lb. To return both from one formula, sort the lanes with SORTBY and keep the first row with TAKE.',
        '`=LET(lanes, UNIQUE(Shipments[Lane]), perLb, SUMIFS(Shipments[Cost], Shipments[Lane], lanes)/SUMIFS(Shipments[Weight lb], Shipments[Lane], lanes), TAKE(SORTBY(HSTACK(lanes, perLb), perLb, -1), 1))`',
      ],
      solution: () =>
        '=LET(lanes,UNIQUE(Shipments[Lane]),perLb,SUMIFS(Shipments[Cost],Shipments[Lane],lanes)/SUMIFS(Shipments[Weight lb],Shipments[Lane],lanes),TAKE(SORTBY(HSTACK(lanes,perLb),perLb,-1),1))',
      answer: () => ({ kind: 'cells', range: 'J8:K8', format: PER_LB, consistency: 'none', spillOk: true }),
      expected: (d) => {
        const [top] = laneRanking(d.rows);
        return [[top.lane, top.perLb]];
      },
      variants: [laneRateRises, moreShipments],
      rules: {
        require: [SHIPMENTS_RULE],
        forbidText: { values: ALL_LANES, advice: 'Find the lane with a lookup over the Shipments Table instead of typing it.' },
        allowNumbers: [0, 1, 2],
      },
    },

    // ---------- 4. SUMIFS grid with a date window ----------
    {
      title: 'Monthly spend by carrier',
      task: () =>
        'Fill `J11:L14` with each carrier’s total Cost for shipments in each month. Row `10` holds the first day of each month; a shipment counts in the month its Ship date falls in, first and last days included. Write one formula in `J11` that still works when you fill it right and down, and build each date window from row `10` so the grid follows when the months change.',
      hints: [
        'Use SUMIFS with three conditions: the carrier in column I, Ship date on or after the month start in row 10, and Ship date on or before that month’s last day, EOMONTH(J$10, 0).',
        'Lock the column on the carrier ($I11) and the row on the month (J$10). Build each date condition by joining an operator to the cell: ">="&J$10.',
        'Filling right also slides Table column names one column, so Shipments[Carrier] would become Shipments[Lane]. Writing a name twice, Shipments[[Carrier]:[Carrier]], locks it the way $ locks a cell.',
        '`=SUMIFS(Shipments[[Cost]:[Cost]], Shipments[[Carrier]:[Carrier]], $I11, Shipments[[Ship date]:[Ship date]], ">="&J$10, Shipments[[Ship date]:[Ship date]], "<="&EOMONTH(J$10, 0))`',
      ],
      // Doubled column names: the fill handle shifts Shipments[Carrier] to Shipments[Lane] one cell to the right.
      solution: () =>
        '=SUMIFS(Shipments[[Cost]:[Cost]],Shipments[[Carrier]:[Carrier]],$I11,Shipments[[Ship date]:[Ship date]],">="&J$10,Shipments[[Ship date]:[Ship date]],"<="&EOMONTH(J$10,0))',
      answer: () => ({ kind: 'cells', range: 'J11:L14', format: FMT.currency, consistency: 'all' }),
      expected: (d) => d.carriers.map((c) => d.months.map((m) => monthlyCost(d.rows, c, m))),
      variants: [monthsShift, shipCostsChange, carriersReordered, moreShipments],
      rules: { require: [SHIPMENTS_RULE], forbidText: CARRIER_TEXT, allowNumbers: [0, 1] },
    },
  ],
});

// =====================================================================================
// Mission 3: Weekly carrier summary refresh (Power Query)
// =====================================================================================

export interface WeeklyShipment {
  id: string;
  date: number;
  carrier: string;
  lane: string;
  weight: number;
  cost: number;
}

export interface RefreshData {
  /** On the sheet at setup: the first three weeks. */
  base: WeeklyShipment[];
  /** Appended when step 2 begins. */
  week2: WeeklyShipment[];
  /** Appended when step 3 begins; includes the new carrier. */
  week3: WeeklyShipment[];
  newCarrier: string;
}

/** Carriers that start partway through the mission, never in CARRIERS. */
export const NEW_CARRIERS = ['Bayline Express', 'Keystone Freight', 'Pioneer Logistics'] as const;

const BASE_WEEKS = [serial(2026, 8, 31), serial(2026, 9, 7), serial(2026, 9, 14)];
const WEEK_2 = serial(2026, 9, 21);
const WEEK_3 = serial(2026, 9, 28);
const LABOR_DAY = serial(2026, 9, 7);

export const REFRESH_COLS: ColumnSpec[] = [
  { header: 'Ship ID' },
  { header: 'Ship date', format: FMT.date },
  { header: 'Carrier' },
  { header: 'Lane' },
  { header: 'Weight lb', format: FMT.int },
  { header: 'Cost', format: FMT.currency },
];

/** The query's output columns. */
export const SUMMARY_COLUMNS = ['Carrier', 'Total cost', 'Shipments'];

export const refreshGrid = (rows: WeeklyShipment[]): Grid => rows.map((r) => [r.id, r.date, r.carrier, r.lane, r.weight, r.cost]);

/** Every row in the Table once step `index` (0-based) has begun. */
export function rowsAtStep(d: RefreshData, index: number): WeeklyShipment[] {
  return [...d.base, ...(index >= 1 ? d.week2 : []), ...(index >= 2 ? d.week3 : [])];
}

/** Table.Group by Carrier with a Sum of Cost and a Count Rows: one row per carrier, in order of first appearance. */
export function carrierSummary(rows: WeeklyShipment[]): Grid {
  const groups = new Map<string, { cost: number; count: number }>();
  for (const r of rows) {
    const g = groups.get(r.carrier) ?? { cost: 0, count: 0 };
    groups.set(r.carrier, { cost: g.cost + r.cost, count: g.count + 1 });
  }
  return [...groups].map(([carrier, g]) => [carrier, g.cost, g.count]);
}

/** A weekday in the week starting `monday`, skipping Labor Day. */
function weekday(rng: Rng, monday: number): number {
  const day = monday + rng.int(0, 4);
  return day === LABOR_DAY ? day + 1 : day;
}

type DraftShipment = Omit<WeeklyShipment, 'id'>;

function draftWeek(rng: Rng, monday: number, carriers: readonly string[], lanes: Lane[], factor: Map<string, number>): DraftShipment[] {
  return carriers
    .map((carrier) => {
      const lane = rng.pick(lanes);
      const weight = shipmentWeight(rng);
      return { date: weekday(rng, monday), carrier, lane: lane.name, weight, cost: freightCost(rng, weight, lane, factor.get(carrier)!) };
    })
    .sort((a, b) => a.date - b.date);
}

const startNoteDates = (rows: WeeklyShipment[]) => `${shortDate(Math.min(...rows.map((r) => r.date)))} to ${shortDate(Math.max(...rows.map((r) => r.date)))}`;

const appendWeek = (index: number) => (d: RefreshData) => [tableWrite('Shipments', REFRESH_COLS, refreshGrid(rowsAtStep(d, index)))];
const summaryAnswer = (): AnswerArea => ({ kind: 'query', columns: SUMMARY_COLUMNS, order: 'any' });

export const weeklyRefresh = defineMission<RefreshData>({
  id: 'mission-weekly-refresh',
  title: 'Weekly carrier summary refresh',
  role: 'ops',
  summary: 'A carrier summary built once in Power Query, then refreshed as each week’s shipments arrive',
  minutes: 12,
  skills: ['pq-group'],
  brief: () => ({
    from: 'Nina Brooks, Transportation manager',
    subject: 'Carrier summary without the Monday rebuild',
    body: [
      'Hi,',
      '',
      'Every Monday I paste last week’s shipments under the `Shipments` Table and rebuild the carrier summary by hand. Set it up in Power Query once, so each week is a refresh instead of a rebuild.',
      '',
      'One row per carrier with its total cost and number of shipments is all I need. I’ll send the next two weeks as they come in, so we can see it hold up.',
      '',
      'Thanks,',
      'Nina',
    ].join('\n'),
  }),
  make: (rng) => {
    const lanes = makeLanes(rng, 6);
    const newCarrier = rng.pick(NEW_CARRIERS);
    const factor = new Map<string, number>([...CARRIERS, newCarrier].map((c) => [c, rng.float(0.85, 1.2, 2)]));
    const busy = (count: number) => Array.from({ length: count }, () => rng.pick(CARRIERS));
    // Each early week has every carrier at least once, so the first summary shows all four.
    const base = BASE_WEEKS.flatMap((monday) => draftWeek(rng, monday, [...CARRIERS, ...busy(rng.int(6, 8))], lanes, factor));
    const week2 = draftWeek(rng, WEEK_2, busy(rng.int(8, 11)), lanes, factor);
    const week3 = draftWeek(rng, WEEK_3, [...busy(rng.int(6, 8)), ...Array.from({ length: rng.int(2, 3) }, () => newCarrier)], lanes, factor);

    // Ship IDs keep rising across all three batches, the way the shipping system numbers them.
    let no = rng.int(60_100, 64_000);
    const number = (rows: DraftShipment[]): WeeklyShipment[] =>
      rows.map((r) => {
        no += rng.int(1, 3);
        return { id: `SH-${no}`, ...r };
      });
    return { base: number(base), week2: number(week2), week3: number(week3), newCarrier };
  },
  // Only the Table goes on the sheet: it grows downward as weeks are appended, and the query
  // output loads to its own sheet.
  blocks: (d) => [dataBlock('Shipments', 'A1', REFRESH_COLS, refreshGrid(d.base))],
  // Restoring inputs after a variant would undo the weekly appends, so there are none.
  inputs: () => [],
  steps: [
    // ---------- 1. Group By with two aggregations ----------
    {
      title: 'Build the carrier summary',
      task: () =>
        'Use Power Query to summarize the `Shipments` Table with one row per carrier and exactly three columns: `Carrier`, `Total cost` (the sum of Cost) and `Shipments` (the number of shipments). Load it to a new sheet, the default, so it never sits below the Table, which grows as each week’s rows are added. Row order doesn’t matter. Before you start, delete any queries left from earlier practice, and their sheets, in Data › Queries & Connections. Later steps use Data › Refresh All, which reruns every query, and one whose source Table is gone or has changed shows an error. {macPqNote}',
      hints: [
        'Choose {fromTable:Shipments}. The Power Query editor opens with the shipment rows.',
        'Choose Transform › Group By, then Advanced. Group by Carrier, and set the first aggregation to New column name Total cost, Operation Sum, Column Cost.',
        'Select Add aggregation for the second column: New column name Shipments, Operation Count Rows. Choose OK, then Home › Close & Load.',
      ],
      solution: () => '{fromTable:Shipments} · Transform › Group By › Advanced: group by Carrier · Total cost: Sum of Cost · Shipments: Count Rows · Home › Close & Load',
      answer: summaryAnswer,
      expected: (d) => carrierSummary(rowsAtStep(d, 0)),
      variants: [],
    },

    // ---------- 2. Refresh picks up new rows ----------
    {
      title: 'Refresh for next week',
      task: () =>
        'Your summary still shows the old totals. Refresh it with Data › Refresh All, then check. The query reads the whole Table each time it runs, so it picks up the new rows without any changes.',
      hints: [
        'Data › Refresh All reruns every query in the workbook. Right-clicking inside the summary table and choosing Refresh does the same for one query.',
        'If the totals haven’t changed, wait for the refresh to finish, then check again.',
        'An error that names another query comes from earlier practice: its source Table is gone or has changed. Delete that query in Data › Queries & Connections, or refresh only yours by right-clicking inside the summary table and choosing Refresh.',
      ],
      solution: () => 'Data › Refresh All, then Check.',
      answer: summaryAnswer,
      expected: (d) => carrierSummary(rowsAtStep(d, 1)),
      variants: [],
      onStart: appendWeek(1),
      startNote: (d) => `Next week’s shipments arrived: ${d.week2.length} rows dated ${startNoteDates(d.week2)} were added to the bottom of the \`Shipments\` Table.`,
    },

    // ---------- 3. Refresh picks up a new group ----------
    {
      title: 'Refresh with a new carrier',
      task: (d) =>
        `Refresh the summary again with Data › Refresh All, then check that ${d.newCarrier} has its own row with its total cost and number of shipments. Group By makes a row for every carrier it finds, so the query needs no changes.`,
      hints: [
        'Data › Refresh All, the same as last week.',
        'If the new carrier is missing after a refresh, open the query (double-click it in Data › Queries & Connections) and look in Applied Steps for a filter on Carrier.',
      ],
      solution: () => 'Data › Refresh All, then Check.',
      answer: summaryAnswer,
      expected: (d) => carrierSummary(rowsAtStep(d, 2)),
      variants: [],
      onStart: appendWeek(2),
      startNote: (d) => {
        const fresh = d.week3.filter((r) => r.carrier === d.newCarrier).length;
        return `Another week arrived: ${d.week3.length} rows dated ${startNoteDates(d.week3)} were added, ${fresh} of them from ${d.newCarrier}, a carrier that wasn’t in the Table before.`;
      },
    },
  ],
});

/** Missions added in v2, in the order they're offered. */
export const EXTRA_MISSIONS: Mission<any>[] = [apAging, carrierCostReview, weeklyRefresh];

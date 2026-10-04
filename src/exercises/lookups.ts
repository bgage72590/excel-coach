import { CARRIERS, GL_ACCOUNTS, ITEMS, VENDORS } from '../engine/data';
import { round, type Rng } from '../engine/rng';
import type { ColumnSpec, Grid } from '../engine/types';
import { FMT, cells, dataBlock, defineExercise, rangeWrite, tableWrite } from './common';

// ---------- shared item master ----------

interface ItemRow {
  sku: string;
  item: string;
  vendor: string;
  leadDays: number;
  cost: number;
}

const ITEM_COLS: ColumnSpec[] = [
  { header: 'SKU' },
  { header: 'Item' },
  { header: 'Vendor' },
  { header: 'Lead days', format: FMT.int },
  { header: 'Unit cost', format: FMT.currency },
];

function itemMaster(rng: Rng, count: number): ItemRow[] {
  const numbers = rng.sample(Array.from({ length: 900 }, (_, i) => 100 + i), count);
  return rng.sample(ITEMS, count).map((it, i) => ({
    sku: `SKU-${numbers[i]}`,
    item: it.item,
    vendor: rng.pick(VENDORS),
    leadDays: rng.int(3, 45),
    cost: it.cost,
  }));
}

const itemGrid = (rows: ItemRow[]) => rows.map((r) => [r.sku, r.item, r.vendor, r.leadDays, r.cost]);

interface PoLine {
  po: string;
  sku: string;
  qty: number;
}

function poLines(rng: Rng, skus: readonly string[], count: number): PoLine[] {
  return Array.from({ length: count }, (_, i) => ({ po: `PO-${5300 + i}`, sku: rng.pick(skus), qty: rng.int(5, 400) }));
}

const poGrid = (rows: PoLine[]) => rows.map((p) => [p.po, p.sku, p.qty]);

interface LeadData {
  items: ItemRow[];
  pos: PoLine[];
}

const reSort = {
  label: 'the Items table is re-sorted',
  explain: 'Match on the SKU itself so the row order of the Items table doesn’t matter.',
  apply: (d: LeadData, rng: Rng) => ({ ...d, items: rng.shuffle(d.items) }),
};

// ---------- Supplier lead time by SKU ----------

export const xlookupLeadTime = defineExercise<LeadData>({
  id: 'xlookup-lead-time',
  module: 'lookups',
  title: 'Supplier lead time by SKU',
  replaces: 'VLOOKUPs that break when someone inserts a column',
  minutes: 4,
  task: () => 'For each purchase order line, return the supplier’s lead time from the `Items` Table into `K2:K21`, matching on SKU.',
  concept: {
    summary:
      'XLOOKUP finds a value in one column and returns the matching value from another. Unlike VLOOKUP it doesn’t care where the columns sit, and it matches exactly by default.',
    syntax: '=XLOOKUP(lookup_value, lookup_column, return_column, [if_not_found])',
    example: '=XLOOKUP(I2, Items[SKU], Items[Lead days])',
  },
  hints: [
    'You’re looking up the SKU in I2.',
    'Look for it in Items[SKU] and return the value from Items[Lead days].',
    '=XLOOKUP(I2, Items[SKU], Items[Lead days]), then fill down.',
  ],
  solution: () => '=XLOOKUP(I2,Items[SKU],Items[Lead days])',
  make: (rng) => {
    const items = itemMaster(rng, 24);
    return { items, pos: poLines(rng, items.map((i) => i.sku), 20) };
  },
  layout: (d) => ({
    blocks: [
      dataBlock('Items', 'A1', ITEM_COLS, itemGrid(d.items)),
      cells('H1', [['PO', 'SKU', 'Qty', 'Lead days']], 'header'),
      cells('H2', poGrid(d.pos), 'input', [undefined, undefined, FMT.int]),
    ],
    answer: { kind: 'cells', range: 'K2:K21', format: FMT.int, consistency: 'all' },
  }),
  expected: (d) => d.pos.map((p) => [d.items.find((i) => i.sku === p.sku)!.leadDays]),
  inputs: (d) => [tableWrite('Items', ITEM_COLS, itemGrid(d.items)), rangeWrite('H2:J21', poGrid(d.pos))],
  variants: [
    reSort,
    { label: 'lead times change', apply: (d, rng) => ({ ...d, items: d.items.map((i) => ({ ...i, leadDays: rng.int(3, 45) })) }) },
    { label: 'the PO lines change', apply: (d, rng) => ({ ...d, pos: poLines(rng, d.items.map((i) => i.sku), 20) }) },
  ],
  rules: {
    require: [
      {
        pattern: /XLOOKUP\(|INDEX\(/i,
        label: 'Uses XLOOKUP (or INDEX and MATCH)',
        advice: 'Try =XLOOKUP(what to find, where to look, what to return).',
      },
    ],
  },
});

// ---------- Handle discontinued SKUs ----------

interface NotFoundData extends LeadData {
  discontinued: string[];
}

function poWithGaps(rng: Rng, items: ItemRow[]): { pos: PoLine[]; discontinued: string[] } {
  const known = new Set(items.map((i) => i.sku));
  const discontinued: string[] = [];
  while (discontinued.length < 4) {
    const sku = `SKU-${rng.int(1000, 1999)}`;
    if (!known.has(sku) && !discontinued.includes(sku)) discontinued.push(sku);
  }
  const pos = poLines(rng, items.map((i) => i.sku), 20);
  for (const [k, idx] of rng.sample(Array.from({ length: 20 }, (_, i) => i), 4).entries()) pos[idx] = { ...pos[idx], sku: discontinued[k] };
  return { pos, discontinued };
}

const costOrDiscontinued = (d: LeadData): Grid => d.pos.map((p) => [d.items.find((i) => i.sku === p.sku)?.cost ?? 'Discontinued']);

export const xlookupNotFound = defineExercise<NotFoundData>({
  id: 'xlookup-not-found',
  module: 'lookups',
  title: 'Handle discontinued SKUs',
  replaces: 'Columns full of #N/A, or IFERROR that hides real mistakes',
  minutes: 3,
  task: () =>
    'Return each PO line’s unit cost into `K2:K21`. Some SKUs were discontinued and aren’t in the `Items` Table; show the text `Discontinued` for those.',
  concept: {
    summary:
      'XLOOKUP’s fourth argument says what to return when nothing matches. It only catches “not found”, so a typo in the formula still shows up as an error instead of being hidden.',
    syntax: '=XLOOKUP(lookup_value, lookup_column, return_column, "Discontinued")',
  },
  hints: [
    'Start with a normal XLOOKUP on Items[SKU], returning Items[Unit cost].',
    'Add a fourth argument with the text to show when there’s no match.',
    '=XLOOKUP(I2, Items[SKU], Items[Unit cost], "Discontinued")',
  ],
  solution: () => '=XLOOKUP(I2,Items[SKU],Items[Unit cost],"Discontinued")',
  make: (rng) => {
    const items = itemMaster(rng, 24);
    return { items, ...poWithGaps(rng, items) };
  },
  layout: (d) => ({
    blocks: [
      dataBlock('Items', 'A1', ITEM_COLS, itemGrid(d.items)),
      cells('H1', [['PO', 'SKU', 'Qty', 'Unit cost']], 'header'),
      cells('H2', poGrid(d.pos), 'input', [undefined, undefined, FMT.int]),
    ],
    answer: { kind: 'cells', range: 'K2:K21', format: FMT.currency, consistency: 'all' },
  }),
  expected: costOrDiscontinued,
  inputs: (d) => [tableWrite('Items', ITEM_COLS, itemGrid(d.items)), rangeWrite('H2:J21', poGrid(d.pos))],
  variants: [
    { ...reSort, apply: (d, rng) => ({ ...d, items: rng.shuffle(d.items) }) },
    {
      label: 'costs change',
      apply: (d, rng) => ({ ...d, items: d.items.map((i) => ({ ...i, cost: round(i.cost * rng.float(0.7, 1.5, 3), 2) })) }),
    },
    { label: 'different SKUs are discontinued', apply: (d, rng) => ({ ...d, ...poWithGaps(rng, d.items) }) },
  ],
  rules: {
    require: [
      {
        pattern: /XLOOKUP\(/i,
        label: 'Uses XLOOKUP’s if-not-found argument',
        advice: 'XLOOKUP’s fourth argument sets what to show when there’s no match.',
      },
    ],
  },
});

// ---------- Freight rate by carrier and zone ----------

interface RateRow {
  carrier: string;
  zone: number;
  rate: number;
}

interface Ship {
  id: string;
  carrier: string;
  zone: number;
  kg: number;
}

const RATE_COLS: ColumnSpec[] = [{ header: 'Carrier' }, { header: 'Zone', format: FMT.plain }, { header: 'Rate per kg', format: FMT.rate }];

function rateTable(rng: Rng): RateRow[] {
  const rows: RateRow[] = [];
  for (const carrier of CARRIERS) {
    const base = rng.float(0.9, 1.6, 2);
    for (let zone = 1; zone <= 5; zone++) rows.push({ carrier, zone, rate: round(base + zone * rng.float(0.35, 0.6, 2), 2) });
  }
  return rng.shuffle(rows);
}

function ships(rng: Rng): Ship[] {
  return Array.from({ length: 20 }, (_, i) => ({ id: `SH-${20400 + i}`, carrier: rng.pick(CARRIERS), zone: rng.int(1, 5), kg: rng.int(20, 900) }));
}

const shipGrid = (rows: Ship[]) => rows.map((s) => [s.id, s.carrier, s.zone, s.kg]);
const rateGrid = (rows: RateRow[]) => rows.map((r) => [r.carrier, r.zone, r.rate]);

interface TwoKeyData {
  rates: RateRow[];
  ships: Ship[];
}

export const xlookupTwoKeys = defineExercise<TwoKeyData>({
  id: 'xlookup-two-keys',
  module: 'lookups',
  title: 'Freight rate by carrier and zone',
  replaces: 'A helper column that glues two keys together for a single lookup',
  minutes: 6,
  task: () =>
    'In `K2:K21`, return each shipment’s rate per kg from the `Rates` Table, matching both its carrier (`H`) and its zone (`I`).',
  concept: {
    summary:
      'To match on two columns, compare each one and multiply the results. (Rates[Carrier]=H2)*(Rates[Zone]=I2) is 1 only on the row where both are true, so XLOOKUP can look for that 1.',
    syntax: '=XLOOKUP(1, (column1=value1)*(column2=value2), return_column)',
    example: '=XLOOKUP(1, (Rates[Carrier]=H2)*(Rates[Zone]=I2), Rates[Rate per kg])',
  },
  hints: [
    'Each comparison like Rates[Carrier]=H2 gives a column of TRUE/FALSE.',
    'Multiplying two of them gives 1 only where both are TRUE. Look up 1 in that result.',
    '=XLOOKUP(1, (Rates[Carrier]=H2)*(Rates[Zone]=I2), Rates[Rate per kg])',
  ],
  solution: () => '=XLOOKUP(1,(Rates[Carrier]=H2)*(Rates[Zone]=I2),Rates[Rate per kg])',
  make: (rng) => ({ rates: rateTable(rng), ships: ships(rng) }),
  layout: (d) => ({
    blocks: [
      dataBlock('Rates', 'A1', RATE_COLS, rateGrid(d.rates)),
      cells('G1', [['Ship ID', 'Carrier', 'Zone', 'Weight kg', 'Rate per kg']], 'header'),
      cells('G2', shipGrid(d.ships), 'input', [undefined, undefined, FMT.plain, FMT.int]),
    ],
    answer: { kind: 'cells', range: 'K2:K21', format: FMT.rate, consistency: 'all' },
  }),
  expected: (d) => d.ships.map((s) => [d.rates.find((r) => r.carrier === s.carrier && r.zone === s.zone)!.rate]),
  inputs: (d) => [tableWrite('Rates', RATE_COLS, rateGrid(d.rates)), rangeWrite('G2:J21', shipGrid(d.ships))],
  variants: [
    {
      label: 'the Rates table is re-sorted',
      explain: 'Match on both conditions so row order doesn’t matter.',
      apply: (d, rng) => ({ ...d, rates: rng.shuffle(d.rates) }),
    },
    { label: 'rates change', apply: (d, rng) => ({ ...d, rates: d.rates.map((r) => ({ ...r, rate: round(r.rate * rng.float(0.8, 1.3, 3), 2) })) }) },
    { label: 'the shipments change', apply: (d, rng) => ({ ...d, ships: ships(rng) }) },
  ],
  rules: {
    forbidText: { values: [...CARRIERS], advice: 'Point to the carrier in column H instead.' },
    allowNumbers: [0, 1],
  },
});

// ---------- GL account name, looking left ----------

interface Account {
  name: string;
  type: string;
  no: number;
}

interface Entry {
  id: string;
  account: number;
  amount: number;
}

const ACCOUNT_COLS: ColumnSpec[] = [{ header: 'Account name' }, { header: 'Type' }, { header: 'Account no', format: FMT.plain }];

function entries(rng: Rng, accounts: Account[]): Entry[] {
  return Array.from({ length: 15 }, (_, i) => ({ id: `JE-${880 + i}`, account: rng.pick(accounts).no, amount: rng.float(-2500, 9000, 2) }));
}

interface LeftData {
  accounts: Account[];
  entries: Entry[];
}

const accountGrid = (rows: Account[]) => rows.map((a) => [a.name, a.type, a.no]);
const entryGrid = (rows: Entry[]) => rows.map((e) => [e.id, e.account, e.amount]);

export const lookupLeft = defineExercise<LeftData>({
  id: 'lookup-left',
  module: 'lookups',
  title: 'GL account name, looking left',
  replaces: 'Rearranging columns because VLOOKUP can only look right',
  minutes: 4,
  task: () =>
    'Journal entries in `E:G` have only an account number. In `H2:H16`, return each account’s name from the `Accounts` Table. The name sits to the left of the number.',
  concept: {
    summary:
      'VLOOKUP can only return columns to the right of the one it searches. XLOOKUP, and the older INDEX/MATCH pair, can return from any side.',
    syntax: '=XLOOKUP(F2, Accounts[Account no], Accounts[Account name])',
    example: 'The INDEX/MATCH version: =INDEX(Accounts[Account name], MATCH(F2, Accounts[Account no], 0))',
    tip: 'You’ll see INDEX/MATCH in older workbooks. It’s worth recognizing even if you write XLOOKUP.',
  },
  hints: [
    'Search Accounts[Account no] for the number in F2.',
    'Return from Accounts[Account name], even though it’s to the left.',
    '=XLOOKUP(F2, Accounts[Account no], Accounts[Account name])',
  ],
  solution: () => '=XLOOKUP(F2,Accounts[Account no],Accounts[Account name])',
  make: (rng) => {
    const accounts = rng.shuffle(GL_ACCOUNTS.map((a) => ({ name: a.name, type: a.type, no: a.no })));
    return { accounts, entries: entries(rng, accounts) };
  },
  layout: (d) => ({
    blocks: [
      dataBlock('Accounts', 'A1', ACCOUNT_COLS, accountGrid(d.accounts)),
      cells('E1', [['Entry', 'Account no', 'Amount', 'Account name']], 'header'),
      cells('E2', entryGrid(d.entries), 'input', [undefined, FMT.plain, FMT.currency]),
    ],
    answer: { kind: 'cells', range: 'H2:H16', consistency: 'all' },
  }),
  expected: (d) => d.entries.map((e) => [d.accounts.find((a) => a.no === e.account)!.name]),
  inputs: (d) => [tableWrite('Accounts', ACCOUNT_COLS, accountGrid(d.accounts)), rangeWrite('E2:G16', entryGrid(d.entries))],
  variants: [
    {
      label: 'the Accounts table is re-sorted',
      explain: 'Look up the account number itself so the row order doesn’t matter.',
      apply: (d, rng) => ({ ...d, accounts: rng.shuffle(d.accounts) }),
    },
    { label: 'the journal entries change', apply: (d, rng) => ({ ...d, entries: entries(rng, d.accounts) }) },
  ],
  rules: {
    require: [
      {
        pattern: /XLOOKUP\(|INDEX\(/i,
        label: 'Uses a lookup that can look left',
        advice: 'VLOOKUP can’t return a column to the left. Use XLOOKUP or INDEX/MATCH.',
      },
    ],
  },
});

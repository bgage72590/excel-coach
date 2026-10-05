import { CARRIERS, GL_ACCOUNTS, ITEMS, VENDORS } from '../engine/data';
import { round, type Rng } from '../engine/rng';
import type { ColumnSpec, Grid, SheetPointer } from '../engine/types';
import { FMT, cells, dataBlock, defineExercise, rangeWrite, tableWrite } from './common';
import { TABLE_TYPING_TIP, cellList, checkStep, fillStep, money, part, raw, rowsWhere, typeStep } from './guides';

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
  guide: (d) => {
    const po = d.pos[0];
    const at = d.items.findIndex((i) => i.sku === po.sku);
    const row = at + 2;
    const days = d.items[at]?.leadDays;
    return [
      {
        do: 'Meet the data. The blue block in columns `A` to `E` is a Table named **Items**, with one row per SKU. The purchase order lines sit in `H:J`, and the lead times go in `K2:K21`.',
        why: 'A Table has a name, and so does each of its columns. In a formula, `Items[SKU]` means the Table’s whole SKU column. Tap the buttons to see each part.',
        show: [
          { label: 'SKUs to look up', at: 'I2:I21', note: 'Each PO line names a SKU in column `I`. You’ll find each one in the Items Table.' },
          { label: 'SKU column', at: 'Items[SKU]', note: 'That’s `Items[SKU]` (column `A`): where XLOOKUP searches for each SKU.' },
          { label: 'Lead days column', at: 'Items[Lead days]', note: 'That’s `Items[Lead days]` (column `D`): the numbers to bring back.' },
        ],
      },
      {
        do: `See what \`K2\` should show. \`I2\` holds ${po.sku}, so find that SKU in the Items Table and read its Lead days.`,
        why: 'This is the lookup you’d do by eye: scan down the SKU column, stop at the match, read across. XLOOKUP does the same for every PO line.',
        show:
          at >= 0
            ? [{ label: `Find ${po.sku}`, at: `A${row}:E${row}`, note: `${po.sku} is on row ${row}. Its lead time is ${days} days, so \`K2\` should show ${days}.` }]
            : undefined,
      },
      typeStep({
        cell: 'K2',
        formula: [
          part('=XLOOKUP(', 'Finds a value in one column and returns the value on the same row from another column.'),
          part('I2', `What to find: the SKU in \`I2\` (${po.sku}). Point at the cell instead of typing the SKU, so each row looks up its own.`, 'I2'),
          raw(', '),
          part('Items[SKU]', 'Where to look for it: the SKU column of the Items Table.', 'Items[SKU]'),
          raw(', '),
          part('Items[Lead days]', 'What to bring back: the Lead days on the row where the SKU matched.', 'Items[Lead days]'),
          raw(')'),
        ],
        why: `${TABLE_TYPING_TIP} XLOOKUP matches exactly unless you tell it otherwise, so there’s no FALSE to add at the end the way VLOOKUP needs for an exact match.`,
      }),
      fillStep({
        from: 'K2',
        range: 'K2:K21',
        direction: 'down',
        why: 'Fill Down copies `K2`’s formula into the cells below. Excel moves `I2` along to `I3`, `I4` and on down, so each PO line looks up its own SKU. The Table columns stay put.',
      }),
      checkStep('The coach re-sorts the Items Table, changes lead times and swaps the PO lines behind the scenes, then puts everything back. XLOOKUP finds each SKU wherever its row ends up.'),
    ];
  },
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
  guide: (d) => {
    const itemOf = (sku: string) => d.items.find((i) => i.sku === sku);
    const foundAt = d.pos.findIndex((p) => itemOf(p.sku));
    const missingAt = d.pos.findIndex((p) => !itemOf(p.sku));
    const first = itemOf(d.pos[0].sku);
    const firstShows = first ? money(first.cost) : 'Discontinued';
    const pointers: SheetPointer[] = [];
    if (foundAt >= 0) {
      const sku = d.pos[foundAt].sku;
      const item = itemOf(sku)!;
      const itemRow = d.items.indexOf(item) + 2;
      pointers.push({
        label: 'A SKU that’s there',
        at: `A${itemRow}:E${itemRow}`,
        note: `\`I${foundAt + 2}\` holds ${sku}, which is on row ${itemRow} of the Items Table. Its Unit cost is ${money(item.cost)}, so \`K${foundAt + 2}\` should show ${money(item.cost)}.`,
      });
    }
    if (missingAt >= 0) {
      pointers.push({
        label: 'A discontinued SKU',
        at: `I${missingAt + 2}`,
        note: `${d.pos[missingAt].sku} isn’t anywhere in \`Items[SKU]\`. A plain XLOOKUP would show \`#N/A\` in \`K${missingAt + 2}\`; yours will show Discontinued.`,
      });
    }
    return [
      {
        do: 'Meet the data. The blue block in columns `A` to `E` is a Table named **Items**. The PO lines sit in `H:J`, and the unit costs go in `K2:K21`.',
        why: 'Four of the PO lines name a SKU that was discontinued, so it isn’t in the Items Table at all. A plain lookup shows the error `#N/A` (“not available”) for those; you’ll show the word Discontinued instead.',
        show: [
          { label: 'SKUs to look up', at: 'I2:I21', note: 'Each PO line names a SKU in column `I`.' },
          { label: 'SKU column', at: 'Items[SKU]', note: 'That’s `Items[SKU]` (column `A`): where XLOOKUP searches.' },
          { label: 'Unit cost column', at: 'Items[Unit cost]', note: 'That’s `Items[Unit cost]` (column `E`): the values to bring back.' },
        ],
      },
      {
        do: 'See the two kinds of answer. Most SKUs are in the Items Table, but a few aren’t.',
        why: `\`I2\` holds ${d.pos[0].sku}, so \`K2\` should show ${firstShows}.`,
        show: pointers.length ? pointers : undefined,
      },
      typeStep({
        cell: 'K2',
        formula: [
          part('=XLOOKUP(', 'Finds a value in one column and returns the value on the same row from another column.'),
          part('I2', `What to find: the SKU in \`I2\` (${d.pos[0].sku}).`, 'I2'),
          raw(', '),
          part('Items[SKU]', 'Where to look for it: the SKU column of the Items Table.', 'Items[SKU]'),
          raw(', '),
          part('Items[Unit cost]', 'What to bring back: the Unit cost on the row where the SKU matched.', 'Items[Unit cost]'),
          raw(', '),
          part('"Discontinued"', 'What to show when the SKU isn’t found anywhere. The pieces between the commas are called arguments, and this optional fourth one is named if_not_found. Text goes inside double quotes; spell it with a capital D.'),
          raw(')'),
        ],
        why: 'Why not wrap it in IFERROR? if_not_found only steps in when nothing matches. If the formula itself has a mistake, the error still shows instead of being hidden.',
      }),
      fillStep({
        from: 'K2',
        range: 'K2:K21',
        direction: 'down',
        why: '`I2` moves down a row at a time, so each PO line looks up its own SKU. The four discontinued ones show Discontinued instead of `#N/A`.',
      }),
      checkStep('Behind the scenes the coach re-sorts the Items Table, changes costs and discontinues different SKUs, then puts it all back. Your fourth argument catches whichever SKUs go missing.'),
    ];
  },
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
  guide: (d) => {
    const ship = d.ships[0];
    const keep = (r: RateRow) => r.carrier === ship.carrier && r.zone === ship.zone;
    const carrierRows = rowsWhere(d.rates, (r) => r.carrier === ship.carrier);
    const zoneRows = rowsWhere(d.rates, (r) => r.zone === ship.zone);
    const row = rowsWhere(d.rates, keep)[0];
    const rate = d.rates.find(keep)?.rate;
    return [
      {
        do: 'Meet the data. The blue block in columns `A` to `C` is a Table named **Rates**: one rate per kg for every carrier and zone. The shipments sit in `G:J`, and the rates go in `K2:K21`.',
        why: 'Each carrier appears five times, once per zone, and each zone appears once per carrier. So neither column alone pins down one row: the lookup has to match both.',
        show: [
          { label: 'Carrier column', at: 'Rates[Carrier]', note: 'That’s `Rates[Carrier]` (column `A`). Each carrier shows up on several rows.' },
          { label: 'Zone column', at: 'Rates[Zone]', note: 'That’s `Rates[Zone]` (column `B`). Each zone shows up on several rows too.' },
          { label: 'Rate column', at: 'Rates[Rate per kg]', note: 'That’s `Rates[Rate per kg]` (column `C`): the values to bring back.' },
        ],
      },
      {
        do: `See what \`K2\` should show: the rate for ${ship.carrier} (in \`H2\`) in zone ${ship.zone} (in \`I2\`).`,
        show: [
          { label: `${ship.carrier} rows`, at: cellList('A', carrierRows), note: `${ship.carrier} has ${carrierRows.length} rows, one per zone. The carrier alone can’t pick one.` },
          { label: `Zone ${ship.zone} rows`, at: cellList('B', zoneRows), note: `Zone ${ship.zone} has ${zoneRows.length} rows, one per carrier. The zone alone can’t pick one either.` },
          ...(row && rate !== undefined
            ? [{ label: 'Both match', at: `A${row}:C${row}`, note: `Row ${row} is the only row where both match. Its rate is ${money(rate)}, so \`K2\` should show ${money(rate)}.` }]
            : []),
        ],
      },
      {
        do: 'Peek at the trick before you use it. Click `E2`, an empty cell, and type this test, then press {enter}.',
        formula: [
          raw('=('),
          part('Rates[Carrier]=H2', `Test 1 asks every Rates row: is your Carrier ${ship.carrier}, the one in \`H2\`? Each row answers TRUE or FALSE.`, 'Rates[Carrier]'),
          raw(')'),
          part('*', 'Multiplies the two tests row by row. Excel counts TRUE as 1 and FALSE as 0, so a row gets 1 only when both tests are TRUE.'),
          raw('('),
          part('Rates[Zone]=I2', `Test 2 asks every Rates row: is your Zone ${ship.zone}, the one in \`I2\`?`, 'Rates[Zone]'),
          raw(')'),
        ],
        why: `The answers spill down \`E2:E21\`: one formula fills the cells below by itself, one result per Rates row, each beside the row it tested. Every row shows 0 except row ${row ?? 'with the match'}, which shows 1.`,
        show: row ? [{ label: 'Find the 1', at: `E${row}`, note: `Row ${row}: ${ship.carrier}, zone ${ship.zone}. That 1 is what XLOOKUP will look for.` }] : undefined,
      },
      {
        do: 'Click `E2` and press Delete.',
        why: 'That removes the peek and its whole spill. The real formula goes in `K2`.',
      },
      typeStep({
        cell: 'K2',
        formula: [
          part('=XLOOKUP(', 'Finds a value in one list and returns the value on the same row from another column.'),
          part('1', 'What to find: a 1, which marks the row where both tests are TRUE.'),
          raw(', ('),
          part('Rates[Carrier]=H2', `Where to look, test 1: does each row’s Carrier match \`H2\` (${ship.carrier})?`, 'Rates[Carrier]'),
          raw(')'),
          part('*', 'Multiplies the tests: 1 where both match, 0 everywhere else. That’s the column of 0s and 1s you saw in the peek.'),
          raw('('),
          part('Rates[Zone]=I2', `Where to look, test 2: does each row’s Zone match \`I2\` (zone ${ship.zone})?`, 'Rates[Zone]'),
          raw('), '),
          part('Rates[Rate per kg]', 'What to bring back: the rate on the row holding the 1.', 'Rates[Rate per kg]'),
          raw(')'),
        ],
        why: `Check the brackets: each test sits in its own pair. ${TABLE_TYPING_TIP}`,
      }),
      fillStep({
        from: 'K2',
        range: 'K2:K21',
        direction: 'down',
        why: '`H2` and `I2` move down a row at a time, so each shipment matches its own carrier and zone. The Rates columns stay put.',
      }),
      checkStep('The coach re-sorts the Rates Table and changes the rates behind the scenes, then puts them back. The 1 follows the matching row wherever it lands.'),
    ];
  },
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
  guide: (d) => {
    const no = d.entries[0].account;
    const at = d.accounts.findIndex((a) => a.no === no);
    const row = at + 2;
    const name = d.accounts[at]?.name;
    return [
      {
        do: 'Meet the data. The blue block in columns `A` to `C` is a Table named **Accounts**. The journal entries sit in `E:G`, and the account names go in `H2:H16`.',
        why: 'The number you search for is in column `C`, but the name you want is in column `A`, to its left. VLOOKUP can only bring back columns to the right of the one it searches. XLOOKUP takes the search column and the return column separately, so either side works.',
        show: [
          { label: 'Numbers to look up', at: 'F2:F16', note: 'Each journal entry has only an account number, in column `F`.' },
          { label: 'Account no column', at: 'Accounts[Account no]', note: 'That’s `Accounts[Account no]` (column `C`): where XLOOKUP searches.' },
          { label: 'Account name column', at: 'Accounts[Account name]', note: 'That’s `Accounts[Account name]` (column `A`): the names to bring back, to the left of the numbers.' },
        ],
      },
      {
        do: `See what \`H2\` should show. \`F2\` holds account ${no}, so find it in the Account no column and read the name to its left.`,
        show: at >= 0 ? [{ label: `Find account ${no}`, at: `A${row}:C${row}`, note: `Account ${no} is on row ${row}. Its name is ${name}, so \`H2\` should show ${name}.` }] : undefined,
      },
      typeStep({
        cell: 'H2',
        formula: [
          part('=XLOOKUP(', 'Finds a value in one column and returns the value on the same row from another column, on either side.'),
          part('F2', `What to find: the account number in \`F2\` (${no}).`, 'F2'),
          raw(', '),
          part('Accounts[Account no]', 'Where to look for it: the Account no column (`C`).', 'Accounts[Account no]'),
          raw(', '),
          part('Accounts[Account name]', 'What to bring back: the Account name on the matching row (`A`). It’s to the left, and XLOOKUP doesn’t mind.', 'Accounts[Account name]'),
          raw(')'),
        ],
        why: `${TABLE_TYPING_TIP} In older workbooks you’ll see the same lookup written as \`=INDEX(Accounts[Account name], MATCH(F2, Accounts[Account no], 0))\`. It gives the same answer.`,
      }),
      fillStep({
        from: 'H2',
        range: 'H2:H16',
        direction: 'down',
        why: '`F2` moves down to each entry’s account number in turn. The Table columns stay put.',
      }),
      checkStep('The coach re-sorts the Accounts Table and swaps in new journal entries behind the scenes, then puts everything back. XLOOKUP finds each number wherever its row ends up.'),
    ];
  },
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

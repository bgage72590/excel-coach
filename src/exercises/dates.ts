import { CARRIERS, VENDORS, eomonth, serial } from '../engine/data';
import type { Rng } from '../engine/rng';
import type { AnswerArea, ColumnSpec, Exercise } from '../engine/types';
import { FMT, cells, column, dataBlock, defineExercise, rangeWrite, tableWrite } from './common';
import { TABLE_TYPING_TIP, checkStep, fillStep, isoDate, part, raw, typeStep } from './guides';

// ---------- Excel business-day arithmetic ----------

/** Serial 0 (1899-12-30) is a Saturday, so serial mod 7 is 0 on Saturdays and 1 on Sundays. */
export function isWeekend(s: number): boolean {
  const r = ((Math.floor(s) % 7) + 7) % 7;
  return r === 0 || r === 1;
}

function isWorkday(s: number, holidays: ReadonlySet<number>): boolean {
  return !isWeekend(s) && !holidays.has(s);
}

/**
 * Excel's WORKDAY(start, days, holidays) with a Saturday/Sunday weekend. The start date is never
 * counted: the result is the `days`-th business day after it (before it when `days` is negative),
 * so a weekend or holiday start begins counting on the next business day. `days` is
 * truncated like Excel does; 0 returns the start date unchanged.
 */
export function workday(start: number, days: number, holidays: readonly number[] = []): number {
  const hol = new Set(holidays.map(Math.floor));
  const step = days < 0 ? -1 : 1;
  let left = Math.abs(Math.trunc(days));
  let d = Math.floor(start);
  while (left > 0) {
    d += step;
    if (isWorkday(d, hol)) left--;
  }
  return d;
}

/**
 * Excel's NETWORKDAYS(start, end, holidays): business days from start to end, counting both ends.
 * Holidays on a weekend, or listed twice, are only excluded once. Negative when end is before start.
 */
export function networkdays(start: number, end: number, holidays: readonly number[] = []): number {
  const a = Math.floor(start);
  const b = Math.floor(end);
  if (a > b) return -networkdays(b, a, holidays);
  const hol = new Set(holidays.map(Math.floor));
  let n = 0;
  for (let d = a; d <= b; d++) if (isWorkday(d, hol)) n++;
  return n;
}

// ---------- company holiday calendar ----------

export interface Holiday {
  name: string;
  date: number;
}

/** Observed dates, so every one falls on a weekday. */
export const HOLIDAYS: readonly Holiday[] = [
  { name: 'New Year’s Day', date: serial(2026, 1, 1) },
  { name: 'Martin Luther King Jr. Day', date: serial(2026, 1, 19) },
  { name: 'Presidents’ Day', date: serial(2026, 2, 16) },
  { name: 'Memorial Day', date: serial(2026, 5, 25) },
  { name: 'Juneteenth', date: serial(2026, 6, 19) },
  { name: 'Independence Day (observed)', date: serial(2026, 7, 3) },
  { name: 'Labor Day', date: serial(2026, 9, 7) },
  { name: 'Thanksgiving', date: serial(2026, 11, 26) },
  { name: 'Day after Thanksgiving', date: serial(2026, 11, 27) },
  { name: 'Christmas Eve', date: serial(2026, 12, 24) },
  { name: 'Christmas Day', date: serial(2026, 12, 25) },
  { name: 'New Year’s Day', date: serial(2027, 1, 1) },
  { name: 'Martin Luther King Jr. Day', date: serial(2027, 1, 18) },
  { name: 'Presidents’ Day', date: serial(2027, 2, 15) },
];

const HOLIDAY_COLS: ColumnSpec[] = [{ header: 'Holiday' }, { header: 'Date', format: FMT.date }];
const holidayGrid = (rows: readonly Holiday[]) => rows.map((h) => [h.name, h.date]);
const holidayDates = (rows: readonly Holiday[]) => rows.map((h) => h.date);

const ADDED_HOLIDAYS = ['Inventory count', 'Plant shutdown', 'Floating holiday'] as const;

/**
 * Appends a company holiday on a business day inside one of `spans` (inclusive), so at least one
 * answer moves when it is added. The row goes at the bottom, the way someone extends a Table.
 */
function addHoliday(holidays: readonly Holiday[], spans: [number, number][], rng: Rng): Holiday[] {
  const taken = new Set(holidayDates(holidays));
  const candidates = new Set<number>();
  for (const [from, to] of spans) for (let s = from; s <= to; s++) if (isWorkday(s, taken)) candidates.add(s);
  const pool = [...candidates].sort((a, b) => a - b);
  let date = pool.length ? rng.pick(pool) : Math.max(...taken) + 1;
  while (!isWorkday(date, taken)) date++;
  return [...holidays, { name: rng.pick(ADDED_HOLIDAYS), date }];
}

const holidayVariantLabel = 'a holiday is added to the Holidays Table';
const holidayVariantExplain = 'Refer to the whole column, Holidays[Date], so a holiday added to the Table counts right away.';

const CLOSED = new Set(holidayDates(HOLIDAYS));

/**
 * A date in [from, to] that's usually a business day. Holidays are always redrawn; weekend dates
 * mostly are, so a few weekend entries stay in the data. Redrawing keeps the dates spread out
 * instead of piling up on Mondays.
 */
function mostlyWeekday(rng: Rng, from: number, to: number, keepWeekend = 0.2): number {
  for (;;) {
    const d = rng.int(from, to);
    if (isWorkday(d, CLOSED)) return d;
    if (isWeekend(d) && rng.chance(keepWeekend)) return d;
  }
}

/** Re-rolls one field on every row, and guarantees at least one row actually changed. */
function reroll<T>(rows: readonly T[], next: (row: T) => T, same: (a: T, b: T) => boolean, bump: (row: T) => T): T[] {
  const out = rows.map(next);
  if (out.every((r, i) => same(r, rows[i]))) out[0] = bump(rows[0]);
  return out;
}

const holidaysRule = {
  pattern: /Holidays\[/i,
  label: 'Takes holidays from the Holidays Table column',
  advice: 'Pass the whole column, Holidays[Date], as the holidays argument, so holidays added to the Table later count too.',
};

// ---------- walkthrough helpers ----------

const DAY_NAMES = ['Saturday', 'Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday'] as const;

/** "Friday 2026-10-02": a date as the sheet shows it, with its weekday. */
const dayDate = (s: number) => `${DAY_NAMES[((Math.floor(s) % 7) + 7) % 7]} ${isoDate(s)}`;

const plural = (n: number, word: string) => `${n} ${word}${n === 1 ? '' : 's'}`;

/** How many days from `from` to `to`, both included, fall on a Saturday or Sunday. */
function weekendDays(from: number, to: number): number {
  let n = 0;
  for (let s = from; s <= to; s++) if (isWeekend(s)) n++;
  return n;
}

/** The names of the weekday holidays dated from `from` to `to`, both included. */
const holidaysIn = (holidays: readonly Holiday[], from: number, to: number) =>
  holidays.filter((h) => h.date >= from && h.date <= to && !isWeekend(h.date)).map((h) => h.name);

/** “Thanksgiving and Day after Thanksgiving”. */
const nameList = (names: readonly string[]) => (names.length > 1 ? `${names.slice(0, -1).join(', ')} and ${names[names.length - 1]}` : (names[0] ?? ''));

/** The Show button for the Holidays Table’s Date column, shared by the business-day walkthroughs. */
const holidayPointer = (holidays: readonly Holiday[]) => ({
  label: 'Holiday dates',
  at: 'Holidays[Date]',
  note: `That’s \`Holidays[Date]\`: ${holidays.length} company holidays to skip, as well as Saturdays and Sundays.`,
});

const HOLIDAY_PART = 'Which other days to skip: the whole Date column of the Holidays Table. Point at the column rather than its cells, so a holiday added to the Table later counts too.';

// ---------- Expected receipt date in business days ----------

interface PoLine {
  po: string;
  vendor: string;
  ordered: number;
  lead: number;
}

interface ReceiptData {
  holidays: Holiday[];
  lines: PoLine[];
}

const PO_COUNT = 24;
const PO_FROM = serial(2026, 10, 1);
const PO_TO = serial(2026, 12, 31);
const LEAD_MIN = 3;
const LEAD_MAX = 20;
/** Rows whose answer moves because of a holiday; enough that leaving out the holidays shows. */
export const MIN_HOLIDAY_LINES = 6;

export const receiptDate = (l: PoLine, holidays: readonly Holiday[]) => workday(l.ordered, l.lead, holidayDates(holidays));

/** Re-rolls rows that don't touch a holiday until at least MIN_HOLIDAY_LINES do. */
function ensureHolidayLines<T>(rows: T[], rng: Rng, crosses: (r: T) => boolean, fix: (r: T) => T): T[] {
  const out = rows.slice();
  for (let guard = 0; out.filter(crosses).length < MIN_HOLIDAY_LINES && guard < 1000; guard++) {
    const i = rng.int(0, out.length - 1);
    if (!crosses(out[i])) out[i] = fix(out[i]);
  }
  return out;
}

export const workdayReceipt = defineExercise<ReceiptData>({
  id: 'workday-receipt',
  module: 'dates',
  title: 'Expected receipt date in business days',
  replaces: 'Counting business days on a wall calendar for every PO',
  minutes: 5,
  task: (d) =>
    `Each PO line has an order date (\`F\`) and the supplier’s lead time in business days (\`G\`). In \`H2:H${d.lines.length + 1}\`, return the date each line should arrive: count the lead days forward from the order date, skipping weekends and every date in the \`Holidays\` Table; refer to those dates as \`Holidays[Date]\` so holidays added later count too. The order date itself doesn’t count. Write one formula in \`H2\` and fill it down.`,
  concept: {
    summary:
      'WORKDAY counts forward a number of business days from a start date, skipping Saturdays, Sundays and any holidays you list. The start date itself isn’t counted, so 1 business day after a Friday order is the following Monday.',
    syntax: '=WORKDAY(start_date, days, [holidays])',
    example: '=WORKDAY(F2, G2, Holidays[Date])',
    tip: 'If a site works Saturdays, WORKDAY.INTL takes a weekend code: WORKDAY.INTL(F2, G2, 11, Holidays[Date]) treats only Sunday as the weekend.',
  },
  hints: [
    'WORKDAY takes a start date, a number of business days, and an optional list of holidays.',
    'The start date is F2 and the number of days is G2. For the holidays, use the whole Date column of the Table: Holidays[Date].',
    '=WORKDAY(F2, G2, Holidays[Date]), then fill down.',
  ],
  solution: () => '=WORKDAY(F2,G2,Holidays[Date])',
  guide: (d) => {
    const end = d.lines.length + 1;
    const first = d.lines[0];
    const due = receiptDate(first, d.holidays);
    const skipped = holidaysIn(d.holidays, first.ordered + 1, due);
    // A later line whose date moves because of a holiday: the reason the third argument matters.
    const k = d.lines.findIndex((l, i) => i > 0 && receiptDate(l, d.holidays) !== workday(l.ordered, l.lead));
    const cross = k > 0 ? d.lines[k] : undefined;
    const crossRow = k + 2;
    const sample = cross ? { row: crossRow, line: cross } : { row: end, line: d.lines[d.lines.length - 1] };
    return [
      {
        do: 'Meet the data. Each PO line has an order date in column `F` and a lead time in column `G`. The blue block in columns `A` and `B` is a Table named **Holidays**.',
        why: 'Lead times are in business days: Monday to Friday, not counting company holidays. In a formula, `Holidays[Date]` means the whole Date column of that Table, and it grows when someone adds a holiday.',
        show: [
          { label: 'Order dates', at: `F2:F${end}`, note: 'The day each PO was placed. It doesn’t count as one of the lead days.' },
          { label: 'Lead days', at: `G2:G${end}`, note: 'How many business days each supplier needs.' },
          holidayPointer(d.holidays),
        ],
      },
      {
        do: `See what \`H2\` should show: ${plural(first.lead, 'business day')} after ${dayDate(first.ordered)}.`,
        why: `Start counting the day after the order and skip Saturdays and Sundays${skipped.length ? `, and ${nameList(skipped)} too` : ''}. Business day ${first.lead} is ${dayDate(due)}: that’s the date your formula in \`H2\` should show.`,
        show: [
          { label: 'Row 2’s order', at: 'F2:G2', note: `Ordered ${dayDate(first.ordered)}, with a lead time of ${plural(first.lead, 'business day')}.` },
          ...(cross
            ? [
                {
                  label: `Row ${crossRow} crosses a holiday`,
                  at: `F${crossRow}:G${crossRow}`,
                  note: `Ordered ${dayDate(cross.ordered)} with ${plural(cross.lead, 'lead day')}. Skipping weekends alone lands on ${dayDate(workday(cross.ordered, cross.lead))}. Skipping ${nameList(holidaysIn(d.holidays, cross.ordered + 1, receiptDate(cross, d.holidays)))} as well lands on ${dayDate(receiptDate(cross, d.holidays))}. That’s why the formula needs the holiday list.`,
                },
              ]
            : []),
        ],
      },
      typeStep({
        cell: 'H2',
        formula: [
          part('=WORKDAY(', 'Counts business days forward from a date, skipping Saturdays and Sundays.'),
          part('F2', `Where to start: this line’s order date (${isoDate(first.ordered)}). The start date itself isn’t counted.`, 'F2'),
          raw(', '),
          part('G2', `How many business days to count: this line’s lead time (${first.lead}).`, 'G2'),
          raw(', '),
          part('Holidays[Date]', HOLIDAY_PART, 'Holidays[Date]'),
          raw(')'),
        ],
        why: TABLE_TYPING_TIP,
      }),
      fillStep({
        from: 'H2',
        range: `H2:H${end}`,
        direction: 'down',
        why: `\`F2\` and \`G2\` move down to each line’s own order date and lead time. \`Holidays[Date]\` stays on the whole holiday column. Row ${sample.row} should show ${isoDate(receiptDate(sample.line, d.holidays))}.`,
      }),
      checkStep('The coach adds a holiday to the Holidays Table, changes the lead times and changes the order dates, to make sure every receipt date follows. Then it puts everything back.'),
    ];
  },
  make: (rng) => {
    const holidays = [...HOLIDAYS];
    const dates = holidayDates(holidays);
    const crosses = (l: Omit<PoLine, 'po'>) => workday(l.ordered, l.lead, dates) !== workday(l.ordered, l.lead);
    let raw = Array.from({ length: PO_COUNT }, () => ({
      vendor: rng.pick(VENDORS),
      ordered: mostlyWeekday(rng, PO_FROM, PO_TO),
      lead: rng.int(LEAD_MIN, LEAD_MAX),
    }));
    // Orders placed in the weeks before Thanksgiving and Christmas run through the holidays.
    raw = ensureHolidayLines(raw, rng, crosses, (l) => ({ ...l, ordered: mostlyWeekday(rng, serial(2026, 11, 16), serial(2026, 12, 22)) }));
    raw.sort((a, b) => a.ordered - b.ordered);
    return { holidays, lines: raw.map((r, i) => ({ po: `PO-${7400 + i}`, ...r })) };
  },
  layout: (d) => {
    const end = d.lines.length + 1;
    return {
      blocks: [
        dataBlock('Holidays', 'A1', HOLIDAY_COLS, holidayGrid(d.holidays)),
        cells('D1', [['PO', 'Vendor', 'Order date', 'Lead days', 'Expected receipt']], 'header'),
        cells('D2', d.lines.map((l) => [l.po, l.vendor]), 'input'),
        cells('F2', column(d.lines.map((l) => l.ordered)), 'input', FMT.date),
        cells('G2', column(d.lines.map((l) => l.lead)), 'input', FMT.int),
      ],
      answer: { kind: 'cells', range: `H2:H${end}`, format: FMT.date, consistency: 'all' },
    };
  },
  expected: (d) => d.lines.map((l) => [receiptDate(l, d.holidays)]),
  inputs: (d) => [
    tableWrite('Holidays', HOLIDAY_COLS, holidayGrid(d.holidays)),
    rangeWrite(`F2:G${d.lines.length + 1}`, d.lines.map((l) => [l.ordered, l.lead])),
  ],
  variants: [
    {
      label: holidayVariantLabel,
      explain: holidayVariantExplain,
      apply: (d, rng) => ({ ...d, holidays: addHoliday(d.holidays, d.lines.map((l) => [l.ordered + 1, receiptDate(l, d.holidays)]), rng) }),
    },
    {
      label: 'lead times change',
      explain: 'Point to the lead days in column G instead of typing a number.',
      apply: (d, rng) => ({
        ...d,
        lines: reroll(
          d.lines,
          (l) => ({ ...l, lead: rng.int(LEAD_MIN, LEAD_MAX) }),
          (a, b) => a.lead === b.lead,
          (l) => ({ ...l, lead: l.lead === LEAD_MAX ? LEAD_MIN : l.lead + 1 }),
        ),
      }),
    },
    {
      label: 'order dates change',
      explain: 'Point to the order dates in column F instead of typing them.',
      apply: (d, rng) => ({
        ...d,
        lines: reroll(
          d.lines,
          (l) => ({ ...l, ordered: mostlyWeekday(rng, PO_FROM, PO_TO) }),
          (a, b) => a.ordered === b.ordered,
          (l) => ({ ...l, ordered: l.ordered + 7 }),
        ),
      }),
    },
  ],
  rules: {
    require: [
      {
        pattern: /WORKDAY(\.INTL)?\(/i,
        label: 'Uses WORKDAY',
        advice: 'WORKDAY(start_date, days, holidays) counts business days for you.',
      },
      holidaysRule,
    ],
    allowNumbers: [1],
  },
});

// ---------- Business days a shipment actually took ----------

interface Shipment {
  id: string;
  carrier: string;
  ordered: number;
  delivered: number;
}

interface TransitData {
  holidays: Holiday[];
  ships: Shipment[];
}

const SHIP_COUNT = 24;
const SHIP_FROM = serial(2026, 5, 11);
const SHIP_TO = serial(2026, 9, 25);

export const businessDays = (s: Shipment, holidays: readonly Holiday[]) => networkdays(s.ordered, s.delivered, holidayDates(holidays));

/** Some carriers deliver on Saturdays; none deliver on Sundays or company holidays. */
function deliveredAfter(rng: Rng, ordered: number, from = 1, to = 12): number {
  let d = ordered + rng.int(from, to);
  const saturdayOk = rng.chance(0.4);
  while ((d % 7 === 0 && !saturdayOk) || d % 7 === 1 || CLOSED.has(d)) d++;
  return d;
}

/** A shipment ordered a few days before a holiday and delivered after it. */
function aroundHoliday(rng: Rng, s: Omit<Shipment, 'id'>): Omit<Shipment, 'id'> {
  const inWindow = HOLIDAYS.filter((h) => h.date > SHIP_FROM && h.date < SHIP_TO);
  const h = rng.pick(inWindow).date;
  let ordered = h - rng.int(1, 4);
  while (isWeekend(ordered)) ordered--;
  return { ...s, ordered, delivered: deliveredAfter(rng, h, 1, 6) };
}

export const networkdaysActual = defineExercise<TransitData>({
  id: 'networkdays-actual',
  module: 'dates',
  title: 'Business days a shipment actually took',
  replaces: 'Counting days on a calendar to check carrier transit times',
  minutes: 4,
  task: (d) =>
    `In \`H2:H${d.ships.length + 1}\`, count the business days each shipment took, from its order date (\`F\`) to its delivered date (\`G\`), skipping weekends and every date in the \`Holidays\` Table; refer to those dates as \`Holidays[Date]\`. Count the order date and the delivered date themselves when they fall on a business day. Write one formula in \`H2\` and fill it down.`,
  concept: {
    summary:
      'NETWORKDAYS counts the business days between two dates, skipping Saturdays, Sundays and any holidays you list. It counts both the start and the end date, so an order placed and delivered on the same Monday counts as 1.',
    syntax: '=NETWORKDAYS(start_date, end_date, [holidays])',
    example: '=NETWORKDAYS(F2, G2, Holidays[Date])',
    tip: 'G2-F2 counts calendar days, weekends included. NETWORKDAYS.INTL does the same job as NETWORKDAYS when the weekend isn’t Saturday and Sunday.',
  },
  hints: [
    'NETWORKDAYS takes a start date, an end date, and an optional list of holidays.',
    'Start at F2 and end at G2. For the holidays, use the whole Date column of the Table: Holidays[Date].',
    '=NETWORKDAYS(F2, G2, Holidays[Date]), then fill down.',
  ],
  solution: () => '=NETWORKDAYS(F2,G2,Holidays[Date])',
  guide: (d) => {
    const end = d.ships.length + 1;
    const first = d.ships[0];
    const days = businessDays(first, d.holidays);
    const span = first.delivered - first.ordered + 1;
    const weekend = weekendDays(first.ordered, first.delivered);
    const hols = holidaysIn(d.holidays, first.ordered, first.delivered);
    const weekendText = weekend === 0 ? 'none falls on a weekend' : weekend === 1 ? 'one falls on a weekend' : `${weekend} fall on a weekend`;
    const holText = hols.length === 0 ? '' : `, and ${nameList(hols)} ${hols.length === 1 ? 'is a holiday' : 'are holidays'}`;
    // A later shipment whose count drops because of a holiday: the reason the third argument matters.
    const k = d.ships.findIndex((s, i) => i > 0 && businessDays(s, d.holidays) !== networkdays(s.ordered, s.delivered));
    const cross = k > 0 ? d.ships[k] : undefined;
    const crossRow = k + 2;
    const crossHols = cross ? holidaysIn(d.holidays, cross.ordered, cross.delivered) : [];
    const sample = cross ? { row: crossRow, ship: cross } : { row: end, ship: d.ships[d.ships.length - 1] };
    return [
      {
        do: 'Meet the data. Each shipment has the date it was ordered in column `F` and the date it was delivered in column `G`. The blue block in columns `A` and `B` is a Table named **Holidays**.',
        why: 'Business days are Monday to Friday, not counting company holidays. In a formula, `Holidays[Date]` means the whole Date column of that Table, and it grows when someone adds a holiday.',
        show: [
          { label: 'Ordered', at: `F2:F${end}`, note: 'The first day of each shipment.' },
          { label: 'Delivered', at: `G2:G${end}`, note: 'The last day. Some carriers deliver on a Saturday, which doesn’t count as a business day.' },
          holidayPointer(d.holidays),
        ],
      },
      {
        do: `See what \`H2\` should show: the business days from ${dayDate(first.ordered)} to ${dayDate(first.delivered)}.`,
        why: `That’s ${plural(span, 'day')}, counting both the first and the last. Of those, ${weekendText}${holText}, so \`H2\` should show ${days}.`,
        show: [
          { label: 'Row 2’s dates', at: 'F2:G2', note: `Ordered ${dayDate(first.ordered)}, delivered ${dayDate(first.delivered)}.` },
          ...(cross
            ? [
                {
                  label: `Row ${crossRow} crosses a holiday`,
                  at: `F${crossRow}:G${crossRow}`,
                  note: `Ordered ${dayDate(cross.ordered)}, delivered ${dayDate(cross.delivered)}. Skipping weekends alone counts ${networkdays(cross.ordered, cross.delivered)}. ${nameList(crossHols)} ${crossHols.length === 1 ? 'is a holiday' : 'are holidays'}, so the answer is ${businessDays(cross, d.holidays)}. That’s why the formula needs the holiday list.`,
                },
              ]
            : []),
        ],
      },
      typeStep({
        cell: 'H2',
        formula: [
          part('=NETWORKDAYS(', 'Counts the business days from one date to another, both included, skipping Saturdays and Sundays.'),
          part('F2', `The first day: when this shipment was ordered (${isoDate(first.ordered)}).`, 'F2'),
          raw(', '),
          part('G2', `The last day: when it was delivered (${isoDate(first.delivered)}).`, 'G2'),
          raw(', '),
          part('Holidays[Date]', HOLIDAY_PART, 'Holidays[Date]'),
          raw(')'),
        ],
        why: TABLE_TYPING_TIP,
      }),
      fillStep({
        from: 'H2',
        range: `H2:H${end}`,
        direction: 'down',
        why: `\`F2\` and \`G2\` move down to each shipment’s own dates. \`Holidays[Date]\` stays on the whole holiday column. Row ${sample.row} should show ${businessDays(sample.ship, d.holidays)}.`,
      }),
      checkStep('The coach adds a holiday to the Holidays Table and changes the delivery dates, to make sure every count follows. Then it puts everything back.'),
    ];
  },
  make: (rng) => {
    const holidays = [...HOLIDAYS];
    const dates = holidayDates(holidays);
    const crosses = (s: Omit<Shipment, 'id'>) => networkdays(s.ordered, s.delivered, dates) !== networkdays(s.ordered, s.delivered);
    let raw = Array.from({ length: SHIP_COUNT }, (): Omit<Shipment, 'id'> => {
      const ordered = mostlyWeekday(rng, SHIP_FROM, SHIP_TO, 0.1);
      return { carrier: rng.pick(CARRIERS), ordered, delivered: deliveredAfter(rng, ordered) };
    });
    raw = ensureHolidayLines(raw, rng, crosses, (s) => aroundHoliday(rng, s));
    raw.sort((a, b) => a.ordered - b.ordered);
    return { holidays, ships: raw.map((r, i) => ({ id: `SH-${41200 + i}`, ...r })) };
  },
  layout: (d) => {
    const end = d.ships.length + 1;
    return {
      blocks: [
        dataBlock('Holidays', 'A1', HOLIDAY_COLS, holidayGrid(d.holidays)),
        cells('D1', [['Ship ID', 'Carrier', 'Ordered', 'Delivered', 'Business days']], 'header'),
        cells('D2', d.ships.map((s) => [s.id, s.carrier]), 'input'),
        cells('F2', d.ships.map((s) => [s.ordered, s.delivered]), 'input', FMT.date),
      ],
      answer: { kind: 'cells', range: `H2:H${end}`, format: FMT.int, consistency: 'all' },
    };
  },
  expected: (d) => d.ships.map((s) => [businessDays(s, d.holidays)]),
  inputs: (d) => [
    tableWrite('Holidays', HOLIDAY_COLS, holidayGrid(d.holidays)),
    rangeWrite(`F2:G${d.ships.length + 1}`, d.ships.map((s) => [s.ordered, s.delivered])),
  ],
  variants: [
    {
      label: holidayVariantLabel,
      explain: holidayVariantExplain,
      apply: (d, rng) => ({ ...d, holidays: addHoliday(d.holidays, d.ships.map((s) => [s.ordered, s.delivered]), rng) }),
    },
    {
      label: 'delivery dates change',
      explain: 'Point to the dates in columns F and G instead of typing them.',
      apply: (d, rng) => ({
        ...d,
        ships: reroll(
          d.ships,
          (s) => ({ ...s, delivered: deliveredAfter(rng, s.ordered) }),
          (a, b) => a.delivered === b.delivered,
          (s) => ({ ...s, delivered: deliveredAfter(rng, s.delivered) }),
        ),
      }),
    },
  ],
  rules: {
    require: [
      {
        pattern: /NETWORKDAYS(\.INTL)?\(/i,
        label: 'Uses NETWORKDAYS',
        advice: 'NETWORKDAYS(start_date, end_date, holidays) counts business days for you.',
      },
      holidaysRule,
    ],
    allowNumbers: [1],
  },
});

// ---------- Due dates on end-of-month terms ----------

interface Invoice {
  no: string;
  customer: string;
  issued: number;
  amount: number;
}

interface TermsData {
  invoices: Invoice[];
  terms: number;
  report: number;
}

const CUSTOMERS = [
  'Alder Hospitality',
  'Brightline Retail',
  'Cedar Grove Foods',
  'Dunmore Logistics',
  'Elm Street Clinics',
  'Fairway Hotels',
  'Glenwood Schools',
  'Harlow Markets',
] as const;

export const TERMS = [15, 30, 45, 60] as const;
/**
 * Report dates late enough that every June invoice is overdue even on 60-day terms, and early
 * enough that no September invoice is overdue even on 15-day terms, so each sheet has both.
 */
export const REPORT_DATES = [serial(2026, 9, 30), serial(2026, 10, 2), serial(2026, 10, 5), serial(2026, 10, 9), serial(2026, 10, 14)] as const;
const INVOICE_MONTHS = [6, 7, 8, 9];
const PER_MONTH = 6;

export const dueDate = (issued: number, terms: number) => eomonth(issued) + terms;
export const daysOverdue = (due: number, report: number) => Math.max(0, report - due);

function invoices(rng: Rng): Invoice[] {
  const rows: Omit<Invoice, 'no'>[] = [];
  for (const m of INVOICE_MONTHS) {
    const first = serial(2026, m, 1);
    const last = eomonth(first);
    for (let i = 0; i < PER_MONTH; i++) {
      let issued = rng.int(first, last);
      while (isWeekend(issued)) issued = rng.int(first, last);
      rows.push({ customer: rng.pick(CUSTOMERS), issued, amount: rng.float(250, 18_000, 2) });
    }
  }
  // One invoice goes out on the last day of its month, where EOMONTH returns the date itself.
  const k = rng.int(0, rows.length - 1);
  rows[k] = { ...rows[k], issued: eomonth(rows[k].issued) };
  rows.sort((a, b) => a.issued - b.issued);
  return rows.map((r, i) => ({ no: `INV-${24051 + i}`, ...r }));
}

const pickOther = <T>(rng: Rng, items: readonly T[], current: T): T => rng.pick(items.filter((x) => x !== current));

export const eomonthTerms = defineExercise<TermsData>({
  id: 'eomonth-terms',
  module: 'dates',
  title: 'Due dates on end-of-month terms',
  replaces: 'Working out end-of-month due dates by hand for the aging report',
  minutes: 6,
  task: (d) => {
    const end = d.invoices.length + 1;
    return `Invoices are due a set number of days after the end of the month they were issued; that number is in \`I1\`. In \`E2:E${end}\`, return each invoice’s due date. In \`F2:F${end}\`, return how many days overdue it is as of the report date in \`I2\`, or \`0\` if it isn’t overdue yet. Refer to \`I1\` and \`I2\` rather than typing their values. Write one formula at the top of each column and fill down.`;
  },
  concept: {
    summary:
      'EOMONTH returns the last day of a month a given number of months before or after a date; with 0 it’s the end of the date’s own month. Terms like “30 days after month end” are that date plus 30. Subtracting two dates gives the days between them, and MAX(0, …) turns a negative result into 0.',
    syntax: '=EOMONTH(start_date, months) + days',
    example: '=EOMONTH(C2, 0) + $I$1 is the due date. =MAX(0, $I$2 - E2) is the days overdue.',
    tip: 'EOMONTH(C2, -1) + 1 is the first day of C2’s month, handy for month-to-date windows.',
  },
  hints: [
    'Find the last day of each invoice’s month with EOMONTH, then add the terms days from I1. For days overdue, subtract the due date from the report date in I2.',
    'EOMONTH(C2, 0) is the end of C2’s own month. Lock the input cells so they stay put when you fill down: press {absKey} to turn I1 into $I$1. MAX(0, …) turns a negative difference into 0.',
    'Due date: =EOMONTH(C2, 0) + $I$1. Days overdue: =MAX(0, $I$2 - E2).',
  ],
  solution: () => 'E2: =EOMONTH(C2,0)+$I$1    F2: =MAX(0,$I$2-E2)',
  guide: (d) => {
    const end = d.invoices.length + 1;
    const first = d.invoices[0];
    const firstDue = dueDate(first.issued, d.terms);
    const late = d.report - firstDue;
    const last = d.invoices[d.invoices.length - 1];
    const lastGap = d.report - dueDate(last.issued, d.terms);
    // An invoice dated on the last day of its month, where EOMONTH hands back the date itself.
    const m = d.invoices.findIndex((inv) => inv.issued === eomonth(inv.issued));
    const monthEnd = m >= 0 ? d.invoices[m] : undefined;
    return [
      {
        do: 'Meet the layout. Each invoice’s date is in column `C`. `I1` holds the payment terms and `I2` the report date.',
        why: `Terms of ${d.terms} days after month end mean an invoice dated any day in a month is due ${d.terms} days after that month’s last day. Your formulas will read \`I1\` and \`I2\`, so changing either one updates every row.`,
        show: [
          { label: 'Invoice dates', at: `C2:C${end}` },
          { label: 'Terms', at: 'I1', note: `${d.terms}: the number of days after the end of the invoice’s month.` },
          { label: 'Report date', at: 'I2', note: `${isoDate(d.report)}: the day the aging report is run. Days overdue count up to this date.` },
        ],
      },
      {
        do: 'See what `E2` should show: the first invoice’s due date.',
        why: `\`C2\` is ${isoDate(first.issued)}, so its month ends on ${isoDate(eomonth(first.issued))}. Add the ${d.terms} days in \`I1\` and it’s due ${isoDate(firstDue)}. That’s the date for \`E2\`.`,
        show: [
          { label: 'First invoice date', at: 'C2' },
          ...(monthEnd
            ? [
                {
                  label: 'A month-end invoice',
                  at: `C${m + 2}`,
                  note: `Dated ${isoDate(monthEnd.issued)}, the last day of its month. EOMONTH hands back that same date, so it’s due exactly ${d.terms} days later, on ${isoDate(dueDate(monthEnd.issued, d.terms))}.`,
                },
              ]
            : []),
        ],
      },
      typeStep({
        cell: 'E2',
        formula: [
          part('=EOMONTH(', 'Finds the last day of a month.'),
          part('C2', `Which date’s month: this invoice’s date (${isoDate(first.issued)}).`, 'C2'),
          raw(', '),
          part('0', 'How many months to move first: 0 keeps the invoice’s own month. 1 would mean the month after.'),
          raw(')'),
          part(' + ', 'Plus: adding a number to a date moves it that many days later.'),
          part('$I$1', `The terms in \`I1\` (${d.terms} days). The \`$\` signs lock it, so every row still reads \`I1\` after you fill down.`, 'I1'),
        ],
        why: 'To add the `$` signs, click inside `I1` while typing and press {absKey} until it reads `$I$1`.',
      }),
      {
        do: 'See what `F2` should show: how many days overdue the first invoice is on the report date.',
        why: `${
          late > 0
            ? `The report date in \`I2\`, ${isoDate(d.report)}, is ${plural(late, 'day')} after the due date, so \`F2\` should show ${late}.`
            : `The report date in \`I2\`, ${isoDate(d.report)}, isn’t past the due date yet, so \`F2\` should show 0.`
        }${
          lastGap < 0
            ? ` The last invoice, on row ${end}, isn’t due until ${isoDate(dueDate(last.issued, d.terms))}. Report date minus due date is ${lastGap} there, a negative number, so that row should show 0 instead.`
            : ''
        }`,
        show: [
          { label: 'Report date', at: 'I2' },
          { label: 'First due date', at: 'E2' },
        ],
      },
      typeStep({
        cell: 'F2',
        formula: [
          part('=MAX(', 'Returns the larger of the two values inside.'),
          part('0', 'The floor: an invoice that isn’t due yet shows 0, not a negative number of days.'),
          raw(', '),
          part('$I$2', `The report date in \`I2\` (${isoDate(d.report)}), locked with \`$\` so every row reads it.`, 'I2'),
          part(' - ', 'Minus: one date minus another gives the number of days between them.'),
          part('E2', 'This invoice’s due date, the one you worked out in `E2`.', 'E2'),
          raw(')'),
        ],
        why: 'Lock `I2` the same way: click inside it while typing and press {absKey} until it reads `$I$2`.',
      }),
      {
        do: `Select \`E2:F${end}\` (click \`E2\`, then Shift-click \`F${end}\`) and press {fillDown}.`,
        why: 'Fill Down copies the top formula of each column into the cells below it. `C2` and `E2` move down a row at a time, while `$I$1` and `$I$2` stay on the inputs.',
        done: { kind: 'answer' },
      },
      checkStep('The coach changes the terms in `I1` and the report date in `I2` to make sure every due date and overdue count follows. Then it puts them back.'),
    ];
  },
  make: (rng) => ({ invoices: invoices(rng), terms: rng.pick(TERMS), report: rng.pick(REPORT_DATES) }),
  layout: (d) => {
    const end = d.invoices.length + 1;
    // AnswerArea 'cells' takes one number format; the per-column formats ride on alsoStyle.
    const dueCol: AnswerArea = { kind: 'cells', range: `E2:E${end}`, format: FMT.date, consistency: 'all' };
    const overdueCol: AnswerArea = { kind: 'cells', range: `F2:F${end}`, format: FMT.int, consistency: 'all' };
    return {
      blocks: [
        cells('A1', [['Invoice', 'Customer', 'Invoice date', 'Amount', 'Due date', 'Days overdue']], 'header'),
        cells('A2', d.invoices.map((i) => [i.no, i.customer]), 'input'),
        cells('C2', column(d.invoices.map((i) => i.issued)), 'input', FMT.date),
        cells('D2', column(d.invoices.map((i) => i.amount)), 'input', FMT.currency),
        cells('H1', [['Terms (days after month end)']], 'label'),
        cells('I1', [[d.terms]], 'input', FMT.plain),
        cells('H2', [['Report date']], 'label'),
        cells('I2', [[d.report]], 'input', FMT.date),
      ],
      answer: { kind: 'cells', range: `E2:F${end}`, consistency: 'columns' },
      alsoStyle: [dueCol, overdueCol],
    };
  },
  expected: (d) =>
    d.invoices.map((inv) => {
      const due = dueDate(inv.issued, d.terms);
      return [due, daysOverdue(due, d.report)];
    }),
  inputs: (d) => [rangeWrite('I1:I2', [[d.terms], [d.report]])],
  variants: [
    {
      label: 'the terms in I1 change',
      explain: 'Add $I$1 instead of typing the number of days, and lock it with $ so it holds when you fill down.',
      apply: (d, rng) => ({ ...d, terms: pickOther(rng, TERMS, d.terms) }),
    },
    {
      label: 'the report date in I2 changes',
      explain: 'Subtract from $I$2 instead of typing the date or using TODAY().',
      apply: (d, rng) => ({ ...d, report: pickOther(rng, REPORT_DATES, d.report) }),
    },
  ],
  rules: {
    // 1 keeps the classic DATE(YEAR(C2), MONTH(C2) + 1, 0) + $I$1 version from being flagged.
    allowNumbers: [0, 1],
  },
});

export const DATES: Exercise<any>[] = [workdayReceipt, networkdaysActual, eomonthTerms];

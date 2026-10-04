import { numberToCol } from '../engine/address';
import { DEPARTMENTS, VENDORS, eomonth, fromSerial, serial, sum } from '../engine/data';
import { round, type Rng } from '../engine/rng';
import type { ColumnSpec, Exercise, Grid, Inspection, PlantedBug, Rules, Variant } from '../engine/types';
import { FMT, cells, column, dataBlock, defineExercise, rangeWrite, tableWrite } from './common';

/** Whole dollars, for fixed costs and expense amounts. */
const USD0 = '$#,##0';

// ---------- shared: scenarios ----------

export const SCENARIOS = ['Base', 'Upside', 'Downside'] as const;
export type Scenario = (typeof SCENARIOS)[number];

/**
 * The assumptions a scenario flexes. All are rates, so the Live column shares one percentage
 * format and filling it down never copies a dollar format onto a percentage.
 */
export const ASSUMPTIONS = ['Monthly unit growth', 'Price change', 'COGS % of revenue', 'Opex change'] as const;

/** Each scenario's value for every assumption, in ASSUMPTIONS order. */
export type ScenarioTable = Record<Scenario, number[]>;

/** Data that carries a scenario switch: the selected case, the column order of B3:D3 and the values. */
export interface Switchable {
  scenario: Scenario;
  columns: Scenario[];
  values: ScenarioTable;
}

const permille = (rng: Rng, lo: number, hi: number) => rng.int(lo, hi) / 1000;

/**
 * Base, Upside and Downside values. Upside grows faster, prices higher and runs leaner; Downside
 * the reverse. The cases sit at least half a point apart on every row, so a Live column that
 * reads the wrong case always shows it.
 */
export function scenarioTable(rng: Rng): ScenarioTable {
  const base = [permille(rng, 5, 20), permille(rng, 10, 40), permille(rng, 540, 620), permille(rng, 20, 50)];
  const upside = [permille(rng, 5, 15), permille(rng, 10, 25), -permille(rng, 10, 30), -permille(rng, 10, 20)];
  const downside = [-permille(rng, 10, 25), -permille(rng, 15, 40), permille(rng, 15, 40), permille(rng, 15, 35)];
  return {
    Base: base,
    Upside: base.map((v, i) => round(v + upside[i], 3)),
    Downside: base.map((v, i) => round(v + downside[i], 3)),
  };
}

/** What the Live column returns: the selected scenario's value for each assumption. */
export const liveValues = (d: Switchable): number[] => d.values[d.scenario];

/** B3:D7 as the sheet shows it: the scenario names, then one row per assumption. */
export const scenarioGrid = (d: Switchable): Grid => [d.columns, ...ASSUMPTIONS.map((_, i) => d.columns.map((c) => d.values[c][i]))];

export function switchTo<D extends Switchable>(scenario: Scenario, explain: string): Variant<D> {
  return { label: `the scenario in B1 is set to ${scenario}`, explain, apply: (d) => ({ ...d, scenario }) };
}

export function swapUpsideDownside<D extends Switchable>(): Variant<D> {
  return {
    label: 'the Upside and Downside columns swap places, with Upside selected',
    explain:
      'Look the scenario up by name in the header row with XLOOKUP or INDEX/MATCH. CHOOSE, nested IFs or a fixed column tie each scenario to a position, so they return the wrong case once the columns move.',
    apply: (d) => ({ ...d, scenario: 'Upside', columns: d.columns.map((c) => (c === 'Upside' ? 'Downside' : c === 'Downside' ? 'Upside' : c)) }),
  };
}

export function newAssumptions<D extends Switchable>(explain: string): Variant<D> {
  return {
    label: 'the scenario assumptions change',
    explain,
    apply: (d, rng) => {
      let values = scenarioTable(rng);
      while (values[d.scenario].join() === d.values[d.scenario].join()) values = scenarioTable(rng);
      return { ...d, values };
    },
  };
}

export const scenarioDropdown = (label: string): Inspection => ({ kind: 'validationList', cell: 'B1', options: [...SCENARIOS], label });

/**
 * B1 back on the case the answer key assumes. A learner who tries the new dropdown and leaves it
 * on Upside sees every value miss; this names the cause instead of leaving them to guess.
 */
export const scenarioLeftOn = (d: Switchable): Inspection => ({
  kind: 'sheet',
  check: { kind: 'values', range: 'B1', expected: [[d.scenario]], describe: 'the scenario' },
  label: `B1 is left on ${d.scenario}`,
  advice: `Pick ${d.scenario} in the B1 dropdown, then check again.`,
});

/** The Live column's rules: a lookup on the header row, no scenario names typed into the formula. */
export const LIVE_RULES: Rules = {
  require: [
    {
      pattern: /XLOOKUP\(|INDEX\(.*MATCH\(/i,
      label: 'Looks the scenario up with XLOOKUP, or INDEX and MATCH',
      advice: 'Use XLOOKUP(scenario, header row, assumption row), or INDEX with MATCH on the header row.',
    },
  ],
  forbidText: { values: [...SCENARIOS], advice: 'Look up the scenario in B1 instead of typing its name.' },
  allowNumbers: [0, 1],
};

// ---------- shared: the 12-month forecast ----------

/** The forecast runs October 2026 through September 2027; September 2026 is the last actual month. */
export const FORECAST_MONTHS = Array.from({ length: 12 }, (_, i) => serial(2026, 10 + i, 1));
/** Month columns C through N. */
export const MONTH_COLS = FORECAST_MONTHS.map((_, i) => numberToCol(3 + i));
/** Number formats for a forecast header row: line label, Last actual, then the twelve months. */
export const MONTH_HEADER_FORMATS = [undefined, undefined, ...FORECAST_MONTHS.map(() => FMT.month)];

const MONTH_ABBR = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

/** "Mar 2027", as the month headers show it. */
export function monthLabel(s: number): string {
  const { year, month } = fromSerial(s);
  return `${MONTH_ABBR[month - 1]} ${year}`;
}

export const LINES = ['Units', 'Revenue', 'COGS', 'Gross margin', 'Opex', 'EBITDA'] as const;

/**
 * Twelve months of each line in LINES order, in the order Excel evaluates the filled-across
 * formulas: Units = last month × (1 + growth), Revenue = revenueOf(Units), COGS = Revenue × COGS %,
 * Gross margin = Revenue − COGS, Opex = a flat monthly figure, EBITDA = Gross margin − Opex.
 */
export function forecastLines(lastUnits: number, growth: number, revenueOf: (units: number) => number, cogsPct: number, opex: number): number[][] {
  const units: number[] = [];
  let prev = lastUnits;
  for (let m = 0; m < FORECAST_MONTHS.length; m++) {
    prev = prev * (1 + growth);
    units.push(prev);
  }
  const revenue = units.map(revenueOf);
  const cogs = revenue.map((r) => r * cogsPct);
  const gm = revenue.map((r, i) => r - cogs[i]);
  const opexRow = units.map(() => opex);
  const ebitda = gm.map((g, i) => g - opexRow[i]);
  return [units, revenue, cogs, gm, opexRow, ebitda];
}

export interface Drivers {
  /** Monthly unit growth, as a fraction. */
  growth: number;
  price: number;
  /** COGS as a fraction of revenue. */
  cogsPct: number;
  /** Fixed operating costs per month. */
  fixed: number;
}

const DRIVER_LABELS = ['Monthly unit growth', 'Price per unit', 'COGS % of revenue', 'Fixed costs per month'] as const;
const DRIVER_FORMATS = [FMT.pct, FMT.currency, FMT.pct, USD0] as const;

/** The forecast lines for a plain driver block: Revenue = Units × price, Opex = fixed costs. */
const driverForecast = (lastUnits: number, dr: Drivers) => forecastLines(lastUnits, dr.growth, (u) => u * dr.price, dr.cogsPct, dr.fixed);

const drawGrowth = (rng: Rng) => permille(rng, 5, 30);
/** $24.00 to $60.00 in 50-cent steps. */
export const drawPrice = (rng: Rng) => rng.int(48, 120) / 2;
/** 52% to 66% in half-point steps. */
const drawCogs = (rng: Rng) => rng.int(104, 132) / 200;
export const drawUnits = (rng: Rng) => rng.int(300, 900) * 10;

/** Fixed costs at 55–85% of month 1's gross margin, rounded to $500, so the forecast starts in profit. */
export const fixedFor = (rng: Rng, units: number, price: number, cogsPct: number) =>
  Math.round((units * price * (1 - cogsPct) * rng.float(0.55, 0.85, 3)) / 500) * 500;

function makeDrivers(rng: Rng, units: number): Drivers {
  const growth = drawGrowth(rng);
  const price = drawPrice(rng);
  const cogsPct = drawCogs(rng);
  return { growth, price, cogsPct, fixed: fixedFor(rng, units, price, cogsPct) };
}

/** Draws until the value differs from `current`, so a variant always changes something. */
export function redraw<T>(current: T, draw: () => T): T {
  let next = draw();
  while (next === current) next = draw();
  return next;
}

/** The last actual month's figures, typed in as the starting column of the forecast. */
export interface LastActual {
  units: number;
  revenue: number;
  cogs: number;
  opex: number;
}

/** Whole-dollar actuals at the price and roughly the COGS % the business ran at last month. */
export function makeActual(rng: Rng, units: number, price: number, cogsPct: number, opex: number): LastActual {
  const revenue = Math.round(units * price);
  return { units, revenue, cogs: Math.round(revenue * (cogsPct + rng.int(-4, 4) / 200)), opex };
}

/** The Last actual column in LINES order. */
export function actualColumn(a: LastActual): number[] {
  const gm = a.revenue - a.cogs;
  return [a.units, a.revenue, a.cogs, gm, a.opex, gm - a.opex];
}

/**
 * The correct formula for every line in one month column of a driver block laid out like the
 * forecast exercises (drivers in B2:B5, Units in row 8, EBITDA margin % in row 14). Units read
 * the column to the left, so month 1 reads the Last actual column like every later month.
 */
function monthFormulas(col: string, prev: string): string[] {
  return [`=${prev}8*(1+$B$2)`, `=${col}8*$B$3`, `=${col}9*$B$4`, `=${col}9-${col}10`, '=$B$5', `=${col}11-${col}12`, `=${col}13/${col}9`];
}

// ---------- Scenario switch with a dropdown ----------

export type ScenarioData = Switchable;

export const modelScenarioSwitch = defineExercise<ScenarioData>({
  id: 'model-scenario-switch',
  module: 'modeling',
  title: 'A Base, Upside and Downside switch',
  replaces: 'Copying the model once per case, or retyping assumptions before every meeting',
  minutes: 6,
  task: () =>
    'Turn `B1` into a scenario switch. First give it a dropdown: Data › Data Validation › Allow: List, with the scenario names in `B3:D3` as the source. Then fill the Live column, `E4:E7`, with each assumption’s value for the scenario in `B1`: one XLOOKUP (or INDEX and MATCH) in `E4` that finds `B1` in the header row, filled down. Leave `B1` on Base when you check. The coach switches it to Upside and Downside, changes the assumptions and swaps the Upside and Downside columns to prove the Live column follows.',
  concept: {
    summary:
      'A scenario switch keeps every case’s assumptions side by side and feeds one Live column into the model. The calculations read only the Live column, so changing one cell reruns the whole model for another case. A dropdown on the switch means nobody can pick a case that doesn’t exist.',
    syntax: '=XLOOKUP(scenario, header_row, assumption_row)',
    example: '=XLOOKUP($B$1, $B$3:$D$3, B4:D4), or with INDEX and MATCH: =INDEX(B4:D4, MATCH($B$1, $B$3:$D$3, 0))',
    tip: 'Find the scenario by name rather than picking a column with CHOOSE or nested IFs. Then adding a case, or moving the columns around, never breaks the model.',
  },
  hints: [
    'Select `B1`, then Data › Data Validation. On the Settings tab, set Allow to List and Source to `=$B$3:$D$3`, the scenario names in the header row. Keep In-cell dropdown ticked.',
    'In `E4`, look up `B1` in the header row `$B$3:$D$3` and return from the same row’s three values, `B4:D4`. Lock `B1` and the header row with $ ({absKey}) so they stay put as you fill down.',
    '`=XLOOKUP($B$1, $B$3:$D$3, B4:D4)` in `E4`, then fill down to `E7`.',
  ],
  solution: () => 'Data › Data Validation on B1: Allow List, Source =$B$3:$D$3. Live column: =XLOOKUP($B$1,$B$3:$D$3,B4:D4) in E4, filled down to E7.',
  make: (rng) => ({ scenario: 'Base', columns: [...SCENARIOS], values: scenarioTable(rng) }),
  layout: (d) => ({
    blocks: [
      cells('A1', [['Scenario']], 'label'),
      cells('B1', [[d.scenario]], 'input'),
      cells('A3', [['Assumption', ...d.columns, 'Live']], 'header'),
      cells('A4', column(ASSUMPTIONS), 'label'),
      cells('B4', scenarioGrid(d).slice(1), 'input', FMT.pct),
    ],
    answer: { kind: 'cells', range: 'E4:E7', format: FMT.pct, consistency: 'all' },
  }),
  expected: (d) => liveValues(d).map((v) => [v]),
  inputs: (d) => [rangeWrite('B1', [[d.scenario]]), rangeWrite('B3:D7', scenarioGrid(d))],
  variants: [
    switchTo('Upside', 'Look up the scenario in B1 so the Live column follows the switch.'),
    switchTo('Downside', 'Look up the scenario in B1 so the Live column follows the switch.'),
    newAssumptions('Return the values from the assumption rows instead of typing them.'),
    swapUpsideDownside(),
  ],
  rules: LIVE_RULES,
  inspections: (d) => [scenarioDropdown('B1 has a dropdown of Base, Upside and Downside'), scenarioLeftOn(d)],
});

// ---------- A driver-based 12-month forecast ----------

export interface ForecastData {
  drivers: Drivers;
  actual: LastActual;
}

/** Last month's price sat a few percent below the forecast price: the increase starts in October. */
const lastPriceFor = (rng: Rng, price: number) => round(price * rng.float(0.94, 0.98, 3), 2);

function forecastActual(rng: Rng, units: number, dr: Drivers): LastActual {
  return makeActual(rng, units, lastPriceFor(rng, dr.price), dr.cogsPct, Math.round((dr.fixed * rng.float(0.95, 1.02, 3)) / 100) * 100);
}

const driverColumn = (dr: Drivers): Grid => column([dr.growth, dr.price, dr.cogsPct, dr.fixed]);

export const modelDriverForecast = defineExercise<ForecastData>({
  id: 'model-driver-forecast',
  module: 'modeling',
  title: 'A driver-based 12-month forecast',
  replaces: 'A forecast where every month was typed or tweaked by hand',
  minutes: 8,
  task: () =>
    'Build the 12-month forecast in `C8:N13` from the drivers in `B2:B5`. Units grow each month by the rate in `B2`, starting from the Last actual units in `B8`. Revenue is Units × the price in `B3`. COGS is Revenue × the COGS % in `B4`. Gross margin is Revenue − COGS. Opex is the fixed costs in `B5`. EBITDA is Gross margin − Opex. Write each line once in column `C` and fill it right to `N`, so every month in a row uses the same formula. Don’t round.',
  concept: {
    summary:
      'A driver-based forecast keeps every assumption in one block of inputs and builds each line from them with formulas. Inputs, calculations and outputs stay apart: you change only the inputs, the calculations never hold a typed number, and the outputs update on their own.',
    syntax: 'Units: =B8*(1+$B$2), filled right, so each month grows from the one before it',
    example: 'Revenue =C8*$B$3 · COGS =C9*$B$4 · Gross margin =C9-C10 · Opex =$B$5 · EBITDA =C11-C12',
    tip: 'Giving the last actual month its own column means month 1 uses the same formula as month 12. One formula per row is quick to review and safe to fill.',
  },
  hints: [
    'Units in `C8` grow from the column to their left: `=B8*(1+$B$2)`. Lock the driver with $ ({absKey}) so it stays on `B2` as you fill right.',
    'Each other line reads its own month and a locked driver: Revenue `=C8*$B$3`, COGS `=C9*$B$4`, Gross margin `=C9-C10`, Opex `=$B$5`, EBITDA `=C11-C12`.',
    'Select `C8:C13` and drag the fill handle right to column `N`, or select `C8:N13` and use Home › Fill › Right. Every month in a row then holds the same formula.',
  ],
  solution: () => 'C8: =B8*(1+$B$2) · C9: =C8*$B$3 · C10: =C9*$B$4 · C11: =C9-C10 · C12: =$B$5 · C13: =C11-C12, then fill C8:C13 right to column N.',
  make: (rng) => {
    const units = drawUnits(rng);
    const drivers = makeDrivers(rng, units);
    return { drivers, actual: forecastActual(rng, units, drivers) };
  },
  layout: (d) => ({
    blocks: [
      cells('A1', [['Driver', 'Value']], 'header'),
      cells('A2', column(DRIVER_LABELS), 'label'),
      ...driverColumn(d.drivers).map((row, i) => cells(`B${2 + i}`, [row], 'input', DRIVER_FORMATS[i])),
      cells('A7', [['Line', 'Last actual', ...FORECAST_MONTHS]], 'header', MONTH_HEADER_FORMATS),
      cells('A8', column(LINES), 'label'),
      cells('B8', column(actualColumn(d.actual)), 'input', FMT.int),
    ],
    answer: { kind: 'cells', range: 'C8:N13', format: FMT.int, consistency: 'rows' },
  }),
  expected: (d) => driverForecast(d.actual.units, d.drivers),
  inputs: (d) => [rangeWrite('B2:B5', driverColumn(d.drivers)), rangeWrite('B8:B13', column(actualColumn(d.actual)))],
  variants: [
    {
      label: 'the growth rate changes',
      explain: 'Point every month at the growth driver with $B$2, so one change flows through the year.',
      apply: (d, rng) => ({ ...d, drivers: { ...d.drivers, growth: redraw(d.drivers.growth, () => drawGrowth(rng)) } }),
    },
    {
      label: 'the price changes',
      explain: 'Read the price from $B$3 instead of typing it.',
      apply: (d, rng) => ({ ...d, drivers: { ...d.drivers, price: redraw(d.drivers.price, () => drawPrice(rng)) } }),
    },
    {
      label: 'COGS % changes',
      explain: 'Read COGS % from $B$4 instead of typing it.',
      apply: (d, rng) => ({ ...d, drivers: { ...d.drivers, cogsPct: redraw(d.drivers.cogsPct, () => drawCogs(rng)) } }),
    },
    {
      label: 'the fixed costs change',
      explain: 'Point Opex in every month at $B$5.',
      apply: (d, rng) => ({ ...d, drivers: { ...d.drivers, fixed: d.drivers.fixed + rng.pick([-4000, -3000, -2000, 2000, 3000, 4000]) } }),
    },
    {
      label: 'the last actual units change',
      explain: 'Month 1 should grow from the Last actual units in B8, the same way every later month grows from the month before it.',
      apply: (d, rng) => {
        const units = redraw(d.actual.units, () => drawUnits(rng));
        return { ...d, actual: forecastActual(rng, units, d.drivers) };
      },
    },
  ],
  rules: { allowNumbers: [0, 1] },
});

// ---------- Check cells and a master check ----------

export interface ExpenseLine {
  date: number;
  dept: string;
  vendor: string;
  amount: number;
}

export interface CheckData {
  /** The departments listed in the summary, F2:F7. */
  departments: string[];
  lines: ExpenseLine[];
  /** The ledger total in H10. */
  ledger: number;
}

const EXPENSE_COLS: ColumnSpec[] = [
  { header: 'Date', format: FMT.date },
  { header: 'Department' },
  { header: 'Vendor' },
  { header: 'Amount', format: USD0 },
];

const CHECK_MONTH = serial(2026, 9, 1);
const SUMMARY_DEPTS = 6;
const CHECK_LABELS = ['Department total minus ledger', 'Table rows minus summary lines', 'All checks'] as const;

const expenseGrid = (lines: ExpenseLine[]): Grid => lines.map((l) => [l.date, l.dept, l.vendor, l.amount]);

/** Whole-dollar amounts: the sums are exact, so a tie is exactly 0 in Excel and in JS. */
function expenseLine(rng: Rng, depts: readonly string[]): ExpenseLine {
  return { date: rng.int(CHECK_MONTH, eomonth(CHECK_MONTH)), dept: rng.pick(depts), vendor: rng.pick(VENDORS), amount: rng.int(120, 9800) };
}

const byDate = (a: ExpenseLine, b: ExpenseLine) => a.date - b.date;

/** The ledger agrees with the Table: the state every "still ties" variant returns to. */
const tied = (d: CheckData): CheckData => ({ ...d, ledger: sum(d.lines.map((l) => l.amount)) });

/**
 * What the three check cells show: the department total minus the ledger, the Table's row count
 * minus the summary's line count, and the master check.
 */
export function checkResults(d: CheckData): [number, number, string] {
  const listed = d.lines.filter((l) => d.departments.includes(l.dept));
  const amount = sum(listed.map((l) => l.amount)) - d.ledger;
  const count = d.lines.length - listed.length;
  return [amount, count, amount === 0 && count === 0 ? 'OK' : 'Check'];
}

/** One line recoded to a department the summary doesn't list, so SUMIFS and COUNTIFS miss it. */
function strayLine(d: CheckData, rng: Rng): CheckData {
  const stray = rng.pick(DEPARTMENTS.filter((x) => !d.departments.includes(x)));
  const k = rng.int(0, d.lines.length - 1);
  return { ...d, lines: d.lines.map((l, i) => (i === k ? { ...l, dept: stray } : l)) };
}

const summaryFormulas = (rows: number): Grid =>
  Array.from({ length: rows }, (_, i) => [`=COUNTIFS(Expenses[Department],F${2 + i})`, `=SUMIFS(Expenses[Amount],Expenses[Department],F${2 + i})`]);

export const modelCheckCells = defineExercise<CheckData>({
  id: 'model-check-cells',
  module: 'modeling',
  title: 'Check cells and a master check',
  replaces: 'Re-adding the summary on a calculator before every review',
  minutes: 5,
  task: () =>
    'The department summary in `F1:H8` is built from the `Expenses` Table, and `H10` holds the ledger total from the GL. Add three check cells so anyone can see at a glance that the summary ties. In `G13`, show the department total in `H8` minus the ledger total in `H10`. In `G14`, show the number of rows in the `Expenses` Table minus the summary’s line count in `G8`. Both are 0 when everything ties. In `G15`, return `OK` when both checks are 0 and `Check` otherwise. The coach then changes the data, and plants mismatches, to prove your checks catch them.',
  concept: {
    summary:
      'A check cell works out the same figure two independent ways and shows the difference, so it reads 0 while the numbers tie. When a line is miscoded, a total is overwritten or a row goes missing, the check moves off 0 the moment it happens instead of in the review meeting. A master check rolls every check into one OK you can see from the top of the sheet.',
    syntax: '=IF(AND(check1=0, check2=0), "OK", "Check")',
    example: '=H8-H10 ties the summary to the ledger. =ROWS(Expenses)-G8 proves no line fell outside the summary.',
    tip: 'With amounts in cents, compare a rounded difference, such as ROUND(G13, 2)=0. Tiny floating-point leftovers can make an exact = test fail.',
  },
  hints: [
    'Each check subtracts one figure from another that should match. In `G13`: `=H8-H10`.',
    'ROWS counts a Table’s data rows however many there are, so `=ROWS(Expenses)-G8` in `G14` catches any line the summary leaves out.',
    'The master check passes only when both differences are 0: `=IF(AND(G13=0, G14=0), "OK", "Check")` in `G15`.',
  ],
  solution: () => 'G13: =H8-H10 · G14: =ROWS(Expenses)-G8 · G15: =IF(AND(G13=0,G14=0),"OK","Check")',
  make: (rng) => {
    const departments = rng.sample(DEPARTMENTS, SUMMARY_DEPTS);
    // Two lines per department first, so every summary row has something to count.
    const lines = [
      ...departments.flatMap((dept) => [0, 1].map(() => ({ ...expenseLine(rng, departments), dept }))),
      ...Array.from({ length: 18 }, () => expenseLine(rng, departments)),
    ].sort(byDate);
    return tied({ departments, lines, ledger: 0 });
  },
  // The Table sits in A:D; column E stays empty so nothing written in F:H joins it.
  layout: (d) => ({
    blocks: [
      dataBlock('Expenses', 'A1', EXPENSE_COLS, expenseGrid(d.lines)),
      cells('F1', [['Department', 'Lines', 'Amount']], 'header'),
      cells('F2', column(d.departments), 'label'),
      cells('G2', summaryFormulas(d.departments.length), 'formula', [FMT.int, USD0]),
      cells('F8', [['Total']], 'label'),
      cells('G8', [['=SUM(G2:G7)', '=SUM(H2:H7)']], 'formula', [FMT.int, USD0]),
      cells('F10', [['Ledger total']], 'label'),
      cells('H10', [[d.ledger]], 'input', USD0),
      cells('F12', [['Check', 'Result']], 'header'),
      cells('F13', column(CHECK_LABELS), 'label'),
    ],
    answer: { kind: 'cells', range: 'G13:G15', format: FMT.int, consistency: 'none' },
  }),
  expected: (d) => checkResults(d).map((v) => [v]),
  inputs: (d) => [tableWrite('Expenses', EXPENSE_COLS, expenseGrid(d.lines)), rangeWrite('H10', [[d.ledger]])],
  variants: [
    {
      label: 'the amounts change and the ledger still ties',
      explain: 'Subtract the ledger total in H10 instead of a typed number.',
      apply: (d, rng) => tied({ ...d, lines: d.lines.map((l) => ({ ...l, amount: rng.int(120, 9800) })) }),
    },
    {
      label: 'four lines are added and the ledger still ties',
      explain: 'Count the rows with ROWS(Expenses) so new lines are included.',
      apply: (d, rng) => tied({ ...d, lines: [...d.lines, ...Array.from({ length: 4 }, () => expenseLine(rng, d.departments))].sort(byDate) }),
    },
    {
      label: 'the ledger total stops matching',
      explain: 'G13 should show the gap, and G15 should switch to Check as soon as either difference isn’t 0.',
      apply: (d, rng) => ({ ...d, ledger: d.ledger + rng.pick([-1, 1]) * rng.int(15, 240) * 10 }),
    },
    {
      label: 'a line is coded to a department the summary doesn’t list',
      explain: 'Count every row of the Expenses Table, not only the listed departments, so a stray line shows up. G15 should then say Check.',
      apply: strayLine,
    },
    {
      // The only run where the amounts tie but the count doesn't, so a master check that reads
      // G13 alone, or compares H8 with H10, shows OK here.
      label: 'a line drops out of the summary and the ledger total drops with it',
      explain: 'G15 must read both checks. G13 ties here, but a line sits outside the summary, so G14 isn’t 0 and G15 should say Check.',
      apply: (d, rng) => {
        const moved = strayLine(d, rng);
        return { ...moved, ledger: sum(moved.lines.filter((l) => moved.departments.includes(l.dept)).map((l) => l.amount)) };
      },
    },
  ],
  // Zero and the usual rounding tolerances only: a typed total or row count is flagged.
  rules: { allowNumbers: [0, 0.005, 0.01, 0.5, 1, 2] },
});

// ---------- Trace a broken model back to its inputs ----------

export interface AuditData {
  forecast: Drivers;
  /** Last year's drivers, shown for comparison. Nothing in the model should read them. */
  lastYear: Drivers;
  actual: LastActual;
  /** Month index (0 = October) of the gross margin that was pasted over with its value. */
  pasted: number;
  /** Month index of the COGS formula that reads last year's COGS % instead of the forecast's. */
  wrongLink: number;
  /** The number pasted over that gross margin: its correct value on the starting data. */
  pastedValue: number;
}

const AUDIT_LINES = [...LINES, 'EBITDA margin %'] as const;
/** The model's calculations and outputs, Units through EBITDA margin %, October through September. */
export const AUDIT_RANGE = 'C8:N14';

/** The model as it should be, row by row in AUDIT_LINES order. */
export function auditGrid(d: Pick<AuditData, 'forecast' | 'actual'>): number[][] {
  const lines = driverForecast(d.actual.units, d.forecast);
  return [...lines, lines[5].map((e, i) => e / lines[1][i])];
}

/** The two planted breaks and every cell each one throws off. Their cells never overlap. */
export function auditBugs(d: AuditData): PlantedBug[] {
  const p = MONTH_COLS[d.pasted];
  const w = MONTH_COLS[d.wrongLink];
  return [
    {
      id: 'pasted-value',
      label: `Gross margin for ${monthLabel(FORECAST_MONTHS[d.pasted])} was a pasted number, so it ignored the drivers`,
      cells: [`${p}11`, `${p}13`, `${p}14`],
    },
    {
      id: 'wrong-input',
      label: `COGS for ${monthLabel(FORECAST_MONTHS[d.wrongLink])} read last year’s COGS % in C4 instead of the forecast’s in B4`,
      cells: [`${w}10`, `${w}11`, `${w}13`, `${w}14`],
    },
  ];
}

const auditActual = (rng: Rng, units: number, ly: Drivers) => makeActual(rng, units, ly.price, ly.cogsPct, ly.fixed);

/**
 * Last year grew a little slower at a lower price, and its COGS % sits 1 to 3 points away from the
 * forecast's, so a formula that reads it always shows a different number.
 */
function lastYearFor(rng: Rng, f: Drivers): Drivers {
  return {
    growth: round(f.growth - permille(rng, 2, 8), 3),
    price: round(f.price * rng.float(0.93, 0.97, 3), 2),
    cogsPct: round(f.cogsPct + (rng.pick([-1, 1]) * rng.int(2, 6)) / 200, 3),
    fixed: Math.round((f.fixed * rng.float(0.9, 0.97, 3)) / 500) * 500,
  };
}

const driverGrid = (d: AuditData): Grid => (['growth', 'price', 'cogsPct', 'fixed'] as const).map((k) => [d.forecast[k], d.lastYear[k]]);
const auditActualColumn = (d: AuditData): Grid => {
  const col = actualColumn(d.actual);
  return column([...col, round(col[5] / col[1], 4)]);
};

/**
 * The model as handed over: every cell in C8:N14 is a formula except the pasted gross margin,
 * and one COGS formula reads $C$4.
 */
function auditModel(d: AuditData) {
  const formulas = MONTH_COLS.map((col, i) => {
    const f = monthFormulas(col, i === 0 ? 'B' : MONTH_COLS[i - 1]);
    if (i === d.wrongLink) f[2] = `=${col}9*$C$4`;
    return f;
  });
  const row = (line: number, from = 0, to = MONTH_COLS.length) => [formulas.slice(from, to).map((f) => f[line])];
  const p = d.pasted;
  return [
    cells('C8', row(0), 'formula', FMT.int),
    cells('C9', row(1), 'formula', FMT.int),
    cells('C10', row(2), 'formula', FMT.int),
    cells('C11', row(3, 0, p), 'formula', FMT.int),
    cells(`${MONTH_COLS[p]}11`, [[d.pastedValue]], 'input', FMT.int),
    cells(`${MONTH_COLS[p + 1]}11`, row(3, p + 1), 'formula', FMT.int),
    cells('C12', row(4), 'formula', FMT.int),
    cells('C13', row(5), 'formula', FMT.int),
    cells('C14', row(6), 'formula', FMT.pct),
  ];
}

const AUDIT_EXPLAIN = {
  drivers: 'A forecast cell that doesn’t move with the drivers was typed over. Trace Precedents from an output: a calculation with no arrows coming in has no formula.',
  cogs: 'Every COGS formula should read the forecast COGS % in B4.',
  lastYear:
    'The forecast shouldn’t move when the Last year column changes. Select each Last year input and choose Formulas › Trace Dependents to see which formula reads it.',
};

export const modelAuditTrace = defineExercise<AuditData>({
  id: 'model-audit-trace',
  module: 'modeling',
  title: 'Trace a broken forecast back to its inputs',
  replaces: 'Clicking through cells one by one to work out why a total looks off',
  minutes: 7,
  task: () =>
    'This forecast came back from review with two breaks in its chain. One calculation was pasted over with a number, so it no longer follows the drivers, and one formula reads the wrong input. Start from the outputs, EBITDA and EBITDA margin % in rows `13` and `14`: select a cell and use Formulas › Trace Precedents, clicking again to step back toward the drivers. On the drivers, Formulas › Trace Dependents shows which formulas read each one, and nothing in the forecast should read the Last year column in `C2:C5`. Fix both cells so every month follows the Forecast drivers in `B2:B5`, and leave the rest of `C8:N14` as it is.',
  concept: {
    summary:
      'Formulas › Trace Precedents draws arrows from the cells a formula reads, and Trace Dependents draws arrows to the cells that read the selected one. Each click goes one level further. Walking back from an output shows where a chain breaks: a calculation with no arrows coming in was typed over, and an arrow from the wrong input is a broken link.',
    syntax: 'Formulas › Trace Precedents · Formulas › Trace Dependents · Formulas › Remove Arrows',
    example: 'Trace Precedents on an EBITDA cell points at Gross margin and Opex in the same month. On Gross margin it points at Revenue and COGS.',
    tip: 'Formulas › Show Formulas swaps every result for its formula, so a typed number or an odd reference stands out in a row of matching formulas.',
  },
  hints: [
    'Select an EBITDA cell in row `13` and choose Formulas › Trace Precedents. Click it again to go one level further back. Formulas › Remove Arrows clears them.',
    'A calculation that shows no arrows coming in was pasted over. Formulas › Show Formulas makes a typed number stand out in a row of formulas, and so does changing a driver and watching for a month that doesn’t move. Change the driver back before you check.',
    'Select each Last year input in `C2:C5` and choose Formulas › Trace Dependents. An arrow into the forecast marks the formula that reads the wrong input.',
    'Rewrite each broken cell with the same formula as its neighbours in that row, reading the Forecast column: Gross margin is Revenue − COGS, and COGS is Revenue × `$B$4`.',
  ],
  solution: (d) => {
    const p = MONTH_COLS[d.pasted];
    const w = MONTH_COLS[d.wrongLink];
    return `Gross margin in ${p}11 held a typed number: =${p}9-${p}10. COGS in ${w}10 read last year’s COGS %: =${w}9*$B$4.`;
  },
  make: (rng) => {
    const units = drawUnits(rng);
    const forecast = makeDrivers(rng, units);
    const lastYear = lastYearFor(rng, forecast);
    const pasted = rng.int(2, 9);
    const wrongLink = rng.pick(Array.from({ length: 10 }, (_, i) => i + 1).filter((i) => i !== pasted));
    const actual = auditActual(rng, units, lastYear);
    return { forecast, lastYear, actual, pasted, wrongLink, pastedValue: auditGrid({ forecast, actual })[3][pasted] };
  },
  layout: (d) => ({
    blocks: [
      cells('A1', [['Driver', 'Forecast', 'Last year']], 'header'),
      cells('A2', column(DRIVER_LABELS), 'label'),
      ...driverGrid(d).map((row, i) => cells(`B${2 + i}`, [row], 'input', DRIVER_FORMATS[i])),
      cells('A7', [['Line', 'Last actual', ...FORECAST_MONTHS]], 'header', MONTH_HEADER_FORMATS),
      cells('A8', column(AUDIT_LINES), 'label'),
      cells('B8', auditActualColumn(d).slice(0, 6), 'input', FMT.int),
      cells('B14', auditActualColumn(d).slice(6), 'input', FMT.pct),
      ...auditModel(d),
    ],
    answer: { kind: 'bugHunt', range: AUDIT_RANGE, bugs: auditBugs(d) },
  }),
  expected: auditGrid,
  inputs: (d) => [rangeWrite('B2:C5', driverGrid(d)), rangeWrite('B8:B14', auditActualColumn(d))],
  variants: [
    {
      label: 'the growth and price drivers change',
      explain: AUDIT_EXPLAIN.drivers,
      apply: (d, rng) => ({
        ...d,
        forecast: { ...d.forecast, growth: redraw(d.forecast.growth, () => drawGrowth(rng)), price: redraw(d.forecast.price, () => drawPrice(rng)) },
      }),
    },
    {
      label: 'the forecast COGS % changes',
      explain: AUDIT_EXPLAIN.cogs,
      apply: (d, rng) => {
        let cogsPct = drawCogs(rng);
        while (cogsPct === d.forecast.cogsPct || cogsPct === d.lastYear.cogsPct) cogsPct = drawCogs(rng);
        return { ...d, forecast: { ...d.forecast, cogsPct } };
      },
    },
    {
      label: 'last year’s figures change',
      explain: AUDIT_EXPLAIN.lastYear,
      apply: (d, rng) => {
        let lastYear = lastYearFor(rng, d.forecast);
        while (lastYear.cogsPct === d.lastYear.cogsPct) lastYear = lastYearFor(rng, d.forecast);
        return { ...d, lastYear };
      },
    },
    {
      label: 'the last actual units change',
      explain: AUDIT_EXPLAIN.drivers,
      apply: (d, rng) => {
        const units = redraw(d.actual.units, () => drawUnits(rng));
        return { ...d, actual: auditActual(rng, units, d.lastYear) };
      },
    },
  ],
  rules: { allowNumbers: [0, 1] },
});

export const MODELING: Exercise<any>[] = [modelScenarioSwitch, modelDriverForecast, modelCheckCells, modelAuditTrace];

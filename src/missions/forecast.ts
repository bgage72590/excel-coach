import { sum } from '../engine/data';
import type { Rng } from '../engine/rng';
import type { Grid, Inspection, Variant } from '../engine/types';
import { FMT, cells, column, rangeWrite } from '../exercises/common';
import {
  ASSUMPTIONS,
  FORECAST_MONTHS,
  LINES,
  LIVE_RULES,
  MONTH_HEADER_FORMATS,
  SCENARIOS,
  actualColumn,
  drawPrice,
  drawUnits,
  fixedFor,
  forecastLines,
  liveValues,
  makeActual,
  monthLabel,
  newAssumptions,
  redraw,
  scenarioDropdown,
  scenarioGrid,
  scenarioLeftOn,
  scenarioTable,
  swapUpsideDownside,
  switchTo,
  type LastActual,
  type Switchable,
} from '../exercises/modeling';
import { defineMission, type Mission } from './types';

// =====================================================================
// Forecast for the quarterly ops review
// =====================================================================

export interface OpsForecastData extends Switchable {
  /** Today's price per unit (H4), before the scenario's price change. */
  price: number;
  /** Today's opex per month (H5), before the scenario's opex change. */
  opex: number;
  /** September's actuals: the forecast's starting column, B10:B15. */
  actual: LastActual;
}

/** Three months per quarter: October to December, then January to March, and so on. */
const QUARTERS = [0, 1, 2, 3].map((q) => `Revenue, ${monthLabel(FORECAST_MONTHS[q * 3]).slice(0, 3)}–${monthLabel(FORECAST_MONTHS[q * 3 + 2])}`);
const OUTPUT_LABELS = [...QUARTERS, 'Annual revenue', 'Annual EBITDA', 'EBITDA margin %'];

/**
 * The forecast in LINES order, in the order Excel evaluates the step 2 formulas:
 * Revenue =C10*$H$4*(1+$E$5) and Opex =$H$5*(1+$E$7), with growth and COGS % from the Live column.
 */
export function opsForecast(d: OpsForecastData): number[][] {
  const [growth, priceChange, cogsPct, opexChange] = liveValues(d);
  return forecastLines(d.actual.units, growth, (u) => u * d.price * (1 + priceChange), cogsPct, d.opex * (1 + opexChange));
}

/** B18:B24: each quarter's revenue, annual revenue, annual EBITDA and EBITDA margin %. */
export function opsSummary(d: OpsForecastData): number[] {
  const lines = opsForecast(d);
  const revenue = lines[1];
  const quarters = [0, 1, 2, 3].map((q) => sum(revenue.slice(q * 3, q * 3 + 3)));
  const annualRevenue = sum(revenue);
  const annualEbitda = sum(lines[5]);
  return [...quarters, annualRevenue, annualEbitda, annualEbitda / annualRevenue];
}

/** The starting point at today's price and opex: September ran at the Base COGS %, give or take. */
function opsStart(rng: Rng, values: OpsForecastData['values']): Pick<OpsForecastData, 'price' | 'opex' | 'actual'> {
  const units = drawUnits(rng);
  const price = drawPrice(rng);
  const opex = fixedFor(rng, units, price, values.Base[2]);
  return { price, opex, actual: makeActual(rng, units, price, values.Base[2], opex) };
}

const startGrid = (d: OpsForecastData): Grid => [[d.price], [d.opex]];

const SWITCH = {
  forecast: 'Read only the Live column, E4:E7, so the switch in B1 drives every month.',
  summary: 'Total the forecast rows so the summary follows the switch.',
  check: 'Total the quarters and the Revenue row with SUM so the check stays 0 in every scenario.',
};

/** B18:B21 as a range, or the four quarters added one by one. */
const QUARTERS_REF = /(?<![A-Z$])\$?B\$?18:\$?B\$?21(?!\d)|(?<![A-Z$])\$?B\$?18\s*[+,]\s*\$?B\$?19\s*[+,]\s*\$?B\$?20\s*[+,]\s*\$?B\$?21(?!\d)/i;
/** The forecast's Revenue row, October through September. */
const REVENUE_ROW_REF = /(?<![A-Z$])\$?C\$?11:\$?N\$?11(?!\d)/i;

const marginAsPercent: Inspection = {
  kind: 'sheet',
  check: { kind: 'numberFormat', range: 'B24', matches: /%/, describe: 'a percentage' },
  label: 'The margin shows as a percentage',
  advice: 'Select B24 and choose Home › Percent Style.',
};

const upside = (explain: string) => switchTo<OpsForecastData>('Upside', explain);
const downside = (explain: string) => switchTo<OpsForecastData>('Downside', explain);

const startChanges: Variant<OpsForecastData> = {
  label: 'the starting point changes',
  explain: 'Grow month 1 from the Last actual units in B10, and read today’s price and opex from H4 and H5.',
  apply: (d, rng) => {
    const units = redraw(d.actual.units, () => drawUnits(rng));
    const price = redraw(d.price, () => drawPrice(rng));
    const opex = d.opex + rng.pick([-3000, -2000, -1000, 1000, 2000, 3000]);
    return { ...d, price, opex, actual: makeActual(rng, units, price, d.values.Base[2], opex) };
  },
};

export const opsForecastReview = defineMission<OpsForecastData>({
  id: 'mission-ops-forecast',
  title: 'Forecast for the quarterly ops review',
  role: 'finance',
  summary: 'A 12-month forecast with a Base, Upside and Downside switch, an outputs summary and a check cell',
  minutes: 25,
  skills: ['model-scenario-switch', 'model-driver-forecast', 'model-check-cells'],
  brief: () => ({
    from: 'Elena Brooks, VP of Operations',
    subject: 'Forecast for the quarterly ops review',
    body: [
      'Hi,',
      'For the quarterly ops review I want a 12-month forecast we can flip between cases in the room. The Base, Upside and Downside assumptions are in `B3:D7`. The starting point is filled in: September’s actuals in `B10:B15` and today’s price and opex in `H4:H5`.',
      'Make `B1` a dropdown of the three cases and pull the chosen case into the Live column. Build October through September off Live only, so changing `B1` moves every number. Then give me a summary I can read out: revenue by quarter, annual revenue, EBITDA and EBITDA margin, with a check that the quarters add back to the forecast.',
      'Leave it on Base when you send it over.',
      'Elena',
    ].join('\n\n'),
  }),
  make: (rng) => {
    const values = scenarioTable(rng);
    return { scenario: 'Base', columns: [...SCENARIOS], values, ...opsStart(rng, values) };
  },
  // Assumptions top left with the starting point beside them (column F stays empty), the forecast
  // below, then the summary and its check under the forecast's label column.
  blocks: (d) => [
    cells('A1', [['Scenario']], 'label'),
    cells('B1', [[d.scenario]], 'input'),
    cells('A3', [['Assumption', ...d.columns, 'Live']], 'header'),
    cells('A4', column(ASSUMPTIONS), 'label'),
    cells('B4', scenarioGrid(d).slice(1), 'input', FMT.pct),
    cells('G3', [['Starting point', 'Value']], 'header'),
    cells('G4', column(['Price per unit today', 'Opex per month today']), 'label'),
    cells('H4', [[d.price]], 'input', FMT.currency),
    cells('H5', [[d.opex]], 'input', '$#,##0'),
    cells('A9', [['Forecast', 'Last actual', ...FORECAST_MONTHS]], 'header', MONTH_HEADER_FORMATS),
    cells('A10', column(LINES), 'label'),
    cells('B10', column(actualColumn(d.actual)), 'input', FMT.int),
    cells('A17', [['Output', 'Value']], 'header'),
    cells('A18', column(OUTPUT_LABELS), 'label'),
    cells('A26', [['Check: quarters minus forecast revenue']], 'label'),
  ],
  inputs: (d) => [
    rangeWrite('B1', [[d.scenario]]),
    rangeWrite('B3:D7', scenarioGrid(d)),
    rangeWrite('H4:H5', startGrid(d)),
    rangeWrite('B10:B15', column(actualColumn(d.actual))),
  ],
  steps: [
    // ---------- 1. Scenario switch ----------
    {
      title: 'Scenario switch and Live column',
      task: () =>
        'Make `B1` a dropdown of the three scenarios: Data › Data Validation › Allow: List, with the names in `B3:D3` as the source. Then fill the Live column, `E4:E7`, with each assumption for the scenario in `B1`: one XLOOKUP (or INDEX and MATCH) in `E4` that finds `B1` in the header row, filled down. Leave `B1` on Base when you check.',
      hints: [
        'Select `B1`, then Data › Data Validation. Set Allow to List and Source to `=$B$3:$D$3`.',
        'In `E4`, find `B1` in the header row and return from the same row’s values. Press {absKey} to lock `B1` and the header row before you fill down.',
        '`=XLOOKUP($B$1, $B$3:$D$3, B4:D4)`, then fill down to `E7`.',
      ],
      solution: () => 'Data › Data Validation on B1: Allow List, Source =$B$3:$D$3. E4: =XLOOKUP($B$1,$B$3:$D$3,B4:D4), filled down to E7.',
      answer: () => ({ kind: 'cells', range: 'E4:E7', format: FMT.pct, consistency: 'all' }),
      expected: (d) => liveValues(d).map((v) => [v]),
      variants: [
        upside('Look up the scenario in B1 so the Live column follows the switch.'),
        downside('Look up the scenario in B1 so the Live column follows the switch.'),
        newAssumptions('Return the values from the assumption rows instead of typing them.'),
        swapUpsideDownside(),
      ],
      rules: LIVE_RULES,
      inspections: (d) => [scenarioDropdown('B1 has a dropdown of Base, Upside and Downside'), scenarioLeftOn(d)],
    },

    // ---------- 2. Forecast from Live ----------
    {
      title: '12-month forecast from the Live column',
      task: () =>
        'Build the forecast in `C10:N15`. Units grow each month by the Live growth in `E4`, starting from the Last actual units in `B10`. Revenue is Units × today’s price in `H4` × (1 + the Live price change in `E5`). COGS is Revenue × the Live COGS % in `E6`. Gross margin is Revenue − COGS. Opex is today’s opex in `H5` × (1 + the Live opex change in `E7`). EBITDA is Gross margin − Opex. Write each line once in column `C` and fill it right to `N`. Read the Live column, never Base, Upside or Downside, so the switch in `B1` drives every month. Don’t round.',
      hints: [
        'Units in `C10`: `=B10*(1+$E$4)`. Lock each Live cell and each starting-point cell with $ ({absKey}) so every month reads it as you fill right.',
        'Revenue `=C10*$H$4*(1+$E$5)`, COGS `=C11*$E$6`, Gross margin `=C11-C12`, Opex `=$H$5*(1+$E$7)`, EBITDA `=C13-C14`.',
        'Select `C10:C15` and fill right to column `N`, so each row holds one formula.',
      ],
      solution: () =>
        'C10: =B10*(1+$E$4) · C11: =C10*$H$4*(1+$E$5) · C12: =C11*$E$6 · C13: =C11-C12 · C14: =$H$5*(1+$E$7) · C15: =C13-C14, then fill C10:C15 right to column N.',
      answer: () => ({ kind: 'cells', range: 'C10:N15', format: FMT.int, consistency: 'rows' }),
      expected: opsForecast,
      variants: [
        upside(SWITCH.forecast),
        downside(SWITCH.forecast),
        newAssumptions<OpsForecastData>('Point at the Live column, which follows the assumptions, instead of typing rates.'),
        startChanges,
      ],
      rules: { allowNumbers: [0, 1] },
      inspections: (d) => [scenarioLeftOn(d)],
    },

    // ---------- 3. Outputs summary ----------
    {
      title: 'Outputs summary',
      task: () =>
        'Summarize the forecast in `B18:B24`: revenue for each quarter (three months each, starting with October), annual revenue, annual EBITDA, and EBITDA margin % (annual EBITDA ÷ annual revenue). Total the forecast rows so every figure follows the switch. Show the margin as a percentage: select `B24` and choose Home › Percent Style.',
      hints: [
        'Each quarter is a SUM over three month columns of the Revenue row: `=SUM(C11:E11)` for October to December, `=SUM(F11:H11)` for January to March.',
        'Annual revenue is `=SUM(C11:N11)` and annual EBITDA is `=SUM(C15:N15)`. The margin divides one by the other: `=B23/B22`.',
      ],
      solution: () =>
        'B18: =SUM(C11:E11) · B19: =SUM(F11:H11) · B20: =SUM(I11:K11) · B21: =SUM(L11:N11) · B22: =SUM(C11:N11) · B23: =SUM(C15:N15) · B24: =B23/B22',
      answer: () => ({ kind: 'cells', range: 'B18:B24', format: FMT.int, consistency: 'none' }),
      expected: (d) => opsSummary(d).map((v) => [v]),
      variants: [upside(SWITCH.summary), downside(SWITCH.summary), startChanges],
      // 0 lets the common =IFERROR(B23/B22,0) margin through.
      rules: { allowNumbers: [0] },
      inspections: (d) => [scenarioLeftOn(d), marginAsPercent],
    },

    // ---------- 4. Check cell ----------
    {
      title: 'Check that the quarters tie',
      task: () =>
        'In `B26`, add a check cell: the four quarters in `B18:B21` added up, minus the total of the forecast’s Revenue row, `C11:N11`. It reads 0 while the summary ties to the forecast, and moves off 0 the moment a quarter misses a month.',
      hints: ['Add up the quarters with one SUM and the Revenue row with another, then subtract.', '`=SUM(B18:B21)-SUM(C11:N11)`'],
      solution: () => '=SUM(B18:B21)-SUM(C11:N11)',
      answer: () => ({ kind: 'cells', range: 'B26', format: '#,##0.00', consistency: 'none' }),
      // 0 whatever the scenario. Adding the months in quarters can leave float noise far inside
      // the 1e-7 tolerance, so the check reads 0 on every run.
      expected: () => [[0]],
      variants: [upside(SWITCH.check), downside(SWITCH.check)],
      // Every variant leaves a correct check at 0, so the rules make sure it reads both sides:
      // =SUM(A2), or the annual figure minus the Revenue row it was built from, can never fire.
      rules: {
        require: [
          { pattern: QUARTERS_REF, label: 'Adds up the four quarters', advice: 'Total the quarters in B18:B21.' },
          {
            pattern: REVENUE_ROW_REF,
            label: 'Totals the forecast’s Revenue row',
            advice: 'Subtract SUM(C11:N11), the Revenue row, not the annual figure built from it.',
          },
        ],
        // The check-cells tolerances, so =ROUND(…, 2) passes as that exercise teaches.
        allowNumbers: [0, 0.005, 0.01, 0.5, 1, 2],
      },
    },
  ],
});

export const FORECAST_MISSIONS: Mission<any>[] = [opsForecastReview];

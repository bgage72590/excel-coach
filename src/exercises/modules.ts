import type { ModuleInfo } from '../engine/types';

export const MODULES: ModuleInfo[] = [
  { id: 'tables', title: 'Tables', blurb: 'Ranges that grow with your data, and formulas you can read.' },
  { id: 'sumifs', title: 'SUMIFS', blurb: 'Totals for any combination of conditions.' },
  { id: 'lookups', title: 'Lookups', blurb: 'Pull matching values from another table.' },
  { id: 'dynamic', title: 'Dynamic arrays', blurb: 'One formula that returns a whole list.' },
  { id: 'cleaning', title: 'Cleaning text', blurb: 'Fix messy exports with formulas instead of retyping.' },
  { id: 'dates', title: 'Dates and schedules', blurb: 'Business days, due dates and lead times.' },
  { id: 'letlambda', title: 'LET and LAMBDA', blurb: 'Readable formulas and functions of your own.' },
  { id: 'pivots', title: 'PivotTables', blurb: 'Summaries you can reshape in seconds.' },
  { id: 'powerquery', title: 'Power Query', blurb: 'Clean and combine data once, then refresh it every month.' },
  { id: 'whatif', title: 'What-if analysis', blurb: 'Goal Seek, data tables and Solver.' },
  { id: 'charts', title: 'Charts and dashboards', blurb: 'Charts and slicers that tell the story.' },
  { id: 'modeling', title: 'Modeling and auditing', blurb: 'Driver-based forecasts, scenario switches and checks you can trust.' },
  { id: 'bughunt', title: 'Bug hunts', blurb: 'Find and fix the mistakes in a workbook someone else built.' },
  { id: 'macros', title: 'Macros', blurb: 'Record the weekly formatting once, then make it work on any report.' },
];

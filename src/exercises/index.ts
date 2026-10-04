import type { Exercise, ModuleId } from '../engine/types';
import { BUG_HUNTS } from './bughunt';
import { comboChart, pivotAverageFilter, slicerDashboard } from './charts';
import { cleanNames, cleanSplit } from './cleaning';
import { eomonthTerms, networkdaysActual, workdayReceipt } from './dates';
import { byrowPeak, lookupTwoWay, pivotbyVendorCategory, xlookupTiered } from './depth';
import { filterLate, groupbyCategory, spillReference, uniqueVendors } from './dynamic';
import { lambdaVarpct, letReorder } from './letlambda';
import { lookupLeft, xlookupLeadTime, xlookupNotFound, xlookupTwoKeys } from './lookups';
import { MACROS } from './macros';
import { MODELING } from './modeling';
import { MODULES } from './modules';
import { pivotShare } from './pivots';
import { pqAppend, pqCleanExport, pqGroup, pqMerge, pqUnpivot } from './powerquery';
import { sumifsGrid, sumifsMonth, sumifsWarehouse } from './sumifs';
import { tablesCalcColumn, tablesConvert } from './tables';
import { dataTableTwoWay, goalSeekPrice, solverShipping } from './whatif';

/** Curriculum order: each exercise builds on the ones before it. */
export const EXERCISES: Exercise<any>[] = [
  tablesConvert,
  tablesCalcColumn,
  sumifsWarehouse,
  sumifsGrid,
  sumifsMonth,
  xlookupLeadTime,
  xlookupNotFound,
  xlookupTwoKeys,
  lookupLeft,
  xlookupTiered,
  lookupTwoWay,
  filterLate,
  uniqueVendors,
  spillReference,
  groupbyCategory,
  byrowPeak,
  pivotbyVendorCategory,
  cleanNames,
  cleanSplit,
  workdayReceipt,
  networkdaysActual,
  eomonthTerms,
  letReorder,
  lambdaVarpct,
  pivotShare,
  pivotAverageFilter,
  pqCleanExport,
  pqUnpivot,
  pqMerge,
  pqAppend,
  pqGroup,
  goalSeekPrice,
  dataTableTwoWay,
  solverShipping,
  comboChart,
  slicerDashboard,
  ...MODELING,
  ...BUG_HUNTS,
  ...MACROS,
];

const byId = new Map(EXERCISES.map((e) => [e.id, e]));

export function getExercise(id: string): Exercise<any> | undefined {
  return byId.get(id);
}

export function exercisesIn(module: ModuleId): Exercise<any>[] {
  return EXERCISES.filter((e) => e.module === module);
}

export { MODULES };

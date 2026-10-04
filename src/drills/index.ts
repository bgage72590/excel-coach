import { sortAndClean } from './clean';
import { numberFormats } from './formats';
import { sheetSetup } from './setup';
import type { Drill, DrillItem } from './types';

export type { Drill, DrillItem };
export { drillItemId, formatDrillGap, formatDrillTime } from './common';
export {
  elapsedMs,
  nextAction,
  pauseClock,
  recordRun,
  resumeClock,
  runTotalMs,
  runningSince,
  startClock,
  uncheckable,
  type DrillClock,
  type ItemTime,
  type RunResult,
} from './run';

/** Drill sets in the order they're offered. */
export const DRILLS: Drill[] = [sheetSetup, sortAndClean, numberFormats];

export function getDrill(id: string): Drill | undefined {
  return DRILLS.find((d) => d.id === id);
}

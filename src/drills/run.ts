import { sum } from '../engine/data';
import { recordDrillRun, type ProgressState } from '../engine/progress';
import type { CheckReport } from '../engine/types';
import { drillItemId } from './common';
import type { Drill } from './types';

/**
 * The timing side of a drill run, kept out of the view so it can be tested without Excel or React.
 * Each action has its own clock. It starts when the practice sheet is ready, pauses while the coach
 * checks the sheet, and stops for good at the Done that passes.
 */

// ---------- an action's clock ----------

export interface DrillClock {
  startedAt: number;
  /** Time spent in checks that didn't pass. It doesn't count. */
  pausedMs: number;
  /** When the check under way started (Done was selected). Set while the clock is paused. */
  checkAt?: number;
}

/** A clock that starts at `now`, when the practice sheet is ready. */
export const startClock = (now: number): DrillClock => ({ startedAt: now, pausedMs: 0 });

/** Pauses the clock while the coach checks the sheet. Pausing a paused clock changes nothing. */
export const pauseClock = (clock: DrillClock, now: number): DrillClock => (clock.checkAt === undefined ? { ...clock, checkAt: now } : clock);

/** Starts the clock again after a check that didn't pass. The check's own time doesn't count. */
export const resumeClock = (clock: DrillClock, now: number): DrillClock =>
  clock.checkAt === undefined ? clock : { startedAt: clock.startedAt, pausedMs: clock.pausedMs + Math.max(0, now - clock.checkAt) };

/** Time on the action: up to `now` while the clock runs, up to Done while it's paused. */
export const elapsedMs = (clock: DrillClock, now: number): number => Math.max(0, (clock.checkAt ?? now) - clock.startedAt - clock.pausedMs);

/**
 * When a running clock would have read zero, so its reading is `now` minus this. It changes each
 * time the clock starts again; undefined while the clock is paused.
 */
export const runningSince = (clock: DrillClock): number | undefined => (clock.checkAt === undefined ? clock.startedAt + clock.pausedMs : undefined);

// ---------- a run ----------

/** A finished action's time, or null when this version of Excel couldn't check it and it was skipped. */
export type ItemTime = number | null;

/** The run's time so far: every checked action's time added up. */
export const runTotalMs = (times: readonly ItemTime[]): number => sum(times.filter((t): t is number => t !== null));

/**
 * True when a check passed without checking anything: every item was skipped because this version
 * of Excel can't report what the action changes. That isn't a pass, and the action can't be timed.
 */
export const uncheckable = (report: CheckReport): boolean => report.items.length > 0 && report.items.every((i) => i.status === 'skip');

/** The action to set up next once `times` holds every finished one, or null when the run is over. */
export const nextAction = (drill: Drill, times: readonly ItemTime[]): number | null => (times.length < drill.items.length ? times.length : null);

export interface RunResult {
  state: ProgressState;
  /** Every checked action's time added up. */
  totalMs: number;
  /** Actions skipped because this version of Excel couldn't check them. */
  skipped: number;
  /** A new best total. Only a run with every action checked can set one. */
  personalBest: boolean;
  /** Item ids that beat an earlier best. */
  newItemBests: string[];
  /** The best total before this run, if there was one. */
  previousBestMs?: number;
}

/**
 * Records a finished run in progress. A skipped action has no time, so a run with one can't set the
 * best total: it would beat a full run by leaving actions out. Its checked actions still count
 * toward their own bests.
 */
export function recordRun(state: ProgressState, drill: Drill, times: readonly ItemTime[], now: number): RunResult {
  const itemMs: Record<string, number> = {};
  drill.items.forEach((item, i) => {
    const ms = times[i];
    if (typeof ms === 'number') itemMs[drillItemId(item)] = ms;
  });
  const totalMs = sum(Object.values(itemMs));
  const skipped = drill.items.length - Object.keys(itemMs).length;
  const previousBestMs = state.drills?.[drill.id]?.bestMs;
  const recorded = recordDrillRun(state, drill.id, totalMs, itemMs, now);
  if (!skipped) return { ...recorded, totalMs, skipped, previousBestMs };

  const entry = { ...recorded.state.drills![drill.id], bestMs: previousBestMs };
  return {
    state: { ...recorded.state, drills: { ...recorded.state.drills, [drill.id]: entry } },
    totalMs,
    skipped,
    personalBest: false,
    newItemBests: recorded.newItemBests,
    previousBestMs,
  };
}

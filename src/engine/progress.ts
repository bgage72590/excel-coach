/**
 * Learner progress, kept in localStorage. Mastery = passing on three different
 * data sets. Mastered skills come back for review at growing intervals.
 */

export const MASTERY_PASSES = 3;
const DAY = 86_400_000;
const STORAGE_KEY = 'excel-coach:progress:v1';

export interface ExerciseProgress {
  attempts: number;
  passSeeds: number[];
  mastered: boolean;
  bestMs?: number;
  lastPassAt?: number;
  /** When the third distinct pass landed. */
  masteredAt?: number;
  /** Next review time (epoch ms) once mastered. */
  reviewDue?: number;
  reviewIntervalDays?: number;
}

export interface Session {
  /** Exercise id, or for missions the current step's id ("mission-id#2"). */
  exerciseId: string;
  /** Set when this session is a mission. */
  missionId?: string;
  /** Zero-based step index for missions. */
  step?: number;
  /** Steps passed in this mission run. */
  stepsPassed?: number[];
  /** Mission steps whose onStart writes have been applied to the sheet. */
  startedSteps?: number[];
  /** Practice time is credited up to here; later credit starts from this point. */
  creditedAt?: number;
  seed: number;
  sheet: string;
  startedAt: number;
  attempts: number;
  hintsShown: number;
  revealed: boolean;
  passedAt?: number;
}

export interface MissionProgress {
  completions: number;
  bestMs?: number;
  lastCompletedAt?: number;
}

export interface DrillProgress {
  runs: number;
  bestMs?: number;
  lastRunAt?: number;
  /** Best time per drill item id. */
  bestItemMs?: Record<string, number>;
}

export interface ProgressState {
  version: 1;
  exercises: Record<string, ExerciseProgress>;
  missions?: Record<string, MissionProgress>;
  drills?: Record<string, DrillProgress>;
  /** Practice time per local day ("2026-10-04" → ms). */
  practice?: Record<string, number>;
  /** Consecutive passes on the first check with no hints. */
  cleanStreak: number;
  session?: Session;
}

export function emptyProgress(): ProgressState {
  return { version: 1, exercises: {}, cleanStreak: 0 };
}

export function loadProgress(): ProgressState {
  try {
    const raw = globalThis.localStorage?.getItem(STORAGE_KEY);
    if (!raw) return emptyProgress();
    const parsed = JSON.parse(raw) as ProgressState;
    return parsed?.version === 1 ? { ...emptyProgress(), ...parsed } : emptyProgress();
  } catch {
    return emptyProgress();
  }
}

export function saveProgress(state: ProgressState): void {
  try {
    globalThis.localStorage?.setItem(STORAGE_KEY, JSON.stringify(state));
  } catch {
    // Storage can be unavailable; progress then lasts for this session only.
  }
}

export function progressFor(state: ProgressState, id: string): ExerciseProgress {
  return state.exercises[id] ?? { attempts: 0, passSeeds: [], mastered: false };
}

export type SkillStatus = 'new' | 'practicing' | 'mastered' | 'review';

export function statusOf(p: ExerciseProgress, now: number): SkillStatus {
  if (p.mastered) return p.reviewDue !== undefined && p.reviewDue <= now ? 'review' : 'mastered';
  return p.passSeeds.length > 0 || p.attempts > 0 ? 'practicing' : 'new';
}

export interface PassOutcome {
  state: ProgressState;
  elapsedMs: number;
  personalBest: boolean;
  newlyMastered: boolean;
  wasReview: boolean;
  passesTowardMastery: number;
  clean: boolean;
}

/** Records a passing check for the current session. */
export function recordPass(state: ProgressState, now: number): PassOutcome {
  const session = state.session;
  if (!session) throw new Error('No active session');
  const prev = progressFor(state, session.exerciseId);
  const elapsedMs = now - session.startedAt;
  const counts = !session.revealed;
  const wasReview = prev.mastered && prev.reviewDue !== undefined && prev.reviewDue <= now;

  const passSeeds = counts && !prev.passSeeds.includes(session.seed) ? [...prev.passSeeds, session.seed] : prev.passSeeds;
  const newlyMastered = !prev.mastered && passSeeds.length >= MASTERY_PASSES;
  const mastered = prev.mastered || newlyMastered;

  let reviewIntervalDays = prev.reviewIntervalDays;
  let reviewDue = prev.reviewDue;
  if (newlyMastered) {
    reviewIntervalDays = 2;
    reviewDue = now + 2 * DAY;
  } else if (wasReview && counts) {
    reviewIntervalDays = Math.min(60, (prev.reviewIntervalDays ?? 2) * 2);
    reviewDue = now + reviewIntervalDays * DAY;
  }

  const personalBest = counts && (prev.bestMs === undefined || elapsedMs < prev.bestMs);
  const clean = counts && session.attempts <= 1 && session.hintsShown === 0;

  const next: ExerciseProgress = {
    ...prev,
    passSeeds,
    mastered,
    bestMs: personalBest ? elapsedMs : prev.bestMs,
    lastPassAt: now,
    masteredAt: newlyMastered ? now : prev.masteredAt,
    reviewDue,
    reviewIntervalDays,
  };

  const credited = creditPractice(state, now);
  return {
    state: {
      ...credited,
      exercises: { ...state.exercises, [session.exerciseId]: next },
      cleanStreak: clean ? state.cleanStreak + 1 : 0,
      session: { ...credited.session!, passedAt: now },
    },
    elapsedMs,
    personalBest,
    newlyMastered,
    wasReview,
    passesTowardMastery: Math.min(passSeeds.length, MASTERY_PASSES),
    clean,
  };
}

export function recordAttempt(state: ProgressState): ProgressState {
  const session = state.session;
  if (!session) return state;
  const prev = progressFor(state, session.exerciseId);
  return {
    ...state,
    exercises: { ...state.exercises, [session.exerciseId]: { ...prev, attempts: prev.attempts + 1 } },
    session: { ...session, attempts: session.attempts + 1 },
  };
}

export function formatDuration(ms: number): string {
  const total = Math.max(0, Math.round(ms / 1000));
  const m = Math.floor(total / 60);
  const s = total % 60;
  return m >= 60 ? `${Math.floor(m / 60)}:${String(m % 60).padStart(2, '0')}:${String(s).padStart(2, '0')}` : `${m}:${String(s).padStart(2, '0')}`;
}

// ---------- missions ----------

export function missionProgressFor(state: ProgressState, id: string): MissionProgress {
  return state.missions?.[id] ?? { completions: 0 };
}

export interface StepOutcome {
  state: ProgressState;
  missionComplete: boolean;
  elapsedMs: number;
  personalBest: boolean;
}

/** Records a passed mission step; completes the mission when every step has passed. */
export function recordStepPass(state: ProgressState, stepCount: number, now: number): StepOutcome {
  const session = state.session;
  if (!session || !session.missionId || session.step === undefined) throw new Error('No active mission');
  const stepsPassed = [...new Set([...(session.stepsPassed ?? []), session.step])].sort((a, b) => a - b);
  const missionComplete = stepsPassed.length >= stepCount;
  const elapsedMs = now - session.startedAt;
  const prev = missionProgressFor(state, session.missionId);
  const counts = !session.revealed;
  const personalBest = missionComplete && counts && (prev.bestMs === undefined || elapsedMs < prev.bestMs);
  const missions = missionComplete
    ? {
        ...state.missions,
        [session.missionId]: {
          completions: prev.completions + 1,
          bestMs: personalBest ? elapsedMs : prev.bestMs,
          lastCompletedAt: now,
        },
      }
    : state.missions;
  const credited = creditPractice(state, now);
  return {
    state: { ...credited, missions, session: { ...credited.session!, stepsPassed, passedAt: missionComplete ? now : undefined } },
    missionComplete,
    elapsedMs,
    personalBest,
  };
}

// ---------- practice time and the weekly summary ----------

/** One rep's credit is capped, so a panel left open overnight doesn't count as practice. */
export const PRACTICE_CAP_MS = 20 * 60_000;
const WEEK = 7 * DAY;

/** Local calendar day, e.g. "2026-10-04". */
export function dayKey(ms: number): string {
  const d = new Date(ms);
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
}

/** Adds practice time to today's total. */
export function recordPractice(state: ProgressState, ms: number, now: number): ProgressState {
  const add = Math.max(0, Math.min(ms, PRACTICE_CAP_MS));
  if (!add) return state;
  const key = dayKey(now);
  return { ...state, practice: { ...state.practice, [key]: (state.practice?.[key] ?? 0) + add } };
}

/** Credits the session's time since it started (or since the last credit) as practice. */
export function creditPractice(state: ProgressState, now: number): ProgressState {
  const session = state.session;
  if (!session) return state;
  const from = Math.max(session.startedAt, session.creditedAt ?? 0);
  return { ...recordPractice(state, now - from, now), session: { ...session, creditedAt: now } };
}

/** Records a finished drill run: total time, per-item bests, and practice time (one capped rep). */
export function recordDrillRun(
  state: ProgressState,
  drillId: string,
  totalMs: number,
  itemMs: Record<string, number>,
  now: number,
): { state: ProgressState; personalBest: boolean; newItemBests: string[] } {
  const prev = state.drills?.[drillId] ?? { runs: 0 };
  const personalBest = prev.bestMs === undefined || totalMs < prev.bestMs;
  const bestItemMs = { ...prev.bestItemMs };
  const newItemBests: string[] = [];
  for (const [id, ms] of Object.entries(itemMs)) {
    if (bestItemMs[id] === undefined || ms < bestItemMs[id]) {
      if (bestItemMs[id] !== undefined) newItemBests.push(id);
      bestItemMs[id] = ms;
    }
  }
  const next: DrillProgress = { runs: prev.runs + 1, bestMs: personalBest ? totalMs : prev.bestMs, lastRunAt: now, bestItemMs };
  const withTime = recordPractice(state, totalMs, now);
  return { state: { ...withTime, drills: { ...state.drills, [drillId]: next } }, personalBest, newItemBests };
}

export interface WeekSummary {
  /** Mastered skills whose review is due now. */
  reviewsDue: number;
  masteredTotal: number;
  masteredThisWeek: number;
  /** Distinct skills passed in the last 7 days. */
  skillsPracticed: number;
  /** Practice time over the last 7 days, today included. */
  practicedMs: number;
  /** Days with any practice in the last 7. */
  activeDays: number;
  /** Practice per day for the last 7 days, oldest first; for a small bar row. */
  days: { key: string; ms: number }[];
}

/** The home screen's weekly summary over the given exercises. */
export function weekSummary(state: ProgressState, exerciseIds: string[], now: number): WeekSummary {
  const since = now - WEEK;
  let reviewsDue = 0;
  let masteredTotal = 0;
  let masteredThisWeek = 0;
  let skillsPracticed = 0;
  for (const id of exerciseIds) {
    const p = progressFor(state, id);
    const status = statusOf(p, now);
    if (status === 'review') reviewsDue++;
    if (p.mastered) masteredTotal++;
    if (p.masteredAt !== undefined && p.masteredAt > since) masteredThisWeek++;
    if (p.lastPassAt !== undefined && p.lastPassAt > since) skillsPracticed++;
  }
  // Step back by calendar day, not 24 hours: the days around a DST change are 23 or 25 hours long.
  const today = new Date(now);
  const days = Array.from({ length: 7 }, (_, i) => {
    const key = dayKey(new Date(today.getFullYear(), today.getMonth(), today.getDate() - (6 - i), 12).getTime());
    return { key, ms: state.practice?.[key] ?? 0 };
  });
  const practicedMs = days.reduce((a, d) => a + d.ms, 0);
  return { reviewsDue, masteredTotal, masteredThisWeek, skillsPracticed, practicedMs, activeDays: days.filter((d) => d.ms > 0).length, days };
}

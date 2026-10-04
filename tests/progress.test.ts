import { afterAll, afterEach, beforeAll, beforeEach, describe, expect, it, vi } from 'vitest';
import { formatPracticeTime, hasProgress, weekSummaryLabel } from '../src/components/WeekSummary';
import {
  PRACTICE_CAP_MS,
  creditPractice,
  dayKey,
  emptyProgress,
  loadProgress,
  progressFor,
  recordDrillRun,
  recordPass,
  recordPractice,
  recordStepPass,
  saveProgress,
  weekSummary,
  type ExerciseProgress,
  type ProgressState,
  type Session,
} from '../src/engine/progress';

// Local noon keeps these tests on the intended calendar day in any time zone. Times near
// midnight around a DST change are covered separately, in a zone that has one.
const noon = (y: number, m: number, d: number) => new Date(y, m, d, 12).getTime();
const MIN = 60_000;
const HOUR = 60 * MIN;
const DAY = 86_400_000;
const NOW = noon(2026, 9, 4); // Sunday 4 October 2026

function withSession(state: ProgressState, session: Partial<Session> & Pick<Session, 'exerciseId' | 'startedAt'>): ProgressState {
  return { ...state, session: { seed: 1, sheet: 'Coach', attempts: 0, hintsShown: 0, revealed: false, ...session } };
}

const practicedOn = (state: ProgressState, at: number) => state.practice?.[dayKey(at)] ?? 0;

describe('practice time', () => {
  it('caps one rep at PRACTICE_CAP_MS', () => {
    const state = recordPractice(emptyProgress(), 3 * HOUR, NOW);
    expect(state.practice).toEqual({ '2026-10-04': PRACTICE_CAP_MS });
  });

  it('adds reps up per local day', () => {
    let state = recordPractice(emptyProgress(), 4 * MIN, noon(2026, 9, 3));
    state = recordPractice(state, 6 * MIN, NOW);
    state = recordPractice(state, 15 * MIN, NOW + 2 * HOUR);
    expect(state.practice).toEqual({ '2026-10-03': 4 * MIN, '2026-10-04': 21 * MIN });
  });

  it('ignores zero and negative time', () => {
    const state = emptyProgress();
    expect(recordPractice(state, 0, NOW)).toBe(state);
    expect(recordPractice(state, -5 * MIN, NOW)).toBe(state);
  });

  it('credits a mission’s time once across its steps', () => {
    const start = NOW - HOUR;
    let state = withSession(emptyProgress(), { exerciseId: 'm#0', missionId: 'm', step: 0, startedAt: start });
    state = recordStepPass(state, 3, start + 6 * MIN).state;
    expect(practicedOn(state, NOW)).toBe(6 * MIN);

    state = { ...state, session: { ...state.session!, exerciseId: 'm#1', step: 1 } };
    state = recordStepPass(state, 3, start + 10 * MIN).state;
    expect(practicedOn(state, NOW)).toBe(10 * MIN);

    // Crediting again at the same moment adds nothing.
    state = creditPractice(state, start + 10 * MIN);
    expect(practicedOn(state, NOW)).toBe(10 * MIN);

    state = { ...state, session: { ...state.session!, exerciseId: 'm#2', step: 2 } };
    const out = recordStepPass(state, 3, start + 25 * MIN);
    expect(out.missionComplete).toBe(true);
    expect(practicedOn(out.state, NOW)).toBe(25 * MIN);
    expect(out.state.session?.creditedAt).toBe(start + 25 * MIN);
  });

  it('counts a step left open for hours as one capped rep', () => {
    const start = NOW - 4 * HOUR;
    let state = withSession(emptyProgress(), { exerciseId: 'm#0', missionId: 'm', step: 0, startedAt: start });
    state = creditPractice(state, start + 3 * HOUR);
    expect(practicedOn(state, NOW)).toBe(PRACTICE_CAP_MS);
    state = creditPractice(state, start + 3 * HOUR + 5 * MIN);
    expect(practicedOn(state, NOW)).toBe(PRACTICE_CAP_MS + 5 * MIN);
  });

  it('leaves state alone with no session', () => {
    const state = emptyProgress();
    expect(creditPractice(state, NOW)).toBe(state);
  });

  it('credits the time since a mission step began', () => {
    const state = withSession(emptyProgress(), { exerciseId: 'm#0', missionId: 'm', step: 0, startedAt: NOW - 8 * MIN });
    const out = recordStepPass(state, 2, NOW);
    expect(out.missionComplete).toBe(false);
    expect(out.state.practice).toEqual({ '2026-10-04': 8 * MIN });
    expect(out.state.session?.stepsPassed).toEqual([0]);
  });

  it('credits the time an exercise took when it passes', () => {
    const state = withSession(emptyProgress(), { exerciseId: 'x', startedAt: NOW - 7 * MIN });
    expect(recordPass(state, NOW).state.practice).toEqual({ '2026-10-04': 7 * MIN });
  });

  it('credits an older session without creditedAt from its start', () => {
    const state = withSession(emptyProgress(), { exerciseId: 'x', startedAt: NOW - 3 * MIN });
    expect(state.session?.creditedAt).toBeUndefined();
    const out = recordPass(state, NOW);
    expect(practicedOn(out.state, NOW)).toBe(3 * MIN);
    expect(out.state.session?.creditedAt).toBe(NOW);
  });
});

describe('mastery', () => {
  it('sets masteredAt only on the third distinct pass', () => {
    let state = emptyProgress();
    const pass = (seed: number, at: number, revealed = false) => {
      state = recordPass(withSession(state, { exerciseId: 'x', seed, startedAt: at - 2 * MIN, revealed }), at).state;
      return progressFor(state, 'x');
    };
    expect(pass(1, NOW - 5 * DAY).masteredAt).toBeUndefined();
    expect(pass(1, NOW - 4 * DAY).masteredAt).toBeUndefined(); // the same data again
    expect(pass(2, NOW - 3 * DAY, true).masteredAt).toBeUndefined(); // the answer was shown
    expect(pass(2, NOW - 2 * DAY).masteredAt).toBeUndefined();

    const third = pass(3, NOW - DAY);
    expect(third.mastered).toBe(true);
    expect(third.passSeeds).toEqual([1, 2, 3]);
    expect(third.masteredAt).toBe(NOW - DAY);

    const later = pass(4, NOW);
    expect(later.masteredAt).toBe(NOW - DAY);
    expect(later.lastPassAt).toBe(NOW);
  });
});

describe('drills', () => {
  it('tracks run bests and item bests', () => {
    const first = recordDrillRun(emptyProgress(), 'basics', 90_000, { freeze: 20_000, filter: 40_000, sort: 30_000 }, NOW - DAY);
    expect(first.personalBest).toBe(true);
    expect(first.newItemBests).toEqual([]); // a first time sets the bar; it isn't a new best
    expect(first.state.drills?.basics).toEqual({
      runs: 1,
      bestMs: 90_000,
      lastRunAt: NOW - DAY,
      bestItemMs: { freeze: 20_000, filter: 40_000, sort: 30_000 },
    });

    const second = recordDrillRun(first.state, 'basics', 95_000, { freeze: 15_000, filter: 45_000, sort: 30_000, format: 5_000 }, NOW);
    expect(second.personalBest).toBe(false);
    expect(second.newItemBests).toEqual(['freeze']); // a tie isn't a new best
    expect(second.state.drills?.basics).toEqual({
      runs: 2,
      bestMs: 90_000,
      lastRunAt: NOW,
      bestItemMs: { freeze: 15_000, filter: 40_000, sort: 30_000, format: 5_000 },
    });

    const third = recordDrillRun(second.state, 'basics', 80_000, { freeze: 16_000, filter: 34_000, sort: 25_000, format: 5_000 }, NOW + MIN);
    expect(third.personalBest).toBe(true);
    expect(third.newItemBests).toEqual(['filter', 'sort']);
    expect(third.state.drills?.basics.bestMs).toBe(80_000);
    expect(third.state.drills?.basics.runs).toBe(3);
  });

  it('keeps each drill’s records apart', () => {
    let state = recordDrillRun(emptyProgress(), 'basics', 60_000, { a: 60_000 }, NOW).state;
    state = recordDrillRun(state, 'format', 45_000, { b: 45_000 }, NOW).state;
    expect(state.drills?.basics).toMatchObject({ runs: 1, bestMs: 60_000 });
    expect(state.drills?.format).toMatchObject({ runs: 1, bestMs: 45_000 });
  });

  it('credits a run as practice, capped', () => {
    expect(practicedOn(recordDrillRun(emptyProgress(), 'basics', 4 * MIN, {}, NOW).state, NOW)).toBe(4 * MIN);
    expect(practicedOn(recordDrillRun(emptyProgress(), 'basics', 15 * MIN, {}, NOW).state, NOW)).toBe(15 * MIN);
    // A run counts as one rep: drills take a minute or two, so a long run is a pane left open.
    expect(practicedOn(recordDrillRun(emptyProgress(), 'basics', 25 * MIN, {}, NOW).state, NOW)).toBe(PRACTICE_CAP_MS);
    expect(practicedOn(recordDrillRun(emptyProgress(), 'basics', 3 * HOUR, {}, NOW).state, NOW)).toBe(PRACTICE_CAP_MS);
  });
});

describe('weekSummary', () => {
  it('lists the last seven local days, oldest first, today last', () => {
    const { days } = weekSummary(emptyProgress(), [], NOW);
    expect(days.map((d) => d.key)).toEqual(['2026-09-28', '2026-09-29', '2026-09-30', '2026-10-01', '2026-10-02', '2026-10-03', '2026-10-04']);
    expect(days[6].key).toBe(dayKey(NOW));
  });

  it('crosses month and year boundaries', () => {
    const { days } = weekSummary(emptyProgress(), [], noon(2027, 0, 2));
    expect(days.map((d) => d.key)).toEqual(['2026-12-27', '2026-12-28', '2026-12-29', '2026-12-30', '2026-12-31', '2027-01-01', '2027-01-02']);
  });

  it('adds up practice inside the window only', () => {
    const practice = {
      '2026-09-27': 50 * MIN, // eight days ago
      '2026-09-28': 10 * MIN,
      '2026-10-02': 25 * MIN,
      '2026-10-04': 5 * MIN,
      '2026-10-05': 90 * MIN, // tomorrow
    };
    const sum = weekSummary({ ...emptyProgress(), practice }, [], NOW);
    expect(sum.days.map((d) => d.ms)).toEqual([10 * MIN, 0, 0, 0, 25 * MIN, 0, 5 * MIN]);
    expect(sum.practicedMs).toBe(40 * MIN);
    expect(sum.activeDays).toBe(3);
  });

  it('counts reviews, mastery and skills practiced over the given exercises', () => {
    const mastered = (masteredAt: number, reviewDue: number): ExerciseProgress => ({
      attempts: 3,
      passSeeds: [1, 2, 3],
      mastered: true,
      lastPassAt: masteredAt,
      masteredAt,
      reviewDue,
      reviewIntervalDays: 2,
    });
    const exercises: Record<string, ExerciseProgress> = {
      due: mastered(NOW - 20 * DAY, NOW - DAY),
      dueNow: mastered(NOW - 10 * DAY, NOW),
      recent: mastered(NOW - 2 * DAY, NOW + 2 * DAY),
      edge: mastered(NOW - 7 * DAY, NOW + DAY), // exactly a week ago falls outside
      practicing: { attempts: 2, passSeeds: [4], mastered: false, lastPassAt: NOW - 3 * DAY },
      stale: { attempts: 1, passSeeds: [5], mastered: false, lastPassAt: NOW - 9 * DAY },
      tried: { attempts: 4, passSeeds: [], mastered: false },
      elsewhere: mastered(NOW - DAY, NOW - HOUR), // not in the list below
    };
    const ids = ['due', 'dueNow', 'recent', 'edge', 'practicing', 'stale', 'tried', 'never'];
    expect(weekSummary({ ...emptyProgress(), exercises }, ids, NOW)).toMatchObject({
      reviewsDue: 2,
      masteredTotal: 4,
      masteredThisWeek: 1,
      skillsPracticed: 2,
    });
  });

  it('is all zeros for a new learner', () => {
    expect(weekSummary(emptyProgress(), ['a', 'b'], NOW)).toMatchObject({
      reviewsDue: 0,
      masteredTotal: 0,
      masteredThisWeek: 0,
      skillsPracticed: 0,
      practicedMs: 0,
      activeDays: 0,
    });
  });
});

describe('weekSummary across daylight saving changes', () => {
  // A day around a change isn't 24 hours long, so stepping back from now in 24-hour steps
  // breaks for an hour a night after it. Pin a zone with DST so the suite catches that anywhere.
  const zone = process.env.TZ;
  beforeAll(() => {
    process.env.TZ = 'America/New_York';
  });
  afterAll(() => {
    if (zone === undefined) delete process.env.TZ;
    else process.env.TZ = zone;
  });

  const keys = (now: number) => weekSummary(emptyProgress(), [], now).days.map((d) => d.key);
  const nextDay = (key: string) => {
    const [y, m, d] = key.split('-').map(Number);
    return dayKey(new Date(y, m - 1, d + 1, 12).getTime());
  };

  it('runs in a zone that changes its clocks', () => {
    expect(new Date(2026, 0, 15).getTimezoneOffset()).toBe(300);
    expect(new Date(2026, 6, 15).getTimezoneOffset()).toBe(240);
  });

  it('keeps seven days in the hour before midnight after the clocks go back', () => {
    expect(keys(new Date(2026, 10, 1, 23, 30).getTime())).toEqual(['2026-10-26', '2026-10-27', '2026-10-28', '2026-10-29', '2026-10-30', '2026-10-31', '2026-11-01']);
  });

  it('keeps seven days in the hour after midnight after the clocks go forward', () => {
    expect(keys(new Date(2026, 2, 9, 0, 30).getTime())).toEqual(['2026-03-03', '2026-03-04', '2026-03-05', '2026-03-06', '2026-03-07', '2026-03-08', '2026-03-09']);
  });

  it('counts today’s practice once', () => {
    const practice = { '2026-10-31': 5 * MIN, '2026-11-01': 10 * MIN };
    const sum = weekSummary({ ...emptyProgress(), practice }, [], new Date(2026, 10, 1, 23, 30).getTime());
    expect(sum.practicedMs).toBe(15 * MIN);
    expect(sum.activeDays).toBe(2);
  });

  it('lists seven consecutive days, today last, at every quarter hour around both changes', () => {
    // From the day before each change through the six nights after it.
    for (const start of [new Date(2026, 2, 7), new Date(2026, 9, 31)]) {
      for (let at = start.getTime(); at < start.getTime() + 9 * DAY; at += 15 * MIN) {
        const days = keys(at);
        expect(days[6], new Date(at).toString()).toBe(dayKey(at));
        for (let i = 1; i < 7; i++) expect(days[i], new Date(at).toString()).toBe(nextDay(days[i - 1]));
      }
    }
  });
});

describe('saved progress', () => {
  // Changing this key would orphan everyone’s saved progress.
  const KEY = 'excel-coach:progress:v1';
  let store: Map<string, string>;

  beforeEach(() => {
    store = new Map();
    vi.stubGlobal('localStorage', {
      getItem: (k: string) => store.get(k) ?? null,
      setItem: (k: string, v: string) => void store.set(k, v),
      removeItem: (k: string) => void store.delete(k),
    });
  });
  afterEach(() => vi.unstubAllGlobals());

  it('loads a save from before practice time and drills', () => {
    store.set(
      KEY,
      JSON.stringify({
        version: 1,
        exercises: {
          tablesConvert: { attempts: 4, passSeeds: [7, 9, 11], mastered: true, bestMs: 61_000, lastPassAt: NOW - 3 * DAY, reviewDue: NOW - DAY, reviewIntervalDays: 2 },
          sumifsGrid: { attempts: 2, passSeeds: [3], mastered: false, lastPassAt: NOW - DAY },
        },
        missions: { 'month-end': { completions: 1, bestMs: 900_000, lastCompletedAt: NOW - 2 * DAY } },
        cleanStreak: 2,
      }),
    );
    const state = loadProgress();
    expect(state.cleanStreak).toBe(2);
    expect(state.practice).toBeUndefined();
    expect(state.drills).toBeUndefined();
    expect(progressFor(state, 'tablesConvert').passSeeds).toEqual([7, 9, 11]);

    // Mastered before masteredAt existed: it counts toward the total, not toward this week.
    expect(weekSummary(state, ['tablesConvert', 'sumifsGrid'], NOW)).toMatchObject({
      reviewsDue: 1,
      masteredTotal: 1,
      masteredThisWeek: 0,
      skillsPracticed: 2,
      practicedMs: 0,
    });

    expect(recordPractice(state, 5 * MIN, NOW).practice).toEqual({ '2026-10-04': 5 * MIN });
    expect(recordDrillRun(state, 'basics', MIN, { a: MIN }, NOW).state.drills?.basics.runs).toBe(1);
  });

  it('fills in fields an older save is missing', () => {
    store.set(KEY, JSON.stringify({ version: 1 }));
    expect(loadProgress()).toEqual(emptyProgress());
    store.set(KEY, JSON.stringify({ version: 1, exercises: { x: { attempts: 1, passSeeds: [], mastered: false } } }));
    expect(loadProgress().cleanStreak).toBe(0);
  });

  it('starts fresh from a damaged or unknown save', () => {
    expect(loadProgress()).toEqual(emptyProgress());
    store.set(KEY, '{"version":1,');
    expect(loadProgress()).toEqual(emptyProgress());
    store.set(KEY, JSON.stringify({ version: 2, exercises: { x: { attempts: 1, passSeeds: [1], mastered: false } }, cleanStreak: 0 }));
    expect(loadProgress()).toEqual(emptyProgress());
  });

  it('round-trips practice time and drills', () => {
    let state = recordPractice(emptyProgress(), 12 * MIN, NOW);
    state = recordDrillRun(state, 'basics', 70_000, { freeze: 10_000 }, NOW).state;
    saveProgress(state);
    expect(loadProgress()).toEqual(state);
  });
});

describe('weekly summary card', () => {
  it('formats practice time', () => {
    expect(formatPracticeTime(85 * MIN)).toBe('1 h 25 min');
    expect(formatPracticeTime(35 * MIN + 50_000)).toBe('35 min');
    expect(formatPracticeTime(2 * HOUR)).toBe('2 h');
    expect(formatPracticeTime(59_999)).toBe('Under a minute');
    expect(formatPracticeTime(MIN)).toBe('1 min');
    expect(formatPracticeTime(0)).toBe('0 min');
  });

  it('spells the units out for screen readers', () => {
    expect(formatPracticeTime(85 * MIN, true)).toBe('1 hour 25 minutes');
    expect(formatPracticeTime(61 * MIN, true)).toBe('1 hour 1 minute');
    expect(formatPracticeTime(2 * HOUR, true)).toBe('2 hours');
    expect(formatPracticeTime(30_000, true)).toBe('under a minute');
  });

  it('sums the card up in one sentence', () => {
    const base = weekSummary(emptyProgress(), [], NOW);
    expect(weekSummaryLabel({ ...base, practicedMs: 85 * MIN, masteredTotal: 12, masteredThisWeek: 2, reviewsDue: 3 }, 36)).toBe(
      'You practiced 1 hour 25 minutes this week, have mastered 12 of 36 skills (2 this week), and have 3 reviews due.',
    );
    expect(weekSummaryLabel({ ...base, masteredTotal: 1, reviewsDue: 1 }, 36)).toBe(
      'You haven’t practiced this week, have mastered 1 of 36 skills, and have 1 review due.',
    );
    expect(weekSummaryLabel({ ...base, practicedMs: 20_000 }, 36)).toBe(
      'You practiced under a minute this week, have mastered 0 of 36 skills, and have no reviews due.',
    );
  });

  it('shows only once the learner has some progress', () => {
    expect(hasProgress(emptyProgress())).toBe(false);
    expect(hasProgress({ ...emptyProgress(), exercises: { x: { attempts: 3, passSeeds: [], mastered: false } } })).toBe(false);
    expect(hasProgress({ ...emptyProgress(), practice: { '2026-10-04': 0 } })).toBe(false);

    expect(hasProgress({ ...emptyProgress(), exercises: { x: { attempts: 1, passSeeds: [], mastered: false, lastPassAt: NOW } } })).toBe(true);
    expect(hasProgress({ ...emptyProgress(), exercises: { x: { attempts: 1, passSeeds: [4], mastered: false } } })).toBe(true);
    expect(hasProgress({ ...emptyProgress(), practice: { '2026-09-01': 4 * MIN } })).toBe(true);
    expect(hasProgress({ ...emptyProgress(), missions: { m: { completions: 1 } } })).toBe(true);
    expect(hasProgress({ ...emptyProgress(), drills: { d: { runs: 1 } } })).toBe(true);
  });
});

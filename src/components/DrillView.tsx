import {
  Badge,
  Body1,
  Button,
  Caption1,
  Link,
  MessageBar,
  MessageBarActions,
  MessageBarBody,
  MessageBarTitle,
  ProgressBar,
  Spinner,
  Text,
  Tooltip,
  makeStyles,
  mergeClasses,
  tokens,
  useId,
} from '@fluentui/react-components';
import {
  ArrowLeft16Regular,
  ArrowSync20Regular,
  CheckmarkCircle20Filled,
  CheckmarkCircle24Filled,
  ChevronRight16Regular,
  Dismiss16Regular,
  Next20Regular,
  Timer16Regular,
  Timer20Regular,
  Trophy24Filled,
} from '@fluentui/react-icons';
import { useCallback, useEffect, useLayoutEffect, useRef, useState, type ReactNode } from 'react';
import {
  DRILLS,
  drillItemId,
  elapsedMs,
  formatDrillGap,
  formatDrillTime,
  nextAction,
  pauseClock,
  recordRun,
  resumeClock,
  runTotalMs,
  runningSince,
  startClock,
  uncheckable,
  type Drill,
  type DrillClock,
  type DrillItem,
  type ItemTime,
  type RunResult,
} from '../drills';
import { formatDuration } from '../engine/progress';
import { Rng, newSeed } from '../engine/rng';
import type { CheckReport } from '../engine/types';
import { CoachError, friendlyError } from '../excel/errors';
import { sheetNameFor } from '../excel/host';
import { useCoach } from '../taskpane/context';
import { RichText, SectionLabel } from './bits';
import { CheckResults } from './CheckResults';
import { ConfirmDialog } from './ConfirmDialog';

/** 'uncheckable': a check came back with every item skipped, because this Excel can't report the change. */
type Phase = 'overview' | 'settingUp' | 'setupFailed' | 'working' | 'checking' | 'passed' | 'uncheckable' | 'complete';

/** How long a passed action stays on screen before the next sheet is set up. */
const ADVANCE_MS = 600;

/** How often a running clock redraws. */
const TICK_MS = 100;

/** The current action's data, fully determined by its seed. */
interface Rep {
  seed: number;
  data: unknown;
}

/** What the results screen shows: the run as recorded, without the progress it was recorded in. */
type RunOutcome = Omit<RunResult, 'state'>;

const slideIn = { from: { opacity: 0, transform: 'translateY(6px)' }, to: { opacity: 1, transform: 'translateY(0)' } };
const pop = { '0%': { transform: 'scale(0.6)' }, '60%': { transform: 'scale(1.18)' }, '100%': { transform: 'scale(1)' } };
const reduce = { '@media (prefers-reduced-motion: reduce)': { animationDuration: '1ms' } };

const useStyles = makeStyles({
  view: { display: 'flex', flexDirection: 'column', minHeight: '100%' },
  content: { flex: 1, display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalL, padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalL} ${tokens.spacingVerticalXL}` },
  back: { alignSelf: 'flex-start', marginLeft: `calc(-1 * ${tokens.spacingHorizontalS})` },
  head: { display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalXS },
  heading: { margin: 0, ':focus-visible': { outline: `2px solid ${tokens.colorStrokeFocus2}`, outlineOffset: '2px', borderRadius: tokens.borderRadiusSmall } },
  muted: { color: tokens.colorNeutralForeground3 },
  eyebrow: { color: tokens.colorBrandForeground1, fontWeight: tokens.fontWeightSemibold },
  tabular: { fontVariantNumeric: 'tabular-nums' },

  // overview
  sets: { listStyle: 'none', margin: 0, padding: 0, display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalS },
  set: {
    display: 'grid',
    gridTemplateColumns: '20px minmax(0, 1fr) 16px',
    alignItems: 'center',
    gap: tokens.spacingHorizontalM,
    width: '100%',
    padding: `${tokens.spacingVerticalM} ${tokens.spacingHorizontalM}`,
    borderRadius: tokens.borderRadiusLarge,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    background: tokens.colorNeutralBackground1,
    color: tokens.colorNeutralForeground1,
    textAlign: 'left',
    cursor: 'pointer',
    fontFamily: 'inherit',
    ':hover': { backgroundColor: tokens.colorSubtleBackgroundHover },
    ':active': { backgroundColor: tokens.colorSubtleBackgroundPressed },
    ':focus-visible': { outline: `2px solid ${tokens.colorStrokeFocus2}`, outlineOffset: '2px' },
  },
  setIcon: { color: tokens.colorBrandForeground1, display: 'inline-flex', alignSelf: 'start', marginTop: '1px' },
  setText: { display: 'flex', flexDirection: 'column', gap: '2px', minWidth: 0 },
  setMeta: {
    display: 'flex',
    flexWrap: 'wrap',
    alignItems: 'center',
    columnGap: tokens.spacingHorizontalS,
    rowGap: tokens.spacingVerticalXXS,
    marginTop: tokens.spacingVerticalXS,
    color: tokens.colorNeutralForeground3,
    fontVariantNumeric: 'tabular-nums',
  },
  chevron: { color: tokens.colorNeutralForeground4, display: 'inline-flex' },

  // run
  runHead: { display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: tokens.spacingHorizontalS },
  runStats: { display: 'inline-flex', alignItems: 'center', gap: tokens.spacingHorizontalS, color: tokens.colorNeutralForeground2, fontVariantNumeric: 'tabular-nums' },
  progress: { marginTop: `calc(-1 * ${tokens.spacingVerticalS})` },
  card: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalM,
    padding: `${tokens.spacingVerticalL} ${tokens.spacingHorizontalL}`,
    borderRadius: tokens.borderRadiusLarge,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    backgroundColor: tokens.colorNeutralBackground1,
    boxShadow: tokens.shadow4,
  },
  promptIn: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalXS,
    animationName: slideIn,
    animationDuration: '220ms',
    animationTimingFunction: tokens.curveDecelerateMid,
    ...reduce,
  },
  prompt: { fontSize: tokens.fontSizeBase500, lineHeight: tokens.lineHeightBase500, fontWeight: tokens.fontWeightSemibold, overflowWrap: 'anywhere' },
  clockRow: { display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: tokens.spacingHorizontalS, minHeight: '20px' },
  clock: { display: 'inline-flex', alignItems: 'center', gap: tokens.spacingHorizontalXS, color: tokens.colorNeutralForeground2, fontVariantNumeric: 'tabular-nums' },
  clockTime: { fontSize: tokens.fontSizeBase400, lineHeight: tokens.lineHeightBase400, fontWeight: tokens.fontWeightSemibold, color: tokens.colorNeutralForeground1 },
  how: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalS,
    padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalM}`,
    borderRadius: tokens.borderRadiusMedium,
    backgroundColor: tokens.colorNeutralBackground2,
  },
  steps: { margin: 0, paddingLeft: tokens.spacingHorizontalL, display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalXS },
  keysLine: { display: 'flex', flexWrap: 'wrap', alignItems: 'baseline', gap: tokens.spacingHorizontalXS },
  keys: {
    fontFamily: tokens.fontFamilyBase,
    fontSize: tokens.fontSizeBase200,
    lineHeight: tokens.lineHeightBase200,
    padding: `1px ${tokens.spacingHorizontalSNudge}`,
    borderRadius: tokens.borderRadiusMedium,
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    boxShadow: `inset 0 -1px 0 ${tokens.colorNeutralStroke1}`,
    backgroundColor: tokens.colorNeutralBackground1,
    color: tokens.colorNeutralForeground1,
    overflowWrap: 'anywhere',
  },
  passed: {
    display: 'grid',
    gridTemplateColumns: '20px minmax(0, 1fr)',
    gap: tokens.spacingHorizontalS,
    alignItems: 'start',
    padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalM}`,
    borderRadius: tokens.borderRadiusMedium,
    border: `1px solid ${tokens.colorPaletteGreenBorder1}`,
    backgroundColor: tokens.colorPaletteGreenBackground1,
  },
  passedIcon: { display: 'inline-flex', color: tokens.colorPaletteGreenForeground1 },
  passedFresh: { animationName: pop, animationDuration: '360ms', animationTimingFunction: 'cubic-bezier(0.2, 0.9, 0.3, 1.3)', ...reduce },
  passedBody: { display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalXXS, minWidth: 0 },
  passedHead: { display: 'flex', justifyContent: 'space-between', gap: tokens.spacingHorizontalS },

  // results
  summary: {
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    textAlign: 'center',
    gap: tokens.spacingVerticalXS,
    padding: `${tokens.spacingVerticalXL} ${tokens.spacingHorizontalM}`,
    borderRadius: tokens.borderRadiusLarge,
    backgroundColor: tokens.colorNeutralBackground2,
    animationName: slideIn,
    animationDuration: '260ms',
    animationTimingFunction: tokens.curveDecelerateMid,
    ...reduce,
  },
  summaryIcon: { display: 'inline-flex', color: tokens.colorBrandForeground1, marginBottom: tokens.spacingVerticalXS },
  bigTime: { fontSize: tokens.fontSizeHero800, lineHeight: tokens.lineHeightHero800, fontWeight: tokens.fontWeightSemibold, fontVariantNumeric: 'tabular-nums' },
  results: { listStyle: 'none', margin: 0, padding: 0, display: 'flex', flexDirection: 'column' },
  result: {
    display: 'grid',
    gridTemplateColumns: '20px minmax(0, 1fr) auto',
    gap: tokens.spacingHorizontalS,
    alignItems: 'start',
    padding: `${tokens.spacingVerticalS} 0`,
    borderBottom: `1px solid ${tokens.colorNeutralStroke3}`,
  },
  resultNum: { color: tokens.colorNeutralForeground3, fontVariantNumeric: 'tabular-nums', paddingTop: '1px' },
  resultText: { display: 'flex', flexDirection: 'column', alignItems: 'flex-start', gap: tokens.spacingVerticalXS, minWidth: 0 },
  resultTime: { display: 'flex', flexDirection: 'column', alignItems: 'flex-end', gap: tokens.spacingVerticalXXS, fontVariantNumeric: 'tabular-nums' },

  footer: {
    position: 'sticky',
    bottom: 0,
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalXS,
    padding: `${tokens.spacingVerticalM} ${tokens.spacingHorizontalL}`,
    borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
    backgroundColor: tokens.colorNeutralBackground1,
  },
  actions: { display: 'flex', gap: tokens.spacingHorizontalS, alignItems: 'center' },
  grow: { flex: 1 },
});

/** A shortcut or menu route, set like a key cap. */
function Keys({ children }: { children: ReactNode }) {
  const s = useStyles();
  return <kbd className={s.keys}>{children}</kbd>;
}

/**
 * A stopwatch reading: `ms`, plus the time on `clock` when there is one. Ticks on its own so the
 * rest of the view doesn't re-render ten times a second.
 */
function LiveTime({ ms, clock, className, label }: { ms: number; clock?: DrillClock; className?: string; label: string }) {
  const since = clock && runningSince(clock);
  const [now, setNow] = useState(() => Date.now());
  // Read the time again before paint whenever the clock starts, so the first frame isn't measured
  // against the moment it last stopped.
  useLayoutEffect(() => {
    if (since === undefined) return;
    setNow(Date.now());
    const t = window.setInterval(() => setNow(Date.now()), TICK_MS);
    return () => window.clearInterval(t);
  }, [since]);
  return (
    <span className={className} role="timer" aria-label={label}>
      {formatDrillTime(ms + (clock ? elapsedMs(clock, now) : 0))}
    </span>
  );
}

/** Timed action drills: a run of one-action practice sheets against the clock. The result is checked, never the keystrokes. */
export function DrillView() {
  const s = useStyles();
  const { host, hostNotice, platform, progress, setProgress, localize, goHome } = useCoach();
  const howId = useId('drill-how');

  const [phase, setPhase] = useState<Phase>('overview');
  const [drill, setDrill] = useState<Drill>();
  const [index, setIndex] = useState(0);
  const [rep, setRep] = useState<Rep>();
  const [clock, setClock] = useState<DrillClock>(() => startClock(0));
  const [times, setTimes] = useState<ItemTime[]>([]);
  const [report, setReport] = useState<CheckReport>();
  const [lastPass, setLastPass] = useState<{ index: number; ms: number }>();
  const [outcome, setOutcome] = useState<RunOutcome>();
  const [error, setError] = useState<string>();
  const [showHow, setShowHow] = useState(false);
  const [confirmLeave, setConfirmLeave] = useState(false);

  // Each run (and each exit) takes a new token, so a setup or check that finishes after the
  // learner has moved on is ignored.
  const runToken = useRef(0);
  const advanceTimer = useRef<number | undefined>(undefined);
  // When a check finishes, the footer's buttons change under the pointer. Ignore clicks for a
  // moment so a quick double-click on Done can't also press the next button.
  const settledAt = useRef(0);
  const settled = () => Date.now() - settledAt.current > 700;
  // Set while a check runs, so Done can't start a second one before the view re-renders.
  const checking = useRef(false);
  const progressRef = useRef(progress);
  progressRef.current = progress;
  // Where keyboard focus goes after the screen changes under it.
  const focusNext = useRef<'overview' | 'prompt' | 'summary' | null>(null);
  const overviewHeading = useRef<HTMLHeadingElement>(null);
  const promptHeading = useRef<HTMLHeadingElement>(null);
  const summaryHeading = useRef<HTMLHeadingElement>(null);

  const mac = platform === 'mac';
  const keysFor = (item: DrillItem) => (mac ? item.shortcut.mac : item.shortcut.windows);
  const sheetFor = (d: Drill) => sheetNameFor(`drill-${d.id}`);

  useEffect(() => {
    const target = focusNext.current;
    if (!target) return;
    const el = { overview: overviewHeading, prompt: promptHeading, summary: summaryHeading }[target].current;
    if (el) {
      el.focus();
      focusNext.current = null;
    }
  });

  useEffect(
    () => () => {
      runToken.current++;
      window.clearTimeout(advanceTimer.current);
    },
    [],
  );

  const setUpItem = useCallback(
    async (d: Drill, i: number) => {
      if (!host) return setError(hostNotice);
      const token = runToken.current;
      const item = d.items[i];
      const seed = newSeed();
      const data = item.exercise.make(new Rng(seed));
      setIndex(i);
      setRep({ seed, data });
      setReport(undefined);
      setShowHow(false);
      setError(undefined);
      setPhase('settingUp');
      try {
        await host.setup(item.exercise, data, sheetFor(d));
      } catch (e) {
        if (token !== runToken.current) return;
        setError(friendlyError(e));
        setPhase('setupFailed');
        return;
      }
      if (token !== runToken.current) return;
      setClock(startClock(Date.now()));
      settledAt.current = Date.now();
      setPhase('working');
    },
    [host, hostNotice],
  );

  const start = (d: Drill) => {
    if (!host) return setError(hostNotice);
    runToken.current++;
    checking.current = false;
    window.clearTimeout(advanceTimer.current);
    setConfirmLeave(false);
    setDrill(d);
    setTimes([]);
    setLastPass(undefined);
    setOutcome(undefined);
    focusNext.current = 'prompt';
    void setUpItem(d, 0);
  };

  const finish = useCallback(
    (d: Drill, finished: ItemTime[]) => {
      const now = Date.now();
      const o = recordRun(progressRef.current, d, finished, now);
      setProgress((prev) => recordRun(prev, d, finished, now).state);
      setOutcome({ totalMs: o.totalMs, skipped: o.skipped, personalBest: o.personalBest, newItemBests: o.newItemBests, previousBestMs: o.previousBestMs });
      settledAt.current = Date.now();
      focusNext.current = 'summary';
      setPhase('complete');
    },
    [setProgress],
  );

  const done = useCallback(async () => {
    if (!host || !drill || !rep || phase !== 'working' || checking.current) return;
    checking.current = true;
    const token = runToken.current;
    const item = drill.items[index];
    // The clock stops at Done. The check's own time doesn't count against the learner.
    const checkAt = Date.now();
    const paused = pauseClock(clock, checkAt);
    const resume = () => setClock((c) => resumeClock(c, Date.now()));
    setError(undefined);
    setClock(paused);
    setPhase('checking');

    let r: CheckReport;
    try {
      r = await host.check(item.exercise, rep.data, sheetFor(drill), rep.seed);
    } catch (e) {
      checking.current = false;
      if (token !== runToken.current) return;
      resume();
      setError(friendlyError(e));
      setPhase(e instanceof CoachError && e.code === 'missing-sheet' ? 'setupFailed' : 'working');
      return;
    }
    checking.current = false;
    if (token !== runToken.current) return;
    settledAt.current = Date.now();

    if (uncheckable(r)) {
      // Nothing was checked, so it isn't a pass, and there's no fair time to record.
      setReport(r);
      setShowHow(false);
      setPhase('uncheckable');
      return;
    }

    if (!r.passed) {
      resume();
      setReport(r);
      setPhase('working');
      return;
    }

    const ms = elapsedMs(paused, checkAt);
    const finished = [...times, ms];
    setTimes(finished);
    setLastPass({ index, ms });
    setReport(undefined);
    setShowHow(false);
    setPhase('passed');
    advanceTimer.current = window.setTimeout(() => {
      if (token !== runToken.current) return;
      const next = nextAction(drill, finished);
      if (next !== null) void setUpItem(drill, next);
      else finish(drill, finished);
    }, ADVANCE_MS);
  }, [host, drill, rep, phase, index, clock, times, setUpItem, finish]);

  // An action this Excel can't check goes without a time, and the run moves on.
  const skip = () => {
    if (!drill || phase !== 'uncheckable') return;
    const finished = [...times, null];
    setTimes(finished);
    setLastPass(undefined);
    setReport(undefined);
    const next = nextAction(drill, finished);
    if (next !== null) void setUpItem(drill, next);
    else finish(drill, finished);
  };

  // ⌘↩ or Ctrl+Enter selects Done from anywhere in the panel, but not from behind the leave dialog.
  const doneRef = useRef(done);
  doneRef.current = done;
  useEffect(() => {
    if (phase !== 'working' || confirmLeave) return;
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Enter' && (e.metaKey || e.ctrlKey)) {
        e.preventDefault();
        void doneRef.current();
      }
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, [phase, confirmLeave]);

  const leave = () => {
    runToken.current++;
    checking.current = false;
    window.clearTimeout(advanceTimer.current);
    setConfirmLeave(false);
    setDrill(undefined);
    setRep(undefined);
    setReport(undefined);
    setError(undefined);
    focusNext.current = 'overview';
    setPhase('overview');
  };

  // Leaving loses the run, so ask once the learner has something to lose.
  const invested = times.length > 0 || phase === 'working' || phase === 'checking' || phase === 'passed' || phase === 'uncheckable';
  const requestLeave = () => (invested && phase !== 'complete' ? setConfirmLeave(true) : leave());

  const goTo = (address: string) => {
    if (host && drill) host.select(sheetFor(drill), address).catch((e) => setError(friendlyError(e)));
  };

  // The current action's clock: live while working, stopped at Done while a check runs. Once the
  // action passes, its time is in `times`.
  const timed = phase === 'working' || phase === 'checking' ? clock : undefined;
  const doneMs = runTotalMs(times);

  const errorBar = error && (
    <MessageBar intent="error" layout="multiline">
      <MessageBarBody>{error}</MessageBarBody>
      <MessageBarActions containerAction={<Button appearance="transparent" size="small" icon={<Dismiss16Regular />} aria-label="Dismiss" onClick={() => setError(undefined)} />} />
    </MessageBar>
  );

  const noticeBar = hostNotice && (
    <MessageBar intent="info">
      <MessageBarBody>{hostNotice}</MessageBarBody>
    </MessageBar>
  );

  // ---------- overview ----------

  if (phase === 'overview' || !drill || !rep) {
    return (
      <div className={s.view}>
        <div className={s.content}>
          <Button className={s.back} appearance="transparent" size="small" icon={<ArrowLeft16Regular />} onClick={goHome}>
            All skills
          </Button>

          <div className={s.head}>
            <Text as="h1" size={500} weight="semibold" className={s.heading} ref={overviewHeading} tabIndex={-1}>
              Timed drills
            </Text>
            <Body1 className={s.muted}>Do each action in Excel as fast as you can, then select Done. The coach checks the result, not the keys you press.</Body1>
          </div>

          {noticeBar}
          {errorBar}

          <ul className={s.sets} aria-label="Drill sets">
            {DRILLS.map((d) => {
              const p = progress.drills?.[d.id];
              const parMs = d.parSeconds * 1000;
              return (
                <li key={d.id}>
                  <button className={s.set} onClick={() => start(d)}>
                    <span className={s.setIcon} aria-hidden="true">
                      <Timer20Regular />
                    </span>
                    <span className={s.setText}>
                      <Text weight="semibold">{d.title}</Text>
                      <Caption1 className={s.muted}>{d.blurb}</Caption1>
                      <span className={s.setMeta}>
                        <Caption1>{d.items.length} actions</Caption1>
                        <Caption1 aria-hidden="true">·</Caption1>
                        <Caption1>Par {formatDuration(parMs)}</Caption1>
                        {p?.bestMs !== undefined && (
                          <>
                            <Caption1 aria-hidden="true">·</Caption1>
                            <Caption1>Best {formatDrillTime(p.bestMs)}</Caption1>
                            {p.bestMs <= parMs && (
                              <Badge appearance="tint" color="success" size="small">
                                Under par
                              </Badge>
                            )}
                          </>
                        )}
                      </span>
                    </span>
                    <span className={s.chevron} aria-hidden="true">
                      <ChevronRight16Regular />
                    </span>
                  </button>
                </li>
              );
            })}
          </ul>

          <Caption1 className={s.muted}>
            The clock starts when each practice sheet is ready and pauses while the coach checks it. Each action gets a fresh sheet, which replaces earlier practice
            sheets.
          </Caption1>
        </div>
      </div>
    );
  }

  const n = drill.items.length;
  const item = drill.items[index];
  const bestItemMs = progress.drills?.[drill.id]?.bestItemMs ?? {};

  // ---------- results ----------

  if (phase === 'complete' && outcome) {
    const parMs = drill.parSeconds * 1000;
    const parGap = outcome.totalMs - parMs;
    const underPar = parGap <= 0;
    const previous = outcome.previousBestMs;
    const firstRun = previous === undefined;
    // A run with skipped actions is short of a full one, so it's not compared with par or the best.
    const partial = outcome.skipped > 0;
    const headline = partial
      ? `This version of Excel couldn’t check ${outcome.skipped} of the ${n} actions, so this run can’t set a best time.`
      : previous === undefined
        ? 'First run. That’s the time to beat.'
        : outcome.personalBest
          ? `New personal best, ${formatDrillGap(outcome.totalMs - previous)} faster.`
          : `${formatDrillGap(outcome.totalMs - previous)} behind your best of ${formatDrillTime(previous)}.`;
    const parLine = partial
      ? `Par (${formatDuration(parMs)}) is for all ${n} actions.`
      : Math.abs(parGap) < 100
        ? `Right on par (${formatDuration(parMs)}).`
        : `${formatDrillGap(parGap)} ${underPar ? 'under' : 'over'} par (${formatDuration(parMs)}).`;
    const newBests = new Set(outcome.newItemBests);

    return (
      <div className={s.view}>
        <div className={s.content}>
          <Button className={s.back} appearance="transparent" size="small" icon={<ArrowLeft16Regular />} onClick={leave}>
            All drills
          </Button>

          <section className={s.summary} aria-label="Result">
            <span className={s.summaryIcon} aria-hidden="true">
              {!partial && ((outcome.personalBest && !firstRun) || underPar) ? <Trophy24Filled /> : <CheckmarkCircle24Filled />}
            </span>
            <Caption1 className={s.eyebrow}>{drill.title}</Caption1>
            <Text as="h1" size={400} weight="semibold" className={s.heading} ref={summaryHeading} tabIndex={-1}>
              Drill complete
            </Text>
            <Text className={s.bigTime}>{formatDrillTime(outcome.totalMs)}</Text>
            <Body1>{headline}</Body1>
            <Caption1 className={s.muted}>{parLine}</Caption1>
          </section>

          <section aria-label="Times per action">
            <SectionLabel>Your times</SectionLabel>
            <ol className={s.results}>
              {drill.items.map((it, i) => {
                const ms = times[i];
                return (
                  <li key={drillItemId(it)} className={s.result}>
                    <Caption1 className={s.resultNum} aria-hidden="true">
                      {i + 1}
                    </Caption1>
                    <div className={s.resultText}>
                      <Text>{it.exercise.title}</Text>
                      <Keys>{keysFor(it)}</Keys>
                    </div>
                    <div className={s.resultTime}>
                      {typeof ms === 'number' ? <Text weight="semibold">{formatDrillTime(ms)}</Text> : <Caption1 className={s.muted}>Not checked</Caption1>}
                      {newBests.has(drillItemId(it)) && (
                        <Badge appearance="tint" color="success" size="small">
                          New best
                        </Badge>
                      )}
                    </div>
                  </li>
                );
              })}
            </ol>
          </section>
        </div>

        <footer className={s.footer}>
          <div className={s.actions}>
            <Button className={s.grow} appearance="primary" icon={<ArrowSync20Regular />} onClick={() => settled() && start(drill)}>
              Run again
            </Button>
            <Button onClick={() => settled() && leave()}>All drills</Button>
          </div>
        </footer>
      </div>
    );
  }

  // ---------- run ----------

  const passedItem = lastPass ? drill.items[lastPass.index] : undefined;
  const best = bestItemMs[drillItemId(item)];
  const skipping = phase === 'uncheckable';
  const cantReport = skipping ? (report?.items.find((i) => i.detail)?.detail ?? 'This version of Excel can’t report this change to add-ins.') : undefined;

  return (
    <div className={s.view}>
      <div className={s.content}>
        <div className={s.runHead}>
          <Button className={s.back} appearance="transparent" size="small" icon={<ArrowLeft16Regular />} onClick={requestLeave}>
            All drills
          </Button>
          <span className={s.runStats}>
            <Caption1>
              {index + 1} of {n}
            </Caption1>
            <Caption1 aria-hidden="true">·</Caption1>
            <Caption1>
              <LiveTime ms={doneMs} clock={timed} label="Total time" />
            </Caption1>
          </span>
        </div>
        <ProgressBar className={s.progress} value={times.length / n} thickness="medium" aria-label={`${times.length} of ${n} actions done`} />

        {noticeBar}

        <section className={s.card} aria-label="Current action">
          <div aria-live="polite" aria-atomic="true">
            <div key={index} className={s.promptIn}>
              <Caption1 className={s.eyebrow}>
                {drill.title} · {item.exercise.title}
              </Caption1>
              <Text as="h1" className={mergeClasses(s.heading, s.prompt)} ref={promptHeading} tabIndex={-1}>
                <RichText text={localize(item.exercise.task(rep.data))} />
              </Text>
            </div>
          </div>

          <div className={s.clockRow}>
            {phase === 'settingUp' ? (
              <Spinner size="extra-tiny" label="Setting up the sheet…" />
            ) : phase === 'setupFailed' ? (
              <Caption1 className={s.muted}>The sheet isn’t ready.</Caption1>
            ) : skipping ? (
              <Caption1 className={s.muted}>This action isn’t timed.</Caption1>
            ) : (
              <span className={s.clock}>
                <Timer16Regular aria-hidden="true" />
                <LiveTime className={s.clockTime} ms={phase === 'passed' && lastPass ? lastPass.ms : 0} clock={timed} label="Time on this action" />
                {best !== undefined && <Caption1 className={s.muted}>Best {formatDrillTime(best)}</Caption1>}
              </span>
            )}
            {(phase === 'working' || phase === 'checking') && (
              <Link as="button" inline={false} onClick={() => setShowHow((v) => !v)} aria-expanded={showHow} aria-controls={howId} style={{ fontSize: tokens.fontSizeBase200 }}>
                {showHow ? 'Hide how' : 'Show how'}
              </Link>
            )}
          </div>

          {showHow && (phase === 'working' || phase === 'checking') && (
            <div id={howId} className={s.how}>
              <ol className={s.steps}>
                {item.exercise.hints.map((h, i) => (
                  <li key={i}>
                    <Caption1>
                      <RichText text={localize(h)} />
                    </Caption1>
                  </li>
                ))}
              </ol>
              <span className={s.keysLine}>
                <Caption1 className={s.muted}>Shortcut</Caption1>
                <Keys>{keysFor(item)}</Keys>
              </span>
            </div>
          )}
        </section>

        {errorBar}

        {phase === 'working' && report && <CheckResults report={report} onGoTo={goTo} />}

        <div role="status">
          {skipping ? (
            <MessageBar intent="warning" layout="multiline">
              <MessageBarBody>
                <MessageBarTitle>Excel can’t check this action</MessageBarTitle>
                {cantReport} Skip it to go on. It won’t count toward your time.
              </MessageBarBody>
            </MessageBar>
          ) : (
            lastPass &&
            passedItem && (
              <div key={lastPass.index} className={s.passed}>
                <span className={mergeClasses(s.passedIcon, phase === 'passed' && lastPass.index === index && s.passedFresh)} aria-hidden="true">
                  <CheckmarkCircle20Filled />
                </span>
                <div className={s.passedBody}>
                  <div className={s.passedHead}>
                    <Text weight="semibold">{passedItem.exercise.title}</Text>
                    <Text className={s.tabular}>{formatDrillTime(lastPass.ms)}</Text>
                  </div>
                  <span className={s.keysLine}>
                    <Caption1>Shortcut</Caption1>
                    <Keys>{keysFor(passedItem)}</Keys>
                  </span>
                </div>
              </div>
            )
          )}
        </div>
      </div>

      <footer className={s.footer}>
        {phase === 'setupFailed' ? (
          <Button appearance="primary" size="large" icon={<ArrowSync20Regular />} onClick={() => settled() && void setUpItem(drill, index)}>
            Set up again
          </Button>
        ) : (
          // One button that turns into Skip, so keyboard focus stays on it when a check can't run.
          <Tooltip content={skipping ? 'Go on without a time for this action' : `Check the sheet (${mac ? '⌘↩' : 'Ctrl+Enter'})`} relationship="description">
            <Button
              appearance="primary"
              size="large"
              icon={
                phase === 'settingUp' || phase === 'checking' ? <Spinner size="tiny" /> : phase === 'passed' ? <CheckmarkCircle20Filled /> : skipping ? <Next20Regular /> : undefined
              }
              disabledFocusable={phase !== 'working' && !skipping}
              aria-keyshortcuts={skipping ? undefined : mac ? 'Meta+Enter' : 'Control+Enter'}
              onClick={() => {
                if (!settled()) return;
                if (skipping) skip();
                else void done();
              }}
            >
              {phase === 'settingUp' ? 'Setting up…' : phase === 'checking' ? 'Checking…' : skipping ? 'Skip this action' : 'Done'}
            </Button>
          </Tooltip>
        )}
        <Caption1 className={s.muted}>
          {phase === 'setupFailed'
            ? 'Sets up this action’s sheet again. Its clock starts over.'
            : skipping
              ? index + 1 < n
                ? 'Sets up the next action.'
                : 'Shows the results of this run.'
              : 'Select Done when the sheet looks right.'}
        </Caption1>
      </footer>

      <ConfirmDialog
        open={confirmLeave}
        title="Leave this drill?"
        body="The run stops here, and its times aren’t saved."
        confirmLabel="Leave"
        cancelLabel="Keep going"
        onCancel={() => setConfirmLeave(false)}
        onConfirm={leave}
      />
    </div>
  );
}

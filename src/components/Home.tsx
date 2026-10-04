import {
  Accordion,
  AccordionHeader,
  AccordionItem,
  AccordionPanel,
  Badge,
  Button,
  Caption1,
  ProgressBar,
  Text,
  makeStyles,
  mergeClasses,
  tokens,
  type AccordionToggleEventHandler,
} from '@fluentui/react-components';
import { ArrowRight16Regular, ChevronRight16Regular, Clock16Regular, DocumentSearch20Regular, Mail16Regular, Timer20Regular } from '@fluentui/react-icons';
import { useEffect, useMemo, useState } from 'react';
import { DRILLS } from '../drills';
import { MASTERY_PASSES, missionProgressFor, progressFor, statusOf, weekSummary, type SkillStatus } from '../engine/progress';
import type { Exercise } from '../engine/types';
import { EXERCISES, MODULES, exercisesIn } from '../exercises';
import { MISSIONS } from '../missions';
import { useCoach } from '../taskpane/context';
import { SectionLabel, StatusIcon } from './bits';
import { WeekSummary, hasProgress } from './WeekSummary';

const useStyles = makeStyles({
  page: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalXL,
    padding: `${tokens.spacingVerticalL} ${tokens.spacingHorizontalL} ${tokens.spacingVerticalXXL}`,
  },
  summary: { display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalS },
  summaryTop: { display: 'flex', alignItems: 'baseline', justifyContent: 'space-between', gap: tokens.spacingHorizontalS },
  steps: { display: 'grid', gap: tokens.spacingVerticalS, margin: 0, padding: 0, listStyle: 'none' },
  step: { display: 'grid', gridTemplateColumns: '22px minmax(0, 1fr)', gap: tokens.spacingHorizontalS, alignItems: 'start' },
  stepNum: {
    width: '22px',
    height: '22px',
    borderRadius: '50%',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    fontSize: tokens.fontSizeBase200,
    fontWeight: tokens.fontWeightSemibold,
    backgroundColor: tokens.colorBrandBackground2,
    color: tokens.colorBrandForeground2,
  },
  next: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalXS,
    padding: `${tokens.spacingVerticalM} ${tokens.spacingHorizontalM}`,
    borderRadius: tokens.borderRadiusLarge,
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    backgroundColor: tokens.colorNeutralBackground1,
    boxShadow: tokens.shadow2,
  },
  eyebrow: { color: tokens.colorBrandForeground1, fontWeight: tokens.fontWeightSemibold },
  muted: { color: tokens.colorNeutralForeground3 },
  meta: { display: 'flex', alignItems: 'center', gap: tokens.spacingHorizontalXS, color: tokens.colorNeutralForeground3 },
  nextActions: { display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginTop: tokens.spacingVerticalS },
  skills: { display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalXS },
  moduleItem: { borderBottom: `1px solid ${tokens.colorNeutralStroke3}` },
  moduleButton: { paddingLeft: 0, paddingRight: 0 },
  modulePanel: { margin: 0 },
  moduleHeader: { display: 'flex', alignItems: 'center', justifyContent: 'space-between', width: '100%', gap: tokens.spacingHorizontalS, fontWeight: tokens.fontWeightSemibold },
  moduleCount: { color: tokens.colorNeutralForeground3, fontVariantNumeric: 'tabular-nums' },
  moduleDone: { color: tokens.colorBrandForeground1 },
  rows: { display: 'flex', flexDirection: 'column', paddingBottom: tokens.spacingVerticalS },
  row: {
    display: 'grid',
    gridTemplateColumns: '20px minmax(0, 1fr) auto 16px',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
    width: `calc(100% + 2 * ${tokens.spacingHorizontalS})`,
    minHeight: '36px',
    margin: `0 calc(-1 * ${tokens.spacingHorizontalS})`,
    padding: `${tokens.spacingVerticalXS} ${tokens.spacingHorizontalS}`,
    border: 'none',
    borderRadius: tokens.borderRadiusMedium,
    background: 'none',
    color: tokens.colorNeutralForeground1,
    textAlign: 'left',
    cursor: 'pointer',
    fontFamily: 'inherit',
    fontSize: tokens.fontSizeBase300,
    ':hover': { backgroundColor: tokens.colorSubtleBackgroundHover },
    ':active': { backgroundColor: tokens.colorSubtleBackgroundPressed },
    ':focus-visible': { outline: `2px solid ${tokens.colorStrokeFocus2}`, outlineOffset: '-2px' },
  },
  rowTitle: { overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' },
  tools: { display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalS },
  tool: {
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
    ':focus-visible': { outline: `2px solid ${tokens.colorStrokeFocus2}`, outlineOffset: '2px' },
  },
  toolIcon: { color: tokens.colorBrandForeground1, display: 'inline-flex' },
  toolText: { display: 'flex', flexDirection: 'column', gap: '2px', minWidth: 0 },
  missionRow: {
    display: 'grid',
    gridTemplateColumns: '20px minmax(0, 1fr) auto 16px',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
    width: `calc(100% + 2 * ${tokens.spacingHorizontalS})`,
    margin: `0 calc(-1 * ${tokens.spacingHorizontalS})`,
    minHeight: '48px',
    padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalS}`,
    border: 'none',
    borderRadius: tokens.borderRadiusMedium,
    background: 'none',
    color: tokens.colorNeutralForeground1,
    textAlign: 'left',
    cursor: 'pointer',
    fontFamily: 'inherit',
    ':hover': { backgroundColor: tokens.colorSubtleBackgroundHover },
    ':focus-visible': { outline: `2px solid ${tokens.colorStrokeFocus2}`, outlineOffset: '-2px' },
  },
  chevron: { color: tokens.colorNeutralForeground4, display: 'inline-flex' },
});

interface Row {
  ex: Exercise<any>;
  status: SkillStatus;
  passes: number;
}

function pickNext(rows: Row[], resumeId?: string): Row | undefined {
  return (
    rows.find((r) => r.ex.id === resumeId && r.status !== 'mastered') ??
    rows.find((r) => r.status === 'review') ??
    rows.find((r) => r.status === 'practicing') ??
    rows.find((r) => r.status === 'new')
  );
}

/**
 * The time, refreshed each minute and whenever the pane comes back into view. A pane left
 * open on Home overnight then moves "today" along and picks up reviews as they fall due.
 */
function useClock(): number {
  const [now, setNow] = useState(() => Date.now());
  useEffect(() => {
    const tick = () => setNow(Date.now());
    const onVisibility = () => {
      if (document.visibilityState === 'visible') tick();
    };
    const t = setInterval(tick, 60_000);
    document.addEventListener('visibilitychange', onVisibility);
    return () => {
      clearInterval(t);
      document.removeEventListener('visibilitychange', onVisibility);
    };
  }, []);
  return now;
}

export function Home() {
  const s = useStyles();
  const { progress, openExercise, openMission, openScan, openDrills } = useCoach();
  const now = useClock();
  const rows: Row[] = useMemo(
    () =>
      EXERCISES.map((ex) => {
        const p = progressFor(progress, ex.id);
        return { ex, status: statusOf(p, now), passes: Math.min(p.passSeeds.length, MASTERY_PASSES) };
      }),
    [progress, now],
  );
  const week = useMemo(() => weekSummary(progress, EXERCISES.map((e) => e.id), now), [progress, now]);

  const inProgress = rows.filter((r) => r.status === 'practicing').length;
  const fresh = rows.every((r) => r.status === 'new');
  const session = progress.session && !progress.session.passedAt ? progress.session : undefined;
  const next = pickNext(rows, session?.exerciseId);
  const moduleTitle = (id: string) => MODULES.find((m) => m.id === id)?.title ?? '';

  const [open, setOpen] = useState<string[]>(() => (next ? [next.ex.module] : [MODULES[0].id]));
  const onToggle: AccordionToggleEventHandler<string> = (_, data) => setOpen(data.openItems as string[]);

  const actionLabel = (r: Row) =>
    r.ex.id === session?.exerciseId ? 'Resume' : r.status === 'review' ? 'Review' : r.status === 'practicing' ? 'Continue' : 'Start';

  const totalMinutes = EXERCISES.reduce((a, e) => a + e.minutes, 0);

  return (
    <div className={s.page}>
      {hasProgress(progress) ? (
        <WeekSummary summary={week} total={EXERCISES.length} />
      ) : (
        <section className={s.summary} aria-label="Progress">
          <div className={s.summaryTop}>
            <Text weight="semibold" size={400}>
              {fresh ? `${EXERCISES.length} advanced skills` : `${week.masteredTotal} of ${EXERCISES.length} skills mastered`}
            </Text>
            {fresh ? (
              <Caption1 className={s.muted}>About {Math.round(totalMinutes / 5) * 5} min</Caption1>
            ) : inProgress > 0 ? (
              <Caption1 className={s.muted}>{inProgress} in progress</Caption1>
            ) : null}
          </div>
          <ProgressBar value={week.masteredTotal / EXERCISES.length} thickness="large" aria-label="Skills mastered" />
        </section>
      )}

      {fresh && (
        <section aria-label="How it works" className={s.summary}>
          <SectionLabel>How it works</SectionLabel>
          <ol className={s.steps}>
            {[
              'The coach adds a practice sheet with fresh data.',
              'You solve it in Excel, the way you would at work.',
              'Check your work. The coach changes the data behind the scenes to make sure your formulas really hold up.',
            ].map((text, i) => (
              <li key={i} className={s.step}>
                <span className={s.stepNum} aria-hidden="true">
                  {i + 1}
                </span>
                <Caption1>{text}</Caption1>
              </li>
            ))}
          </ol>
        </section>
      )}

      {next && (
        <section className={s.next} aria-label="Up next">
          <Caption1 className={s.eyebrow}>
            {next.status === 'review' ? 'Review' : 'Up next'} · {moduleTitle(next.ex.module)}
          </Caption1>
          <Text weight="semibold" size={400}>
            {next.ex.title}
          </Text>
          <Caption1 className={s.muted}>{next.ex.replaces}</Caption1>
          <div className={s.nextActions}>
            <span className={s.meta}>
              <Clock16Regular aria-hidden="true" />
              <Caption1>About {next.ex.minutes} min</Caption1>
            </span>
            <Button appearance="primary" icon={<ArrowRight16Regular />} iconPosition="after" onClick={() => openExercise(next.ex.id)}>
              {actionLabel(next)}
            </Button>
          </div>
        </section>
      )}

      <div className={s.tools}>
        <button className={s.tool} onClick={openScan}>
          <span className={s.toolIcon} aria-hidden="true">
            <DocumentSearch20Regular />
          </span>
          <span className={s.toolText}>
            <Text weight="semibold">Scan your own workbook</Text>
            <Caption1 className={s.muted}>Find spots where these skills would save you time.</Caption1>
          </span>
          <span className={s.chevron} aria-hidden="true">
            <ChevronRight16Regular />
          </span>
        </button>
        {DRILLS.length > 0 && (
          <button className={s.tool} onClick={openDrills}>
            <span className={s.toolIcon} aria-hidden="true">
              <Timer20Regular />
            </span>
            <span className={s.toolText}>
              <Text weight="semibold">Timed drills</Text>
              <Caption1 className={s.muted}>Quick actions against the clock: freeze, filter, sort, format.</Caption1>
            </span>
            <span className={s.chevron} aria-hidden="true">
              <ChevronRight16Regular />
            </span>
          </button>
        )}
      </div>

      {MISSIONS.length > 0 && (
        <section className={s.skills} aria-label="Missions">
          <SectionLabel>Missions</SectionLabel>
          <div className={s.rows}>
            {MISSIONS.map((m) => {
              const mp = missionProgressFor(progress, m.id);
              const live = progress.session?.missionId === m.id && !progress.session.passedAt ? progress.session : undefined;
              return (
                <button key={m.id} className={s.missionRow} onClick={() => openMission(m.id)}>
                  <span className={s.toolIcon} aria-hidden="true">
                    <Mail16Regular />
                  </span>
                  <span className={s.toolText}>
                    <span className={s.rowTitle}>{m.title}</span>
                    <Caption1 className={s.muted}>
                      {live ? `In progress · ${live.stepsPassed?.length ?? 0} of ${m.steps.length} steps` : `${m.role === 'finance' ? 'Finance' : m.role === 'sales' ? 'Sales' : 'Ops'} · ${m.steps.length} steps · about ${m.minutes} min`}
                    </Caption1>
                  </span>
                  {mp.completions > 0 ? (
                    <Badge appearance="tint" color="success" size="small">
                      Done
                    </Badge>
                  ) : (
                    <span />
                  )}
                  <span className={s.chevron} aria-hidden="true">
                    <ChevronRight16Regular />
                  </span>
                </button>
              );
            })}
          </div>
        </section>
      )}

      <section className={s.skills} aria-label="All skills">
        <SectionLabel>All skills</SectionLabel>
        <Accordion multiple collapsible openItems={open} onToggle={onToggle}>
          {MODULES.filter((m) => exercisesIn(m.id).length > 0).map((m) => {
            const items = exercisesIn(m.id);
            const done = items.filter((e) => {
              const st = rows.find((r) => r.ex.id === e.id)!.status;
              return st === 'mastered' || st === 'review';
            }).length;
            return (
              <AccordionItem key={m.id} value={m.id} className={s.moduleItem}>
                <AccordionHeader expandIconPosition="end" size="medium" button={{ className: s.moduleButton }}>
                  <span className={s.moduleHeader}>
                    <span>{m.title}</span>
                    <Caption1 className={mergeClasses(s.moduleCount, done === items.length && s.moduleDone)}>
                      {done}/{items.length}
                    </Caption1>
                  </span>
                </AccordionHeader>
                <AccordionPanel className={s.modulePanel}>
                  <div className={s.rows}>
                    {items.map((ex) => {
                      const r = rows.find((x) => x.ex.id === ex.id)!;
                      return (
                        <button key={ex.id} className={s.row} onClick={() => openExercise(ex.id)}>
                          <StatusIcon status={r.status} />
                          <span className={s.rowTitle}>{ex.title}</span>
                          {r.status === 'review' ? (
                            <Badge appearance="tint" color="warning" size="small">
                              Review
                            </Badge>
                          ) : r.status === 'practicing' ? (
                            <Caption1 className={s.moduleCount}>
                              {r.passes}/{MASTERY_PASSES}
                            </Caption1>
                          ) : ex.m365 && r.status === 'new' ? (
                            <Badge appearance="outline" size="small" color="informative">
                              365
                            </Badge>
                          ) : (
                            <span />
                          )}
                          <span className={s.chevron} aria-hidden="true">
                            <ChevronRight16Regular />
                          </span>
                        </button>
                      );
                    })}
                  </div>
                </AccordionPanel>
              </AccordionItem>
            );
          })}
        </Accordion>
      </section>
    </div>
  );
}

import {
  Accordion,
  AccordionHeader,
  AccordionItem,
  AccordionPanel,
  Badge,
  Body1,
  Button,
  Caption1,
  Menu,
  MenuItem,
  MenuList,
  MenuPopover,
  MenuTrigger,
  MessageBar,
  MessageBarBody,
  Spinner,
  Text,
  makeStyles,
  mergeClasses,
  shorthands,
  tokens,
} from '@fluentui/react-components';
import {
  ArrowLeft16Regular,
  ArrowRight16Regular,
  ArrowSync20Regular,
  CheckmarkCircle20Filled,
  Eye20Regular,
  Lightbulb20Regular,
  Mail16Regular,
  MoreHorizontal20Regular,
  Table20Regular,
  Timer16Regular,
} from '@fluentui/react-icons';
import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { buildHintPayload, canExplain } from '../ai/hints';
import { formatDuration, missionProgressFor, recordStepPass, type StepOutcome } from '../engine/progress';
import { Rng, newSeed } from '../engine/rng';
import type { CheckReport } from '../engine/types';
import { CoachError, friendlyError } from '../excel/errors';
import { sheetNameFor } from '../excel/host';
import { getExercise } from '../exercises';
import { missionStepExercise } from '../missions/compile';
import { getMission } from '../missions';
import { useCoach } from '../taskpane/context';
import { RichText, SectionLabel, useNow } from './bits';
import { CheckResults } from './CheckResults';
import { ConfirmDialog } from './ConfirmDialog';
import { HintPanel, hintKey, useHintsEnabled } from './HintPanel';

type Phase = 'loading' | 'brief' | 'settingUp' | 'working' | 'checking' | 'stepPassed' | 'complete';

const useStyles = makeStyles({
  view: { display: 'flex', flexDirection: 'column', minHeight: '100%' },
  content: { flex: 1, display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalL, padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalL} ${tokens.spacingVerticalXL}` },
  back: { alignSelf: 'flex-start', marginLeft: `calc(-1 * ${tokens.spacingHorizontalS})` },
  titleBlock: { display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalXXS },
  eyebrow: { color: tokens.colorBrandForeground1, fontWeight: tokens.fontWeightSemibold },
  muted: { color: tokens.colorNeutralForeground3 },
  meta: { display: 'flex', alignItems: 'center', gap: tokens.spacingHorizontalM, marginTop: tokens.spacingVerticalXS, color: tokens.colorNeutralForeground3, fontVariantNumeric: 'tabular-nums' },
  metaItem: { display: 'inline-flex', alignItems: 'center', gap: tokens.spacingHorizontalXS },
  email: { borderRadius: tokens.borderRadiusLarge, border: `1px solid ${tokens.colorNeutralStroke2}`, overflow: 'hidden' },
  emailHead: { display: 'flex', flexDirection: 'column', gap: '2px', padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalM}`, backgroundColor: tokens.colorNeutralBackground2, borderBottom: `1px solid ${tokens.colorNeutralStroke2}` },
  emailFrom: { display: 'flex', alignItems: 'center', gap: tokens.spacingHorizontalXS, color: tokens.colorNeutralForeground3 },
  emailBody: { padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalM} ${tokens.spacingVerticalM}`, whiteSpace: 'pre-line' },
  steps: { listStyle: 'none', margin: 0, padding: 0, display: 'flex', flexDirection: 'column' },
  step: { display: 'grid', gridTemplateColumns: '24px minmax(0, 1fr)', gap: tokens.spacingHorizontalS, alignItems: 'start', padding: `${tokens.spacingVerticalS} 0`, borderBottom: `1px solid ${tokens.colorNeutralStroke3}` },
  stepNum: {
    width: '22px',
    height: '22px',
    borderRadius: '50%',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    fontSize: tokens.fontSizeBase200,
    fontWeight: tokens.fontWeightSemibold,
    border: `1.5px solid ${tokens.colorNeutralStroke1}`,
    color: tokens.colorNeutralForeground3,
    boxSizing: 'border-box',
  },
  stepNumCurrent: { ...shorthands.borderColor(tokens.colorBrandStroke1), color: tokens.colorBrandForeground1, backgroundColor: tokens.colorBrandBackground2 },
  stepDone: { color: tokens.colorBrandForeground1, display: 'inline-flex' },
  stepTitle: { paddingTop: '2px' },
  stepTitleCurrent: { fontWeight: tokens.fontWeightSemibold },
  stepBody: { display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalS, marginTop: tokens.spacingVerticalXS },
  hint: { display: 'grid', gridTemplateColumns: '20px minmax(0, 1fr)', gap: tokens.spacingHorizontalS, padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalM}`, borderRadius: tokens.borderRadiusMedium, backgroundColor: tokens.colorNeutralBackground2 },
  hintIcon: { color: tokens.colorPaletteMarigoldForeground1, display: 'inline-flex' },
  answer: { fontFamily: tokens.fontFamilyMonospace, fontSize: tokens.fontSizeBase200, padding: tokens.spacingHorizontalS, borderRadius: tokens.borderRadiusMedium, border: `1px dashed ${tokens.colorNeutralStroke1}`, overflowWrap: 'anywhere', whiteSpace: 'pre-wrap', margin: 0 },
  done: { display: 'flex', flexDirection: 'column', alignItems: 'center', textAlign: 'center', gap: tokens.spacingVerticalS, padding: `${tokens.spacingVerticalL} ${tokens.spacingHorizontalM}`, borderRadius: tokens.borderRadiusLarge, backgroundColor: tokens.colorNeutralBackground2 },
  // Scrolled into view after a check; the margin clears the sticky footer (about 65px) with room to spare.
  result: { display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalM, scrollMarginBottom: '88px' },
  stepOk: { display: 'flex', alignItems: 'center', gap: tokens.spacingHorizontalS, padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalM}`, borderRadius: tokens.borderRadiusMedium, backgroundColor: tokens.colorPaletteGreenBackground1, color: tokens.colorPaletteGreenForeground1 },
  footer: { position: 'sticky', bottom: 0, display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalXS, padding: `${tokens.spacingVerticalM} ${tokens.spacingHorizontalL}`, borderTop: `1px solid ${tokens.colorNeutralStroke2}`, backgroundColor: tokens.colorNeutralBackground1 },
  actions: { display: 'flex', gap: tokens.spacingHorizontalS, alignItems: 'center' },
  grow: { flex: 1 },
});

export function MissionView({ id }: { id: string }) {
  const s = useStyles();
  const { host, hostNotice, platform, progress, setProgress, localize, goHome, openExercise } = useCoach();
  const m = getMission(id)!;
  const mp = missionProgressFor(progress, id);
  const session = progress.session?.missionId === id ? progress.session : undefined;
  const step = session?.step ?? 0;
  const passed = new Set(session?.stepsPassed ?? []);

  const [phase, setPhase] = useState<Phase>(session && host ? (session.passedAt ? 'complete' : 'loading') : 'brief');
  const [report, setReport] = useState<CheckReport>();
  const [outcome, setOutcome] = useState<StepOutcome>();
  const [error, setError] = useState<string>();
  const [confirm, setConfirm] = useState<'restart' | 'reveal' | null>(null);
  const [advancing, setAdvancing] = useState(false);
  const settledAt = useRef(0);
  const settled = () => Date.now() - settledAt.current > 700;
  const resultRef = useRef<HTMLDivElement>(null);
  const hintsOn = useHintsEnabled();

  const data = useMemo(() => m.make(new Rng(session?.seed ?? 1)), [m, session?.seed]);
  const brief = m.brief(data);
  const stepEx = useMemo(() => missionStepExercise(m, step), [m, step]);
  const now = useNow(phase === 'working' || phase === 'checking' || phase === 'stepPassed');
  const elapsed = session ? (session.passedAt ?? now) - session.startedAt : 0;
  const busy = phase === 'loading' || phase === 'settingUp' || phase === 'checking';

  useEffect(() => {
    if (phase !== 'loading' || !host || !session) return;
    let cancelled = false;
    host
      .sheetExists(session.sheet)
      .then((exists) => !cancelled && setPhase(exists ? 'working' : 'brief'))
      .catch(() => !cancelled && setPhase('brief'));
    return () => {
      cancelled = true;
    };
  }, [phase, host, session]);

  useEffect(() => {
    if (!report && phase !== 'complete') return;
    resultRef.current?.scrollIntoView({ behavior: matchMedia('(prefers-reduced-motion: reduce)').matches ? 'auto' : 'smooth', block: 'nearest' });
  }, [report, phase]);

  const start = useCallback(async () => {
    if (!host) return setError(hostNotice);
    setConfirm(null);
    setError(undefined);
    setReport(undefined);
    setOutcome(undefined);
    setPhase('settingUp');
    const seed = newSeed();
    const sheet = sheetNameFor(m.id);
    try {
      const fresh = m.make(new Rng(seed));
      await host.setup(missionStepExercise(m, 0), fresh, sheet);
      const opening = m.steps[0].onStart?.(fresh) ?? [];
      if (opening.length) await host.writeInputs(sheet, opening);
      setProgress((prev) => ({
        ...prev,
        session: {
          exerciseId: `${m.id}#1`,
          missionId: m.id,
          step: 0,
          stepsPassed: [],
          startedSteps: opening.length ? [0] : [],
          seed,
          sheet,
          startedAt: Date.now(),
          attempts: 0,
          hintsShown: 0,
          revealed: false,
        },
      }));
      setPhase('working');
    } catch (e) {
      setError(friendlyError(e));
      setPhase('brief');
    }
  }, [host, hostNotice, m, setProgress]);

  const check = useCallback(async () => {
    if (!host || !session || !settled()) return;
    setError(undefined);
    setPhase('checking');
    try {
      const r = await host.check(stepEx, data, session.sheet, session.seed);
      setReport(r);
      settledAt.current = Date.now();
      if (r.passed) {
        const o = recordStepPass(progress, m.steps.length, Date.now());
        setOutcome(o);
        setProgress(() => o.state);
        setPhase(o.missionComplete ? 'complete' : 'stepPassed');
      } else {
        setProgress((prev) => (prev.session ? { ...prev, session: { ...prev.session, attempts: prev.session.attempts + 1 } } : prev));
        setPhase('working');
      }
    } catch (e) {
      setError(friendlyError(e));
      setPhase(e instanceof CoachError && e.code === 'missing-sheet' ? 'brief' : 'working');
    }
  }, [host, session, stepEx, data, progress, m, setProgress]);

  const nextStep = async () => {
    if (!settled() || !session) return;
    const next = m.steps.findIndex((_, i) => !passed.has(i) && i !== step);
    const target = next === -1 ? step : next;
    // Some steps change the data when they begin (new rows arrive); apply that once.
    const writes = session.startedSteps?.includes(target) ? [] : (m.steps[target].onStart?.(data) ?? []);
    if (writes.length) {
      if (!host) return setError(hostNotice);
      setAdvancing(true);
      try {
        await host.writeInputs(session.sheet, writes);
      } catch (e) {
        setError(friendlyError(e));
        return;
      } finally {
        setAdvancing(false);
      }
    }
    setProgress((prev) =>
      prev.session
        ? {
            ...prev,
            session: {
              ...prev.session,
              step: target,
              exerciseId: `${m.id}#${target + 1}`,
              hintsShown: 0,
              revealed: false,
              startedSteps: writes.length ? [...(prev.session.startedSteps ?? []), target] : prev.session.startedSteps,
            },
          }
        : prev,
    );
    setReport(undefined);
    setPhase('working');
  };

  const showHint = () => setProgress((prev) => (prev.session ? { ...prev, session: { ...prev.session, hintsShown: prev.session.hintsShown + 1 } } : prev));
  const reveal = () => {
    setConfirm(null);
    setProgress((prev) => (prev.session ? { ...prev, session: { ...prev.session, revealed: true } } : prev));
  };
  const goTo = (address: string) => host && session && host.select(session.sheet, address).catch((e) => setError(friendlyError(e)));

  const current = m.steps[step];
  const hintsShown = session?.hintsShown ?? 0;
  const started = !!session && phase !== 'brief' && phase !== 'settingUp' && phase !== 'loading';

  return (
    <div className={s.view}>
      <div className={s.content}>
        <Button className={s.back} appearance="transparent" size="small" icon={<ArrowLeft16Regular />} onClick={goHome}>
          All skills
        </Button>

        <div className={s.titleBlock}>
          <Caption1 className={s.eyebrow}>Mission · {m.role === 'finance' ? 'Finance' : m.role === 'sales' ? 'Sales' : 'Operations'}</Caption1>
          <Text as="h1" size={500} weight="semibold" style={{ margin: 0 }}>
            {m.title}
          </Text>
          <Caption1 className={s.muted}>{m.summary}</Caption1>
          <div className={s.meta}>
            <span className={s.metaItem}>
              <Timer16Regular aria-hidden="true" />
              <Caption1>{started ? formatDuration(elapsed) : `About ${m.minutes} min`}</Caption1>
            </span>
            <Caption1>
              {started ? `${passed.size} of ${m.steps.length} steps` : `${m.steps.length} steps`}
            </Caption1>
            {mp.bestMs !== undefined && <Caption1>Best {formatDuration(mp.bestMs)}</Caption1>}
            {mp.completions > 0 && (
              <Badge appearance="tint" color="success" size="small">
                Done {mp.completions}×
              </Badge>
            )}
          </div>
        </div>

        {hostNotice && (
          <MessageBar intent="info">
            <MessageBarBody>{hostNotice}</MessageBarBody>
          </MessageBar>
        )}

        <div className={s.email}>
          <Accordion collapsible defaultOpenItems={started ? [] : ['brief']}>
            <AccordionItem value="brief">
              <AccordionHeader expandIconPosition="end" size="medium" icon={<Mail16Regular />}>
                {brief.subject}
              </AccordionHeader>
              <AccordionPanel style={{ margin: 0 }}>
                <div className={s.emailHead}>
                  <Caption1 className={s.emailFrom}>From {brief.from}</Caption1>
                </div>
                <div className={s.emailBody}>
                  <Body1 style={{ whiteSpace: 'pre-line', display: 'block' }}>
                    <RichText text={brief.body.trim().replace(/[ \t]+\n/g, '\n').replace(/\n[ \t]+/g, '\n')} />
                  </Body1>
                </div>
              </AccordionPanel>
            </AccordionItem>
          </Accordion>
        </div>

        {!started && m.skills.length > 0 && (
          <Caption1 className={s.muted}>
            Uses:{' '}
            {m.skills
              .map((sid) => getExercise(sid))
              .filter(Boolean)
              .map((e, i, arr) => (
                <span key={e!.id}>
                  <button onClick={() => openExercise(e!.id)} style={{ background: 'none', border: 'none', padding: 0, color: tokens.colorBrandForegroundLink, cursor: 'pointer', font: 'inherit' }}>
                    {e!.title}
                  </button>
                  {i < arr.length - 1 ? ', ' : ''}
                </span>
              ))}
          </Caption1>
        )}

        <section aria-label="Steps">
          <SectionLabel>Steps</SectionLabel>
          <ol className={s.steps}>
            {m.steps.map((st, i) => {
              const isCurrent = started && i === step && phase !== 'complete';
              const isDone = passed.has(i);
              return (
                <li key={i} className={s.step}>
                  {isDone ? (
                    <span className={s.stepDone} role="img" aria-label="Done">
                      <CheckmarkCircle20Filled />
                    </span>
                  ) : (
                    <span className={mergeClasses(s.stepNum, isCurrent && s.stepNumCurrent)} aria-hidden="true">
                      {i + 1}
                    </span>
                  )}
                  <div>
                    <Text className={mergeClasses(s.stepTitle, isCurrent && s.stepTitleCurrent)} block>
                      {st.title}
                    </Text>
                    {isCurrent && (
                      <div className={s.stepBody}>
                        {st.startNote && session?.startedSteps?.includes(i) && (
                          <MessageBar intent="info" layout="multiline">
                            <MessageBarBody>
                              <RichText text={localize(st.startNote(data))} />
                            </MessageBarBody>
                          </MessageBar>
                        )}
                        <Body1>
                          <RichText text={localize(st.task(data))} />
                        </Body1>
                        {current.hints.slice(0, hintsShown).map((h, k) => (
                          <div key={k} className={s.hint}>
                            <span className={s.hintIcon} aria-hidden="true">
                              <Lightbulb20Regular />
                            </span>
                            <Caption1>
                              <RichText text={localize(h)} />
                            </Caption1>
                          </div>
                        ))}
                        {session?.revealed && <pre className={s.answer}>{localize(current.solution(data))}</pre>}
                      </div>
                    )}
                  </div>
                </li>
              );
            })}
          </ol>
        </section>

        {error && (
          <MessageBar intent="error" layout="multiline">
            <MessageBarBody>{error}</MessageBarBody>
          </MessageBar>
        )}

        <div ref={resultRef} className={s.result}>
          {report && phase === 'working' && <CheckResults report={report} onGoTo={goTo} />}
          {/* A new report remounts it, which clears the explanation. */}
          {hintsOn && report && canExplain(report) && phase === 'working' && (
            <HintPanel
              key={hintKey(report)}
              payload={buildHintPayload({
                title: `${m.title} · ${current.title}`,
                task: localize(current.task(data)),
                platform,
                report,
                hintsShown: current.hints.slice(0, hintsShown).map(localize),
              })}
            />
          )}
          {phase === 'stepPassed' && (
            <div className={s.stepOk} role="status">
              <CheckmarkCircle20Filled />
              <Text weight="semibold">Step {step + 1} done</Text>
            </div>
          )}
          {phase === 'complete' && (
            <section className={s.done} role="status">
              <CheckmarkCircle20Filled style={{ width: 32, height: 32, color: tokens.colorBrandForeground1 }} />
              <Text weight="semibold" size={500}>
                Mission complete{outcome ? ` in ${formatDuration(outcome.elapsedMs)}` : ''}
              </Text>
              <Caption1 className={s.muted}>
                {session?.revealed ? 'An answer was shown, so this run doesn’t set a best time.' : outcome?.personalBest && mp.completions > 1 ? 'New personal best.' : 'Every step checked out.'}
              </Caption1>
            </section>
          )}
        </div>
      </div>

      <footer className={s.footer}>
        {phase === 'brief' || phase === 'settingUp' || phase === 'loading' ? (
          <>
            <Button appearance="primary" size="large" icon={phase === 'brief' ? <Table20Regular /> : <Spinner size="tiny" />} disabled={busy} onClick={() => void start()}>
              {phase === 'settingUp' ? 'Setting up…' : phase === 'loading' ? 'Opening…' : 'Start mission'}
            </Button>
            <Caption1 className={s.muted}>Adds a sheet with this mission’s data and replaces earlier practice sheets.</Caption1>
          </>
        ) : phase === 'complete' ? (
          <div className={s.actions}>
            <Button className={s.grow} appearance="primary" icon={<ArrowSync20Regular />} onClick={() => void start()}>
              Run it again
            </Button>
            <Button onClick={goHome}>All missions</Button>
          </div>
        ) : phase === 'stepPassed' ? (
          <Button
            appearance="primary"
            size="large"
            icon={advancing ? <Spinner size="tiny" /> : <ArrowRight16Regular />}
            iconPosition="after"
            disabled={advancing}
            onClick={() => void nextStep()}
          >
            {advancing ? 'Updating the data…' : 'Next step'}
          </Button>
        ) : (
          <div className={s.actions}>
            <Button className={s.grow} appearance="primary" size="large" icon={phase === 'checking' ? <Spinner size="tiny" /> : undefined} disabled={phase === 'checking'} onClick={() => void check()}>
              {phase === 'checking' ? 'Checking…' : `Check step ${step + 1}`}
            </Button>
            <Button size="large" icon={<Lightbulb20Regular />} disabled={hintsShown >= current.hints.length || busy} onClick={showHint} aria-label={`Show hint ${hintsShown + 1} of ${current.hints.length}`}>
              {hintsShown === 0 ? 'Hint' : `${hintsShown}/${current.hints.length}`}
            </Button>
            <Menu positioning="above-end">
              <MenuTrigger disableButtonEnhancement>
                <Button size="large" icon={<MoreHorizontal20Regular />} aria-label="More actions" disabled={busy} />
              </MenuTrigger>
              <MenuPopover>
                <MenuList>
                  <MenuItem icon={<Eye20Regular />} disabled={session?.revealed} onClick={() => setConfirm('reveal')}>
                    Show answer for this step
                  </MenuItem>
                  <MenuItem icon={<ArrowSync20Regular />} onClick={() => setConfirm('restart')}>
                    Start over with new data
                  </MenuItem>
                </MenuList>
              </MenuPopover>
            </Menu>
          </div>
        )}
      </footer>

      <ConfirmDialog
        open={confirm === 'restart'}
        title="Start the mission over?"
        body="The mission sheet is replaced with new data, and your work on it is cleared."
        confirmLabel="Start over"
        onCancel={() => setConfirm(null)}
        onConfirm={() => void start()}
      />
      <ConfirmDialog
        open={confirm === 'reveal'}
        title="Show this step’s answer?"
        body="This run won’t set a best time. You can still enter the answer and check it."
        confirmLabel="Show answer"
        cancelLabel="Keep trying"
        onCancel={() => setConfirm(null)}
        onConfirm={reveal}
      />
    </div>
  );
}

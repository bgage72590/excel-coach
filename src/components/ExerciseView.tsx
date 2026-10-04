import {
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
  MessageBarActions,
  MessageBarBody,
  Spinner,
  Text,
  Tooltip,
  makeStyles,
  tokens,
} from '@fluentui/react-components';
import {
  ArrowLeft16Regular,
  ArrowSync20Regular,
  Copy16Regular,
  Dismiss16Regular,
  Eye20Regular,
  Lightbulb20Regular,
  MoreHorizontal20Regular,
  Table20Regular,
  Timer16Regular,
} from '@fluentui/react-icons';
import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { buildHintPayload, canExplain } from '../ai/hints';
import { MASTERY_PASSES, formatDuration, progressFor, recordAttempt, recordPass, statusOf, type PassOutcome } from '../engine/progress';
import { Rng, newSeed } from '../engine/rng';
import type { CheckReport } from '../engine/types';
import { CoachError, friendlyError } from '../excel/errors';
import { sheetNameFor } from '../excel/host';
import { EXERCISES, MODULES, getExercise } from '../exercises';
import { useCoach } from '../taskpane/context';
import { MasteryDots, RichText, SectionLabel, useNow } from './bits';
import { CheckResults } from './CheckResults';
import { ConceptPanel } from './ConceptPanel';
import { ConfirmDialog } from './ConfirmDialog';
import { HintPanel, hintKey, useHintsEnabled } from './HintPanel';
import { PassPanel } from './PassPanel';

type Phase = 'loading' | 'brief' | 'settingUp' | 'working' | 'checking' | 'passed';

const useStyles = makeStyles({
  view: { display: 'flex', flexDirection: 'column', minHeight: '100%' },
  content: {
    flex: 1,
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalL,
    padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalL} ${tokens.spacingVerticalXL}`,
  },
  back: { alignSelf: 'flex-start', marginLeft: `calc(-1 * ${tokens.spacingHorizontalS})` },
  titleBlock: { display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalXXS },
  eyebrow: { color: tokens.colorBrandForeground1, fontWeight: tokens.fontWeightSemibold },
  replaces: { color: tokens.colorNeutralForeground3 },
  meta: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalM,
    marginTop: tokens.spacingVerticalXS,
    color: tokens.colorNeutralForeground3,
    fontVariantNumeric: 'tabular-nums',
  },
  metaItem: { display: 'inline-flex', alignItems: 'center', gap: tokens.spacingHorizontalXS },
  task: { display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalXS },
  hints: { listStyle: 'none', margin: 0, padding: 0, display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalS },
  hint: {
    display: 'grid',
    gridTemplateColumns: '20px minmax(0, 1fr)',
    gap: tokens.spacingHorizontalS,
    padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalM}`,
    borderRadius: tokens.borderRadiusMedium,
    backgroundColor: tokens.colorNeutralBackground2,
  },
  hintIcon: { color: tokens.colorPaletteMarigoldForeground1, display: 'inline-flex' },
  solution: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalXS,
    padding: tokens.spacingHorizontalM,
    borderRadius: tokens.borderRadiusLarge,
    border: `1px dashed ${tokens.colorNeutralStroke1}`,
  },
  solutionHead: { display: 'flex', alignItems: 'center', justifyContent: 'space-between' },
  code: {
    fontFamily: tokens.fontFamilyMonospace,
    fontSize: tokens.fontSizeBase200,
    lineHeight: tokens.lineHeightBase300,
    overflowWrap: 'anywhere',
    margin: 0,
    whiteSpace: 'pre-wrap',
    userSelect: 'all',
  },
  // Scrolled into view after a check; the margin clears the sticky footer (about 65px) with room to spare.
  outcome: { display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalM, scrollMarginBottom: '88px' },
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
  footnote: { color: tokens.colorNeutralForeground3 },
});

export function ExerciseView({ id }: { id: string }) {
  const s = useStyles();
  const { host, hostNotice, platform, progress, setProgress, localize, goHome, openExercise } = useCoach();
  const ex = getExercise(id)!;
  const moduleTitle = MODULES.find((m) => m.id === ex.module)?.title ?? '';
  const index = EXERCISES.findIndex((e) => e.id === id);
  const p = progressFor(progress, id);
  const session = progress.session?.exerciseId === id ? progress.session : undefined;

  const [phase, setPhase] = useState<Phase>(session && !session.passedAt && host ? 'loading' : 'brief');
  const [report, setReport] = useState<CheckReport>();
  const [outcome, setOutcome] = useState<PassOutcome>();
  const [error, setError] = useState<string>();
  const [confirm, setConfirm] = useState<'newData' | 'reveal' | null>(null);
  const [copied, setCopied] = useState(false);
  const hintsOn = useHintsEnabled();
  const outcomeRef = useRef<HTMLDivElement>(null);
  // When a check finishes, the footer's buttons change under the pointer. Ignore clicks
  // for a moment so a quick double-click on Check can't also press "Next rep".
  const settledAt = useRef(0);
  const settled = () => Date.now() - settledAt.current > 700;
  const busy = phase === 'settingUp' || phase === 'checking' || phase === 'loading';

  // The rep's data is fully determined by its seed.
  const data = useMemo(() => ex.make(new Rng(session?.seed ?? 1)), [ex, session?.seed]);
  const now = useNow(phase === 'working' || phase === 'checking');
  const elapsed = session ? (session.passedAt ?? now) - session.startedAt : 0;

  // Resume an unfinished rep if its sheet is still in the workbook.
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

  const setUp = useCallback(async () => {
    if (!host) return setError(hostNotice);
    setConfirm(null);
    setError(undefined);
    setReport(undefined);
    setOutcome(undefined);
    setPhase('settingUp');
    const seed = newSeed();
    const sheet = sheetNameFor(id);
    try {
      await host.setup(ex, ex.make(new Rng(seed)), sheet);
      setProgress((prev) => ({
        ...prev,
        session: { exerciseId: id, seed, sheet, startedAt: Date.now(), attempts: 0, hintsShown: 0, revealed: false },
      }));
      setPhase('working');
    } catch (e) {
      setError(friendlyError(e));
      setPhase(session && !session.passedAt ? 'working' : 'brief');
    }
  }, [host, hostNotice, id, ex, setProgress, session]);

  const check = useCallback(async () => {
    if (!host) return setError(hostNotice);
    if (!session) return;
    setError(undefined);
    setPhase('checking');
    try {
      const r = await host.check(ex, data, session.sheet, session.seed);
      let next = recordAttempt(progress);
      if (r.passed) {
        const o = recordPass(next, Date.now());
        next = o.state;
        setOutcome(o);
        setPhase('passed');
      } else {
        setPhase('working');
      }
      setReport(r);
      setProgress(() => next);
      settledAt.current = Date.now();
    } catch (e) {
      setError(friendlyError(e));
      setPhase(e instanceof CoachError && e.code === 'missing-sheet' ? 'brief' : 'working');
    }
  }, [host, hostNotice, session, ex, data, progress, setProgress]);

  // Bring the result of a check into view; the task text can push it below the fold.
  useEffect(() => {
    if (!report || phase === 'checking') return;
    const reduce = matchMedia('(prefers-reduced-motion: reduce)').matches;
    outcomeRef.current?.scrollIntoView({ behavior: reduce ? 'auto' : 'smooth', block: 'nearest' });
  }, [report, phase]);

  // ⌘/Ctrl+Enter checks from anywhere in the panel.
  const checkRef = useRef(check);
  checkRef.current = check;
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Enter' && (e.metaKey || e.ctrlKey) && phase === 'working') {
        e.preventDefault();
        void checkRef.current();
      }
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, [phase]);

  const showHint = () =>
    setProgress((prev) => (prev.session ? { ...prev, session: { ...prev.session, hintsShown: Math.min(ex.hints.length, prev.session.hintsShown + 1) } } : prev));

  const reveal = () => {
    setConfirm(null);
    setProgress((prev) => (prev.session ? { ...prev, session: { ...prev.session, revealed: true } } : prev));
  };

  const requestNewData = () => {
    const invested = session && !session.passedAt && (session.attempts > 0 || Date.now() - session.startedAt > 30_000);
    if (invested && phase !== 'passed') setConfirm('newData');
    else void setUp();
  };

  const goTo = (address: string) => {
    if (host && session) host.select(session.sheet, address).catch((e) => setError(friendlyError(e)));
  };

  const nextRep = () => {
    if (settled()) void setUp();
  };

  const nextSkill = () => {
    if (!settled()) return;
    const nowMs = Date.now();
    const after = [...EXERCISES.slice(index + 1), ...EXERCISES.slice(0, index)];
    const target = after.find((e) => statusOf(progressFor(progress, e.id), nowMs) !== 'mastered') ?? after[0];
    openExercise(target.id);
  };

  const copySolution = async (text: string) => {
    try {
      await navigator.clipboard.writeText(text);
      setCopied(true);
      setTimeout(() => setCopied(false), 1600);
    } catch {
      // Clipboard can be blocked inside Office; the text is selectable instead.
    }
  };

  const passes = Math.min(p.passSeeds.length, MASTERY_PASSES);
  const hintsShown = session?.hintsShown ?? 0;
  const revealed = session?.revealed ?? false;
  const status = statusOf(p, Date.now());
  const concept = {
    summary: localize(ex.concept.summary),
    syntax: localize(ex.concept.syntax),
    example: ex.concept.example && localize(ex.concept.example),
    tip: ex.concept.tip && localize(ex.concept.tip),
  };
  const solution = localize(ex.solution(data));

  return (
    <div className={s.view}>
      <div className={s.content}>
        <Button className={s.back} appearance="transparent" size="small" icon={<ArrowLeft16Regular />} onClick={goHome}>
          All skills
        </Button>

        <div className={s.titleBlock}>
          <Caption1 className={s.eyebrow}>
            {moduleTitle} · Skill {index + 1} of {EXERCISES.length}
          </Caption1>
          <Text as="h1" size={500} weight="semibold" style={{ margin: 0 }}>
            {ex.title}
          </Text>
          <Caption1 className={s.replaces}>Replaces: {ex.replaces}</Caption1>
          <div className={s.meta}>
            {session && (phase === 'working' || phase === 'checking' || phase === 'passed') ? (
              <span className={s.metaItem} aria-label="Time on this rep">
                <Timer16Regular aria-hidden="true" />
                <Caption1>{formatDuration(elapsed)}</Caption1>
              </span>
            ) : (
              <span className={s.metaItem}>
                <Timer16Regular aria-hidden="true" />
                <Caption1>About {ex.minutes} min</Caption1>
              </span>
            )}
            <span className={s.metaItem}>
              <MasteryDots passes={passes} />
            </span>
            {p.bestMs !== undefined && <Caption1>Best {formatDuration(p.bestMs)}</Caption1>}
            {ex.m365 && (
              <Tooltip content="Uses a function that needs Microsoft 365" relationship="description">
                <Badge appearance="outline" size="small" color="informative">
                  365
                </Badge>
              </Tooltip>
            )}
          </div>
        </div>

        {hostNotice && (
          <MessageBar intent="info">
            <MessageBarBody>{hostNotice}</MessageBarBody>
          </MessageBar>
        )}

        <section className={s.task} aria-label="Task">
          <SectionLabel>Your task</SectionLabel>
          <Body1>
            <RichText text={localize(ex.task(data))} />
          </Body1>
        </section>

        <ConceptPanel concept={concept} defaultOpen={status === 'new' && !session} />

        {hintsShown > 0 && (
          <section aria-label="Hints">
            <ul className={s.hints}>
              {ex.hints.slice(0, hintsShown).map((h, i) => (
                <li key={i} className={s.hint}>
                  <span className={s.hintIcon} aria-hidden="true">
                    <Lightbulb20Regular />
                  </span>
                  <Caption1>
                    <RichText text={localize(h)} />
                  </Caption1>
                </li>
              ))}
            </ul>
          </section>
        )}

        {revealed && (
          <section className={s.solution} aria-label="Answer">
            <div className={s.solutionHead}>
              <SectionLabel as="h3">Answer</SectionLabel>
              <Button size="small" appearance="subtle" icon={<Copy16Regular />} onClick={() => copySolution(solution)}>
                {copied ? 'Copied' : 'Copy'}
              </Button>
            </div>
            <pre className={s.code}>{solution}</pre>
          </section>
        )}

        {error && (
          <MessageBar intent="error" layout="multiline">
            <MessageBarBody>{error}</MessageBarBody>
            <MessageBarActions
              containerAction={<Button appearance="transparent" size="small" icon={<Dismiss16Regular />} aria-label="Dismiss" onClick={() => setError(undefined)} />}
            />
          </MessageBar>
        )}

        <div ref={outcomeRef} className={s.outcome}>
          {report && phase !== 'passed' && phase !== 'settingUp' && <CheckResults report={report} onGoTo={goTo} />}
          {/* A new report remounts it, which clears the explanation. */}
          {hintsOn && report && canExplain(report) && phase === 'working' && (
            <HintPanel
              key={hintKey(report)}
              payload={buildHintPayload({ title: ex.title, task: localize(ex.task(data)), platform, report, hintsShown: ex.hints.slice(0, hintsShown).map(localize) })}
            />
          )}
          {phase === 'passed' && report && (
            <PassPanel outcome={outcome} report={report} revealed={revealed} passes={Math.min(progressFor(progress, id).passSeeds.length, MASTERY_PASSES)} mastered={p.mastered} />
          )}
        </div>
      </div>

      <footer className={s.footer}>
        {phase === 'brief' || phase === 'settingUp' || phase === 'loading' ? (
          <>
            <Button
              appearance="primary"
              size="large"
              icon={phase === 'brief' ? <Table20Regular /> : <Spinner size="tiny" />}
              disabled={busy}
              onClick={() => void setUp()}
            >
              {phase === 'settingUp' ? 'Setting up…' : phase === 'loading' ? 'Opening…' : 'Set up practice sheet'}
            </Button>
            <Caption1 className={s.footnote}>Adds a sheet with fresh data and replaces earlier practice sheets.</Caption1>
          </>
        ) : phase === 'passed' ? (
          <div className={s.actions}>
            {p.mastered ? (
              <>
                <Button className={s.grow} appearance="primary" onClick={nextSkill}>
                  Next skill
                </Button>
                <Button onClick={nextRep}>Another rep</Button>
              </>
            ) : (
              <>
                <Button className={s.grow} appearance="primary" icon={<ArrowSync20Regular />} onClick={nextRep}>
                  Next rep
                </Button>
                <Button onClick={nextSkill}>Next skill</Button>
              </>
            )}
          </div>
        ) : (
          <div className={s.actions}>
            <Tooltip content={`Check your work (${navigator.platform.includes('Mac') ? '⌘' : 'Ctrl+'}↩)`} relationship="description">
              <Button
                className={s.grow}
                appearance="primary"
                size="large"
                icon={phase === 'checking' ? <Spinner size="tiny" /> : undefined}
                disabled={phase === 'checking'}
                onClick={() => settled() && void check()}
              >
                {phase === 'checking' ? 'Checking…' : 'Check'}
              </Button>
            </Tooltip>
            <Button
              size="large"
              icon={<Lightbulb20Regular />}
              disabled={hintsShown >= ex.hints.length || busy}
              onClick={showHint}
              aria-label={hintsShown >= ex.hints.length ? 'No more hints' : `Show hint ${hintsShown + 1} of ${ex.hints.length}`}
            >
              {hintsShown === 0 ? 'Hint' : `${hintsShown}/${ex.hints.length}`}
            </Button>
            <Menu positioning="above-end">
              <MenuTrigger disableButtonEnhancement>
                <Button size="large" icon={<MoreHorizontal20Regular />} aria-label="More actions" disabled={busy} />
              </MenuTrigger>
              <MenuPopover>
                <MenuList>
                  <MenuItem icon={<ArrowSync20Regular />} onClick={requestNewData}>
                    New data
                  </MenuItem>
                  <MenuItem icon={<Eye20Regular />} disabled={revealed} onClick={() => setConfirm('reveal')}>
                    Show answer
                  </MenuItem>
                </MenuList>
              </MenuPopover>
            </Menu>
          </div>
        )}
      </footer>

      <ConfirmDialog
        open={confirm === 'newData'}
        title="Start over with new data?"
        body="The practice sheet is replaced with a fresh one, and your formulas on it are cleared."
        confirmLabel="Start over"
        onCancel={() => setConfirm(null)}
        onConfirm={() => void setUp()}
      />
      <ConfirmDialog
        open={confirm === 'reveal'}
        title="Show the answer?"
        body="This rep won’t count toward mastery. You can still enter the answer and check it."
        confirmLabel="Show answer"
        cancelLabel="Keep trying"
        onCancel={() => setConfirm(null)}
        onConfirm={reveal}
      />
    </div>
  );
}

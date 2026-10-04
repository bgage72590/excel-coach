import {
  Body1,
  Button,
  Caption1,
  Divider,
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
  ArrowReset20Regular,
  CheckmarkCircle20Filled,
  Dismiss16Regular,
  DocumentCopy20Regular,
  LockClosed16Regular,
  MoreHorizontal20Regular,
  Open20Regular,
} from '@fluentui/react-icons';
import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { rangeSize } from '../engine/address';
import { copyReferences, gradeFix, planFix, scanEdgeBelow, type FixCopy } from '../engine/fix';
import { isFormula } from '../engine/formula';
import type { CheckReport } from '../engine/types';
import { CoachError, friendlyError } from '../excel/errors';
import { fixSheetNameFor } from '../excel/host';
import { useCoach, type FixTarget } from '../taskpane/context';
import { RichText, SectionLabel } from './bits';
import { CheckResults } from './CheckResults';
import { ConfirmDialog } from './ConfirmDialog';

/** 'found': a copy from an earlier visit is still in the workbook. */
type Phase = 'loading' | 'brief' | 'found' | 'copying' | 'working' | 'checking' | 'passed';

const rise = { from: { opacity: 0, transform: 'translateY(4px)' }, to: { opacity: 1, transform: 'translateY(0)' } };
/** Counts in the copy, which is English: 1,240 cells. */
const count = new Intl.NumberFormat('en-US');

const useStyles = makeStyles({
  view: { display: 'flex', flexDirection: 'column', minHeight: '100%' },
  content: { flex: 1, display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalL, padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalL} ${tokens.spacingVerticalXL}` },
  back: { alignSelf: 'flex-start', marginLeft: `calc(-1 * ${tokens.spacingHorizontalS})` },
  titleBlock: { display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalXXS },
  eyebrow: { color: tokens.colorBrandForeground1, fontWeight: tokens.fontWeightSemibold, overflowWrap: 'anywhere' },
  muted: { color: tokens.colorNeutralForeground3 },
  privacy: { display: 'grid', gridTemplateColumns: '16px minmax(0, 1fr)', gap: tokens.spacingHorizontalS, color: tokens.colorNeutralForeground2, alignItems: 'start' },
  privacyIcon: { marginTop: '2px' },
  section: { display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalXS },
  steps: { listStyle: 'none', margin: 0, padding: 0, display: 'flex', flexDirection: 'column' },
  step: {
    display: 'grid',
    gridTemplateColumns: '24px minmax(0, 1fr)',
    gap: tokens.spacingHorizontalS,
    alignItems: 'start',
    padding: `${tokens.spacingVerticalS} 0`,
    borderBottom: `1px solid ${tokens.colorNeutralStroke3}`,
    ':last-child': { borderBottom: 'none' },
  },
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
  stepText: { paddingTop: '1px', overflowWrap: 'anywhere' },
  empty: { display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalXS, padding: tokens.spacingHorizontalM, borderRadius: tokens.borderRadiusLarge, backgroundColor: tokens.colorNeutralBackground2 },
  outcome: { display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalM, scrollMarginBottom: '16px' },
  done: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalS,
    padding: `${tokens.spacingVerticalL} ${tokens.spacingHorizontalM}`,
    borderRadius: tokens.borderRadiusLarge,
    backgroundColor: tokens.colorNeutralBackground2,
    animationName: rise,
    animationDuration: '260ms',
    animationTimingFunction: tokens.curveDecelerateMid,
    '@media (prefers-reduced-motion: reduce)': { animationDuration: '1ms' },
  },
  doneHead: { display: 'grid', gridTemplateColumns: '32px minmax(0, 1fr)', gap: tokens.spacingHorizontalS, alignItems: 'center' },
  doneIcon: { width: '32px', height: '32px', color: tokens.colorBrandForeground1 },
  divider: { margin: `${tokens.spacingVerticalXS} 0` },
  nextSteps: {
    margin: 0,
    paddingLeft: tokens.spacingHorizontalL,
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalXS,
    fontSize: tokens.fontSizeBase200,
    lineHeight: tokens.lineHeightBase200,
    overflowWrap: 'anywhere',
  },
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
  actions: { display: 'flex', flexWrap: 'wrap', gap: tokens.spacingHorizontalS, alignItems: 'center' },
  grow: { flex: 1 },
  footnote: { color: tokens.colorNeutralForeground3 },
});

/** Guided fix on a copy of the learner's sheet: steps, then a check that the copy's results still match. */
export function FixView({ target }: { target: FixTarget }) {
  const s = useStyles();
  const { host, hostNotice, localize, openScan } = useCoach();
  // True once the panel finds the column going on below the rows the scan read.
  const [partial, setPartial] = useState(false);
  const plan = useMemo(() => planFix(target.finding, target.sheet, { partial }), [target, partial]);
  const copy = fixSheetNameFor(target.sheet.sheet);
  const fixCopy = useMemo<FixCopy>(() => ({ name: copy, original: target.sheet }), [copy, target.sheet]);

  const [phase, setPhase] = useState<Phase>(host && plan ? 'loading' : 'brief');
  const [report, setReport] = useState<CheckReport>();
  const [error, setError] = useState<string>();
  const [confirm, setConfirm] = useState(false);
  /** Tables the passing copy's formulas name that the original's don't: maybe the copy's own. */
  const [unplaced, setUnplaced] = useState<string[]>([]);
  const outcomeRef = useRef<HTMLDivElement>(null);
  // The footer's buttons change under the pointer when a copy or check finishes. Ignore clicks for
  // a moment so a quick double-click can't also press the button that takes the old one's place.
  const settledAt = useRef(0);
  const settled = () => Date.now() - settledAt.current > 700;
  const busy = phase === 'loading' || phase === 'copying' || phase === 'checking';

  // Offer to pick up where the learner left off when a copy is still in the workbook. The scan reads
  // very large sheets in part, so when the block ends on the last row it read, look one row further:
  // a formula there means the column goes on, and the steps say so.
  useEffect(() => {
    if (phase !== 'loading' || !host || !plan) return;
    let cancelled = false;
    const below = scanEdgeBelow(plan, target.sheet);
    void Promise.all([
      host.sheetExists(copy).catch(() => false),
      below
        ? host
            .readRange(plan.sheet, below)
            .then((r) => isFormula(r.formulas[0]?.[0]))
            .catch(() => false)
        : false,
    ]).then(([exists, continues]) => {
      if (cancelled) return;
      if (continues) setPartial(true);
      setPhase(exists ? 'found' : 'brief');
    });
    return () => {
      cancelled = true;
    };
  }, [phase, host, copy, plan, target.sheet]);

  const makeCopy = useCallback(async () => {
    if (!host || !plan) return setError(hostNotice);
    const from = phase;
    setConfirm(false);
    setError(undefined);
    setReport(undefined);
    setPhase('copying');
    try {
      const name = await host.copySheetForFix(plan.sheet);
      await host.select(name, plan.range);
      settledAt.current = Date.now();
      setPhase('working');
    } catch (e) {
      setError(friendlyError(e));
      // Starting over replaces the earlier copy, so it may be gone even though the new one failed.
      const exists = await host.sheetExists(copy).catch(() => false);
      setPhase(!exists ? 'brief' : from === 'working' ? 'working' : 'found');
    }
  }, [host, hostNotice, plan, phase, copy]);

  const goToCopy = async (address = plan?.range): Promise<boolean> => {
    if (!host || !address) return false;
    try {
      await host.select(copy, address);
      return true;
    } catch (e) {
      // The learner can delete the copy while the panel is open. Then the way forward is a new one.
      if (!(await host.sheetExists(copy).catch(() => true))) {
        setError(`${copy} isn’t in this workbook anymore. Make a new copy to continue.`);
        setReport(undefined);
        setPhase('brief');
      } else {
        setError(friendlyError(e));
      }
      return false;
    }
  };

  const continueOnCopy = async () => {
    setError(undefined);
    if (!(await goToCopy())) return;
    settledAt.current = Date.now();
    setPhase('working');
  };

  const check = useCallback(async () => {
    if (!host || !plan) return setError(hostNotice);
    setError(undefined);
    setPhase('checking');
    try {
      const before = await host.readRange(plan.sheet, plan.range);
      const after = await host.readRange(copy, plan.range);
      const items = gradeFix(plan, before, after, fixCopy);
      const passed = items.every((i) => i.status === 'pass');
      setUnplaced(passed ? copyReferences(plan, after, fixCopy).unknown : []);
      setReport({ passed, items, marks: [], focus: items.find((i) => i.status === 'fail' && i.focus)?.focus });
      settledAt.current = Date.now();
      setPhase(passed ? 'passed' : 'working');
    } catch (e) {
      setError(friendlyError(e));
      // When the copy itself is gone, the way forward is a fresh one.
      const copyGone = e instanceof CoachError && e.code === 'missing-sheet' && !(await host.sheetExists(copy).catch(() => false));
      setPhase(copyGone ? 'brief' : 'working');
    }
  }, [host, hostNotice, plan, copy, fixCopy]);

  // Bring the result of a check into view; the steps can push it below the fold.
  useEffect(() => {
    if (!report || phase === 'checking') return;
    const reduce = matchMedia('(prefers-reduced-motion: reduce)').matches;
    outcomeRef.current?.scrollIntoView({ behavior: reduce ? 'auto' : 'smooth', block: 'nearest' });
  }, [report, phase]);

  // ⌘/Ctrl+Enter checks from anywhere in the panel, but not from the Start over dialog.
  const checkRef = useRef(check);
  checkRef.current = check;
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Enter' && (e.metaKey || e.ctrlKey) && phase === 'working' && !confirm && settled()) {
        e.preventDefault();
        void checkRef.current();
      }
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, [phase, confirm]);

  const back = (
    <Button className={s.back} appearance="transparent" size="small" icon={<ArrowLeft16Regular />} onClick={openScan}>
      My Work
    </Button>
  );

  if (!plan) {
    return (
      <div className={s.view}>
        <div className={s.content}>
          {back}
          <div className={s.titleBlock}>
            <Caption1 className={s.eyebrow}>{target.sheet.sheet}</Caption1>
            <Text as="h1" size={500} weight="semibold" style={{ margin: 0 }}>
              {target.finding.title}
            </Text>
          </div>
          <section className={s.empty} role="status">
            <Text weight="semibold">No guided fix for this spot</Text>
            <Caption1>
              The coach can’t plan a fix here that it can check automatically. Usually that’s because the change can give different results on purpose. Use Practice the fix on My Work to learn
              the change, then make it by hand.
            </Caption1>
          </section>
        </div>
      </div>
    );
  }

  const { rows, cols } = rangeSize(plan.range);
  const cells = rows * cols;
  const changed = plan.valuesMayChange.length;
  const shortcut = navigator.platform.includes('Mac') ? '⌘' : 'Ctrl+';
  const copyBack = [
    plan.copyBack ?? `When you’re ready, copy the new formulas back: select \`${plan.range}\` on \`${copy}\`, copy, and paste onto the same range on \`${plan.sheet}\`.`,
    plan.continues,
  ]
    .filter(Boolean)
    .join(' ');
  const tableCheck =
    unplaced.length > 0 &&
    `First, check the Table names in the formulas. Excel adds a number to a Table’s name when it copies a sheet, so if ${unplaced.map((t) => `\`${t}\``).join(' or ')} is on the copy, change it to your own Table’s name.`;

  return (
    <div className={s.view}>
      <div className={s.content}>
        {back}

        <div className={s.titleBlock}>
          <Caption1 className={s.eyebrow}>
            {plan.sheet} · {plan.range}
          </Caption1>
          <Text as="h1" size={500} weight="semibold" style={{ margin: 0 }}>
            {plan.title}
          </Text>
          <Caption1 className={s.muted}>{target.finding.title}</Caption1>
        </div>

        <div className={s.privacy}>
          <LockClosed16Regular className={s.privacyIcon} aria-hidden="true" />
          <Caption1>Works on a copy named {copy}. Your original sheet isn’t changed.</Caption1>
        </div>

        {hostNotice && (
          <MessageBar intent="info">
            <MessageBarBody>{hostNotice}</MessageBarBody>
          </MessageBar>
        )}

        <Body1>
          <RichText text={localize(plan.intro)} />
        </Body1>

        <section className={s.section} aria-label="Steps">
          <SectionLabel>Steps</SectionLabel>
          {/* role="list" keeps list semantics in Safari, which drops them when the markers are hidden. */}
          <ol className={s.steps} role="list">
            {plan.steps.map((step, i) => (
              <li key={i} className={s.step}>
                <span className={s.stepNum} aria-hidden="true">
                  {i + 1}
                </span>
                <Body1 className={s.stepText}>
                  <RichText text={localize(step)} />
                </Body1>
              </li>
            ))}
          </ol>
        </section>

        {error && (
          <MessageBar intent="error" layout="multiline">
            <MessageBarBody>{error}</MessageBarBody>
            <MessageBarActions
              containerAction={<Button appearance="transparent" size="small" icon={<Dismiss16Regular />} aria-label="Dismiss" onClick={() => setError(undefined)} />}
            />
          </MessageBar>
        )}

        <div ref={outcomeRef} className={s.outcome}>
          {report && phase === 'working' && <CheckResults report={report} onGoTo={(address) => void goToCopy(address)} />}
          {phase === 'passed' && (
            <section className={s.done} role="status" aria-live="polite">
              <div className={s.doneHead}>
                <CheckmarkCircle20Filled className={s.doneIcon} aria-hidden="true" />
                <div>
                  <Text as="p" weight="semibold" size={400} block style={{ margin: 0 }}>
                    Your fix holds
                  </Text>
                  <Caption1 className={s.muted}>
                    {changed === 0
                      ? `The copy matches the original in all ${count.format(cells)} ${cells === 1 ? 'cell' : 'cells'}.`
                      : `The copy matches the original in ${count.format(cells - changed)} of ${count.format(cells)} cells. The other ${count.format(changed)} ${changed === 1 ? 'is the cell' : 'are the cells'} the fix was allowed to change.`}
                  </Caption1>
                </div>
              </div>
              <Divider className={s.divider} />
              <SectionLabel as="h3">Next steps</SectionLabel>
              <ol className={s.nextSteps}>
                {tableCheck && (
                  <li>
                    <Caption1>
                      <RichText text={tableCheck} />
                    </Caption1>
                  </li>
                )}
                <li>
                  <Caption1>
                    <RichText text={localize(copyBack)} />
                  </Caption1>
                </li>
                <li>
                  <Caption1>Then delete the copy: right-click its tab › Delete.</Caption1>
                </li>
              </ol>
              <Caption1 className={s.muted}>The coach never writes to or deletes your original sheet.</Caption1>
            </section>
          )}
        </div>
      </div>

      <footer className={s.footer}>
        {phase === 'loading' || phase === 'brief' || phase === 'copying' ? (
          <>
            <Button
              appearance="primary"
              size="large"
              icon={phase === 'brief' ? <DocumentCopy20Regular /> : <Spinner size="tiny" />}
              disabled={busy}
              onClick={() => void makeCopy()}
            >
              {phase === 'copying' ? 'Making a copy…' : phase === 'loading' ? 'Opening…' : 'Make a copy'}
            </Button>
            <Caption1 className={s.footnote}>Adds {copy} next to your sheet and selects the cells to fix.</Caption1>
          </>
        ) : phase === 'found' ? (
          <>
            <div className={s.actions}>
              <Button className={s.grow} appearance="primary" size="large" icon={<Open20Regular />} onClick={() => void continueOnCopy()}>
                Continue on the copy
              </Button>
              <Button size="large" icon={<ArrowReset20Regular />} onClick={() => setConfirm(true)}>
                Start over
              </Button>
            </div>
            <Caption1 className={s.footnote}>{copy} from earlier is still in this workbook.</Caption1>
          </>
        ) : phase === 'passed' ? (
          <div className={s.actions}>
            <Button className={s.grow} appearance="primary" onClick={() => settled() && openScan()}>
              Back to My Work
            </Button>
            <Button icon={<Open20Regular />} onClick={() => settled() && void goToCopy()}>
              Go to the copy
            </Button>
          </div>
        ) : (
          <div className={s.actions}>
            <Tooltip content={`Check the copy against your original (${shortcut}↩)`} relationship="description">
              <Button
                className={s.grow}
                appearance="primary"
                size="large"
                icon={phase === 'checking' ? <Spinner size="tiny" /> : undefined}
                disabled={phase === 'checking'}
                onClick={() => settled() && void check()}
              >
                {phase === 'checking' ? 'Checking…' : 'Check my fix'}
              </Button>
            </Tooltip>
            <Menu positioning="above-end">
              <MenuTrigger disableButtonEnhancement>
                <Button size="large" icon={<MoreHorizontal20Regular />} aria-label="More actions" disabled={busy} />
              </MenuTrigger>
              <MenuPopover>
                <MenuList>
                  <MenuItem icon={<Open20Regular />} onClick={() => void goToCopy()}>
                    Go to the cells to fix
                  </MenuItem>
                  <MenuItem icon={<ArrowReset20Regular />} onClick={() => setConfirm(true)}>
                    Start over with a fresh copy
                  </MenuItem>
                </MenuList>
              </MenuPopover>
            </Menu>
          </div>
        )}
      </footer>

      <ConfirmDialog
        open={confirm}
        title="Start over with a fresh copy?"
        body={`${copy} is replaced with a new copy of ${plan.sheet}, and your changes on it are cleared. Your original sheet isn’t changed.`}
        confirmLabel="Start over"
        onCancel={() => setConfirm(false)}
        onConfirm={() => void makeCopy()}
      />
    </div>
  );
}

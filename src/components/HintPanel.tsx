import { Body1, Button, Caption1, MessageBar, MessageBarActions, MessageBarBody, Skeleton, SkeletonItem, makeStyles, tokens } from '@fluentui/react-components';
import { ArrowClockwise16Regular, Sparkle16Regular, Sparkle20Regular } from '@fluentui/react-icons';
import { useEffect, useRef, useState } from 'react';
import type { HintPayload } from '../ai/contract';
import type { CheckReport } from '../engine/types';
import { CLIENT_MESSAGES, HintError, hintStatus, knownHintStatus, requestHint } from '../ai/hints';
import { RichText } from './bits';

const slideIn = { from: { opacity: 0, transform: 'translateY(4px)' }, to: { opacity: 1, transform: 'translateY(0)' } };

const useStyles = makeStyles({
  ask: { display: 'flex', flexDirection: 'column', alignItems: 'flex-start', gap: tokens.spacingVerticalXS },
  note: { color: tokens.colorNeutralForeground3 },
  // scrollIntoView ignores the views' sticky footer (about 65px), so the margin keeps the footnote clear of it.
  region: { scrollMarginBottom: '88px', ':focus': { outlineStyle: 'none' } },
  card: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalS,
    padding: tokens.spacingHorizontalM,
    borderRadius: tokens.borderRadiusLarge,
    backgroundColor: tokens.colorNeutralBackground1,
    border: `1px solid ${tokens.colorBrandStroke2}`,
    animationName: slideIn,
    animationDuration: '220ms',
    animationTimingFunction: tokens.curveDecelerateMid,
    '@media (prefers-reduced-motion: reduce)': { animationDuration: '1ms' },
  },
  head: { display: 'flex', alignItems: 'center', gap: tokens.spacingHorizontalXS, color: tokens.colorBrandForeground1 },
  label: { fontWeight: tokens.fontWeightSemibold },
  lines: { display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalS, paddingBlock: tokens.spacingVerticalXXS },
  short: { width: '60%' },
  // Claude may quote a whole formula; let it wrap instead of pushing the panel sideways.
  text: { display: 'block', '& code': { whiteSpace: 'normal', overflowWrap: 'anywhere' } },
  footnote: { color: tokens.colorNeutralForeground3 },
});

type State = { kind: 'idle' } | { kind: 'loading' } | { kind: 'error'; message: string } | { kind: 'done'; hint: string };

/** Whether the local hint server is available. Starts from the cached answer, so there's no flash. */
export function useHintsEnabled(): boolean {
  const [enabled, setEnabled] = useState(() => knownHintStatus() ?? false);
  useEffect(() => {
    let live = true;
    void hintStatus().then((on) => live && setEnabled(on));
    return () => {
      live = false;
    };
  }, []);
  return enabled;
}

const reportKeys = new WeakMap<CheckReport, number>();
let lastReportKey = 0;

/** A React key that's new for every check report. Use it on HintPanel to remount it per check. */
export function hintKey(report: CheckReport): number {
  let key = reportKeys.get(report);
  if (key === undefined) reportKeys.set(report, (key = ++lastReportKey));
  return key;
}

/**
 * "Explain my mistake": asks Claude about the last failed check. Render it only on practice sheets
 * (exercises and missions), keyed with hintKey(report) so an explanation never outlives the results
 * it explains.
 */
export function HintPanel({ payload }: { payload: HintPayload }) {
  const s = useStyles();
  const [state, setState] = useState<State>({ kind: 'idle' });
  const pending = useRef<AbortController | null>(null);
  const regionRef = useRef<HTMLDivElement>(null);

  useEffect(() => () => pending.current?.abort(), []);

  // The button that was pressed goes away, so keep keyboard focus on the answer as it arrives.
  useEffect(() => {
    if (state.kind === 'idle') return;
    const region = regionRef.current;
    if (!region) return;
    if (state.kind === 'loading') region.focus({ preventScroll: true });
    const reduce = matchMedia('(prefers-reduced-motion: reduce)').matches;
    region.scrollIntoView({ behavior: reduce ? 'auto' : 'smooth', block: 'nearest' });
  }, [state.kind]);

  const ask = async () => {
    pending.current?.abort();
    const ctrl = new AbortController();
    pending.current = ctrl;
    setState({ kind: 'loading' });
    try {
      const hint = await requestHint(payload, ctrl.signal);
      if (!ctrl.signal.aborted) setState({ kind: 'done', hint });
    } catch (e) {
      if (!ctrl.signal.aborted) setState({ kind: 'error', message: e instanceof HintError ? e.message : CLIENT_MESSAGES.failed });
    }
  };

  if (state.kind === 'idle') {
    return (
      <div className={s.ask}>
        <Button icon={<Sparkle20Regular />} onClick={() => void ask()}>
          Explain my mistake
        </Button>
        <Caption1 className={s.note}>Sends your formulas and these results to Claude.</Caption1>
      </div>
    );
  }

  return (
    // The MessageBar announces itself; the card is a live region so the answer is read when it lands.
    <div ref={regionRef} className={s.region} tabIndex={-1}>
      {state.kind === 'error' ? (
        <MessageBar intent="error" layout="multiline">
          <MessageBarBody>{state.message}</MessageBarBody>
          <MessageBarActions>
            <Button size="small" icon={<ArrowClockwise16Regular />} onClick={() => void ask()}>
              Try again
            </Button>
          </MessageBarActions>
        </MessageBar>
      ) : (
        <section className={s.card} aria-label="Explanation from Claude" aria-live="polite" aria-busy={state.kind === 'loading'}>
          <div className={s.head}>
            <Sparkle16Regular aria-hidden="true" />
            <Caption1 className={s.label}>From Claude</Caption1>
          </div>
          {state.kind === 'loading' ? (
            <Skeleton className={s.lines} aria-label="Claude is reading your formulas">
              <SkeletonItem size={12} />
              <SkeletonItem size={12} />
              <SkeletonItem size={12} className={s.short} />
            </Skeleton>
          ) : (
            <>
              <Body1 className={s.text}>
                <RichText text={state.hint} />
              </Body1>
              <Caption1 className={s.footnote}>AI can make mistakes. Check the explanation against your sheet.</Caption1>
            </>
          )}
        </section>
      )}
    </div>
  );
}

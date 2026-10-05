import { Body1, Button, Caption1, Link, Text, makeStyles, mergeClasses, tokens } from '@fluentui/react-components';
import { ArrowLeft16Regular, ArrowRight16Regular, Checkmark16Filled, CheckmarkCircle16Filled, Dismiss16Regular, Eye16Regular, Warning16Filled } from '@fluentui/react-icons';
import { useEffect, useRef, useState } from 'react';
import { partsFormula, type StepProbe } from '../engine/guide';
import type { FormulaPart, GuideStep, SheetPointer, SheetSpot, StepDone } from '../engine/types';
import { RichText, SectionLabel } from './bits';

/** How often to look at the workbook while a step is waiting. */
const POLL_MS = 1000;
/** How long "Done" shows before the next step opens. */
const ADVANCE_MS = 1400;

const pulse = { '0%, 100%': { opacity: 0.35 }, '50%': { opacity: 1 } };
const slideIn = { from: { opacity: 0, transform: 'translateY(4px)' }, to: { opacity: 1, transform: 'translateY(0)' } };

// Part colors, in order. Readable on light and dark themes.
const PART_COLORS = [
  tokens.colorPaletteBerryForeground1,
  tokens.colorPaletteBlueForeground2,
  tokens.colorPaletteDarkOrangeForeground1,
  tokens.colorPaletteGreenForeground1,
  tokens.colorPaletteGrapeForeground2,
  tokens.colorPaletteTealForeground2,
];

const useStyles = makeStyles({
  guide: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalM,
    padding: tokens.spacingHorizontalM,
    borderRadius: tokens.borderRadiusLarge,
    border: `1px solid ${tokens.colorBrandStroke2}`,
    backgroundColor: tokens.colorNeutralBackground1,
  },
  head: { display: 'flex', alignItems: 'center', gap: tokens.spacingHorizontalS },
  headText: { flex: 1, display: 'flex', alignItems: 'baseline', gap: tokens.spacingHorizontalS },
  count: { color: tokens.colorNeutralForeground3, fontVariantNumeric: 'tabular-nums' },
  bar: { display: 'flex', gap: '3px' },
  seg: { flex: 1, height: '4px', borderRadius: '2px', backgroundColor: tokens.colorNeutralStroke2 },
  segDone: { backgroundColor: tokens.colorBrandBackground },
  segNow: { backgroundColor: tokens.colorBrandStroke2 },
  steps: { listStyle: 'none', margin: 0, padding: 0, display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalS },
  step: { display: 'grid', gridTemplateColumns: '22px minmax(0, 1fr)', gap: tokens.spacingHorizontalS, alignItems: 'start' },
  marker: {
    width: '22px',
    height: '22px',
    borderRadius: '50%',
    boxSizing: 'border-box',
    display: 'inline-flex',
    alignItems: 'center',
    justifyContent: 'center',
    fontSize: tokens.fontSizeBase100,
    fontWeight: tokens.fontWeightSemibold,
    border: `1.5px solid ${tokens.colorNeutralStroke1}`,
    color: tokens.colorNeutralForeground3,
  },
  markerNow: { backgroundColor: tokens.colorBrandBackground, border: `1.5px solid ${tokens.colorBrandBackground}`, color: tokens.colorNeutralForegroundOnBrand },
  markerDone: { border: 'none', color: tokens.colorPaletteGreenForeground1 },
  stepBody: { display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalS, minWidth: 0, paddingTop: '1px' },
  future: { color: tokens.colorNeutralForeground3 },
  past: { color: tokens.colorNeutralForeground3 },
  current: {
    animationName: slideIn,
    animationDuration: '200ms',
    animationTimingFunction: tokens.curveDecelerateMid,
    '@media (prefers-reduced-motion: reduce)': { animationDuration: '1ms' },
  },
  why: { color: tokens.colorNeutralForeground2 },
  // The formula, each part in its color, then one row per part saying what it means.
  card: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalS,
    padding: tokens.spacingHorizontalM,
    borderRadius: tokens.borderRadiusMedium,
    backgroundColor: tokens.colorNeutralBackground2,
  },
  formula: {
    fontFamily: tokens.fontFamilyMonospace,
    fontSize: tokens.fontSizeBase300,
    lineHeight: tokens.lineHeightBase300,
    overflowWrap: 'anywhere',
    margin: 0,
    whiteSpace: 'pre-wrap',
  },
  partText: { fontWeight: tokens.fontWeightSemibold },
  parts: { listStyle: 'none', margin: 0, padding: 0, display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalS },
  part: { display: 'flex', flexDirection: 'column', gap: '1px', paddingLeft: tokens.spacingHorizontalS, borderLeft: '3px solid' },
  partHead: { display: 'flex', alignItems: 'center', gap: tokens.spacingHorizontalS, minWidth: 0 },
  partCode: { fontFamily: tokens.fontFamilyMonospace, fontSize: tokens.fontSizeBase200, fontWeight: tokens.fontWeightSemibold, overflowWrap: 'anywhere', flex: 1 },
  pointers: { display: 'flex', flexWrap: 'wrap', gap: tokens.spacingHorizontalS },
  pointerNote: {
    padding: `${tokens.spacingVerticalXS} ${tokens.spacingHorizontalS}`,
    borderRadius: tokens.borderRadiusMedium,
    backgroundColor: tokens.colorBrandBackground2,
    color: tokens.colorNeutralForeground1,
  },
  status: { display: 'flex', alignItems: 'flex-start', gap: tokens.spacingHorizontalS, color: tokens.colorNeutralForeground3 },
  dot: {
    width: '8px',
    height: '8px',
    marginTop: '5px',
    flexShrink: 0,
    borderRadius: '50%',
    backgroundColor: tokens.colorBrandBackground,
    animationName: pulse,
    animationDuration: '1.4s',
    animationIterationCount: 'infinite',
    '@media (prefers-reduced-motion: reduce)': { animationName: 'none' },
  },
  doneLine: { display: 'flex', alignItems: 'center', gap: tokens.spacingHorizontalXS, color: tokens.colorPaletteGreenForeground1, fontWeight: tokens.fontWeightSemibold },
  note: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalXS,
    padding: tokens.spacingHorizontalS,
    borderRadius: tokens.borderRadiusMedium,
    backgroundColor: tokens.colorPaletteMarigoldBackground1,
    border: `1px solid ${tokens.colorPaletteMarigoldBorder1}`,
  },
  noteHead: { display: 'flex', gap: tokens.spacingHorizontalXS, alignItems: 'flex-start' },
  noteIcon: { color: tokens.colorPaletteMarigoldForeground1, display: 'inline-flex', marginTop: '1px' },
  compare: { display: 'grid', gridTemplateColumns: 'auto minmax(0, 1fr)', columnGap: tokens.spacingHorizontalS, rowGap: '2px', alignItems: 'baseline' },
  compareCode: { fontFamily: tokens.fontFamilyMonospace, fontSize: tokens.fontSizeBase200, overflowWrap: 'anywhere' },
  nav: { display: 'flex', gap: tokens.spacingHorizontalS, alignItems: 'center' },
  grow: { flex: 1 },
});

interface Props {
  steps: GuideStep[];
  index: number;
  onIndex(index: number): void;
  /** True while the rep's sheet is ready and nothing else is talking to Excel. */
  watching: boolean;
  /** The rep has passed its check. */
  passed: boolean;
  /** Where the answer goes, for "watching" lines: "I2:I5", or a Table column. */
  answerAt: string;
  probe(done: StepDone): Promise<StepProbe>;
  onShow(at: SheetSpot): void;
  onClose(): void;
  localize(text: string): string;
}

function waitingText(done: StepDone, answerAt: string): string {
  switch (done.kind) {
    case 'select':
      return `Waiting for you to click \`${done.range}\`.`;
    case 'answer':
      return `Watching \`${done.cells ?? answerAt}\`. Excel shows your work to the coach once you press {enter}.`;
    case 'inspect':
    case 'tableAt':
      return 'Watching the workbook. This step ticks itself off when it’s done.';
    case 'check':
      return 'Click **Check** below when you’re ready.';
  }
}

/** A part as its own chip: without the comma or bracket that joins it on, unless that's all there is (", ," marks an empty argument). */
function chipText(text: string): string {
  const bare = text.trim().replace(/^[,(]\s*/, '');
  return /[^\s,()]/.test(bare) ? bare : text.trim();
}

function FormulaCard({ parts, onShow, localize }: { parts: FormulaPart[]; onShow(at: SheetSpot): void; localize(t: string): string }) {
  const s = useStyles();
  let color = 0;
  const colored = parts.map((p) => (p.means ? PART_COLORS[color++ % PART_COLORS.length] : undefined));
  return (
    <div className={s.card}>
      <pre className={s.formula} aria-label={`Formula: ${partsFormula(parts)}`}>
        {parts.map((p, i) =>
          colored[i] ? (
            <span key={i} className={s.partText} style={{ color: colored[i] }}>
              {p.text}
            </span>
          ) : (
            <span key={i}>{p.text}</span>
          ),
        )}
      </pre>
      <ul className={s.parts}>
        {parts.map((p, i) =>
          p.means ? (
            <li key={i} className={s.part} style={{ borderLeftColor: colored[i] }}>
              <div className={s.partHead}>
                <span className={s.partCode} style={{ color: colored[i] }}>
                  {chipText(p.text)}
                </span>
                {p.at && (
                  <Button size="small" appearance="subtle" icon={<Eye16Regular />} onClick={() => onShow(p.at!)} aria-label={`Show ${p.text.trim()} on the sheet`}>
                    Show
                  </Button>
                )}
              </div>
              <Caption1>
                <RichText text={localize(p.means)} />
              </Caption1>
            </li>
          ) : null,
        )}
      </ul>
    </div>
  );
}

function Pointers({ pointers, onShow, localize }: { pointers: SheetPointer[]; onShow(at: SheetSpot): void; localize(t: string): string }) {
  const s = useStyles();
  const [shown, setShown] = useState<number>();
  const note = shown !== undefined ? pointers[shown]?.note : undefined;
  return (
    <>
      <div className={s.pointers}>
        {pointers.map((p, i) => (
          <Button
            key={i}
            size="small"
            icon={<Eye16Regular />}
            onClick={() => {
              setShown(i);
              onShow(p.at);
            }}
          >
            {localize(p.label)}
          </Button>
        ))}
      </div>
      {note && (
        <Caption1 className={s.pointerNote} role="status">
          <RichText text={localize(note)} />
        </Caption1>
      )}
    </>
  );
}

export function GuidePanel({ steps, index, onIndex, watching, passed, answerAt, probe, onShow, onClose, localize }: Props) {
  const s = useStyles();
  const current = Math.min(index, steps.length - 1);
  const step = steps[current];
  const [seen, setSeen] = useState<StepProbe>();
  // Steps the learner went back to: they show as done but don't move on by themselves.
  const revisited = useRef(new Set<number>());
  // The parent re-renders every second (its timer); keep the watch loop from restarting with it.
  const probeRef = useRef(probe);
  probeRef.current = probe;
  const onIndexRef = useRef(onIndex);
  onIndexRef.current = onIndex;

  useEffect(() => setSeen(undefined), [current]);

  // Look at the workbook until the step is done, then open the next one.
  useEffect(() => {
    const done = step?.done;
    if (!watching || passed || !done || done.kind === 'check') return;
    let live = true;
    let timer: ReturnType<typeof setTimeout> | undefined;
    const tick = async () => {
      try {
        const r = await probeRef.current(done);
        if (!live) return;
        setSeen(r);
        if (r.done) {
          if (current < steps.length - 1 && !revisited.current.has(current)) timer = setTimeout(() => live && onIndexRef.current(current + 1), ADVANCE_MS);
          return;
        }
      } catch {
        // Excel is busy (a dialog, or another request); look again shortly.
      }
      if (live) timer = setTimeout(() => void tick(), POLL_MS);
    };
    void tick();
    return () => {
      live = false;
      if (timer) clearTimeout(timer);
    };
  }, [watching, passed, step, current, steps.length]);

  const go = (i: number) => {
    if (i < current) revisited.current.add(i);
    onIndex(i);
  };

  const stepDone = passed || (seen?.done ?? false);
  const last = current === steps.length - 1;
  const ours = step?.formula ? partsFormula(step.formula) : undefined;

  return (
    <section className={s.guide} aria-label="Walkthrough">
      <div className={s.head}>
        <div className={s.headText}>
          <SectionLabel>Walkthrough</SectionLabel>
          <Caption1 className={s.count}>
            {passed ? 'Done' : `Step ${current + 1} of ${steps.length}`}
          </Caption1>
        </div>
        <Button size="small" appearance="subtle" icon={<Dismiss16Regular />} onClick={onClose} aria-label="Hide the walkthrough" />
      </div>
      <div className={s.bar} aria-hidden="true">
        {steps.map((_, i) => (
          <span key={i} className={mergeClasses(s.seg, (i < current || passed) && s.segDone, i === current && !passed && s.segNow)} />
        ))}
      </div>

      <ol className={s.steps}>
        {steps.map((st, i) => {
          const isNow = i === current && !passed;
          const isPast = i < current || passed;
          return (
            <li key={i} className={s.step} aria-current={isNow ? 'step' : undefined}>
              <span className={mergeClasses(s.marker, isNow && s.markerNow, isPast && s.markerDone)} role="img" aria-label={isPast ? `Step ${i + 1}, done` : `Step ${i + 1}`}>
                {isPast ? <CheckmarkCircle16Filled /> : i + 1}
              </span>
              <div className={mergeClasses(s.stepBody, isNow && s.current)}>
                {isNow ? (
                  <Text weight="semibold">
                    <RichText text={localize(st.do)} />
                  </Text>
                ) : (
                  <Body1 className={isPast ? s.past : s.future}>
                    <RichText text={localize(st.do)} />
                  </Body1>
                )}

                {isNow && (
                  <>
                    {st.formula && <FormulaCard parts={st.formula} onShow={onShow} localize={localize} />}
                    {st.why && (
                      <Caption1 className={s.why}>
                        <RichText text={localize(st.why)} />
                      </Caption1>
                    )}
                    {st.show && st.show.length > 0 && <Pointers key={i} pointers={st.show} onShow={onShow} localize={localize} />}

                    <div aria-live="polite">
                      {st.done && stepDone && st.done.kind !== 'check' ? (
                        <Caption1 className={s.doneLine}>
                          <Checkmark16Filled aria-hidden="true" /> Done
                        </Caption1>
                      ) : seen?.note ? (
                        <div className={s.note}>
                          <div className={s.noteHead}>
                            <span className={s.noteIcon} aria-hidden="true">
                              <Warning16Filled />
                            </span>
                            <Caption1>{seen.note}</Caption1>
                          </div>
                          {seen.formula && (
                            <div className={s.compare}>
                              <Caption1>Yours</Caption1>
                              <span className={s.compareCode}>{seen.formula}</span>
                              {ours && (
                                <>
                                  <Caption1>Guide</Caption1>
                                  <span className={s.compareCode}>{localize(ours)}</span>
                                </>
                              )}
                            </div>
                          )}
                          {seen.focus && (
                            <Link as="button" inline={false} onClick={() => onShow(seen.focus!)} style={{ alignSelf: 'flex-start', fontSize: tokens.fontSizeBase200 }}>
                              Go to {seen.focus}
                            </Link>
                          )}
                        </div>
                      ) : st.done ? (
                        <div className={s.status}>
                          {st.done.kind !== 'check' && watching && <span className={s.dot} aria-hidden="true" />}
                          <Caption1>
                            <RichText text={localize(waitingText(st.done, answerAt))} />
                          </Caption1>
                        </div>
                      ) : null}
                    </div>

                    <div className={s.nav}>
                      {current > 0 && (
                        <Button size="small" appearance="subtle" icon={<ArrowLeft16Regular />} onClick={() => go(current - 1)}>
                          Back
                        </Button>
                      )}
                      <span className={s.grow} />
                      {!last && (
                        <Button
                          size="small"
                          appearance={!st.done || stepDone ? 'primary' : 'secondary'}
                          icon={<ArrowRight16Regular />}
                          iconPosition="after"
                          onClick={() => go(current + 1)}
                        >
                          {!st.done || stepDone ? 'Next' : 'Skip ahead'}
                        </Button>
                      )}
                    </div>
                  </>
                )}
              </div>
            </li>
          );
        })}
      </ol>
    </section>
  );
}

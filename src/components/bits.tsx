import { makeStyles, mergeClasses, tokens } from '@fluentui/react-components';
import { CheckmarkCircle20Filled, Circle20Regular, CircleHalfFill20Regular, History20Regular } from '@fluentui/react-icons';
import { Fragment, useEffect, useState, type ReactNode } from 'react';
import { MASTERY_PASSES, type SkillStatus } from '../engine/progress';

// ---------- logo ----------

export function Logo({ size = 20 }: { size?: number }) {
  return (
    <svg width={size} height={size} viewBox="0 0 20 20" aria-hidden="true" focusable="false">
      <rect width="20" height="20" rx="4.5" fill="#107C41" />
      <path d="M6.7 0v20M13.3 0v20M0 6.7h20M0 13.3h20" stroke="#0B5C30" strokeWidth="0.6" />
      <path d="M5.3 10.5l3.1 3.1 6.3-6.9" fill="none" stroke="#fff" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round" />
    </svg>
  );
}

// ---------- skill status ----------

const useStatusStyles = makeStyles({
  icon: { flexShrink: 0, display: 'inline-flex' },
  new: { color: tokens.colorNeutralForeground4 },
  practicing: { color: tokens.colorBrandForeground1 },
  mastered: { color: tokens.colorBrandForeground1 },
  review: { color: tokens.colorPaletteMarigoldForeground1 },
});

const STATUS_LABEL: Record<SkillStatus, string> = {
  new: 'Not started',
  practicing: 'In progress',
  mastered: 'Mastered',
  review: 'Review due',
};

export function StatusIcon({ status }: { status: SkillStatus }) {
  const s = useStatusStyles();
  const Icon = { new: Circle20Regular, practicing: CircleHalfFill20Regular, mastered: CheckmarkCircle20Filled, review: History20Regular }[status];
  return (
    <span className={mergeClasses(s.icon, s[status])} role="img" aria-label={STATUS_LABEL[status]}>
      <Icon />
    </span>
  );
}

// ---------- mastery dots ----------

const pop = { '0%': { transform: 'scale(0.4)' }, '60%': { transform: 'scale(1.25)' }, '100%': { transform: 'scale(1)' } };

const useDotStyles = makeStyles({
  row: { display: 'inline-flex', gap: '5px', alignItems: 'center' },
  dot: {
    width: '9px',
    height: '9px',
    borderRadius: '50%',
    boxSizing: 'border-box',
    border: `1.5px solid ${tokens.colorNeutralStroke1}`,
  },
  filled: { backgroundColor: tokens.colorBrandBackground, border: `1.5px solid ${tokens.colorBrandBackground}` },
  fresh: {
    animationName: pop,
    animationDuration: '420ms',
    animationTimingFunction: 'cubic-bezier(0.2, 0.9, 0.3, 1.3)',
    '@media (prefers-reduced-motion: reduce)': { animationDuration: '1ms' },
  },
});

export function MasteryDots({ passes, fresh = -1 }: { passes: number; fresh?: number }) {
  const s = useDotStyles();
  return (
    <span className={s.row} role="img" aria-label={`${Math.min(passes, MASTERY_PASSES)} of ${MASTERY_PASSES} passes toward mastery`}>
      {Array.from({ length: MASTERY_PASSES }, (_, i) => (
        <span key={i} className={mergeClasses(s.dot, i < passes && s.filled, i === fresh && s.fresh)} />
      ))}
    </span>
  );
}

// ---------- inline code in copy ----------

const useRichStyles = makeStyles({
  code: {
    fontFamily: tokens.fontFamilyMonospace,
    fontSize: '0.9em',
    backgroundColor: tokens.colorNeutralBackground3,
    color: tokens.colorNeutralForeground1,
    padding: '1px 5px',
    borderRadius: tokens.borderRadiusSmall,
    whiteSpace: 'nowrap',
  },
});

/** Renders `backticked` spans as inline code. */
export function RichText({ text }: { text: string }) {
  const s = useRichStyles();
  const parts = text.split('`');
  return (
    <>
      {parts.map((part, i) =>
        i % 2 === 1 ? (
          <code key={i} className={s.code}>
            {part}
          </code>
        ) : (
          <Fragment key={i}>{part}</Fragment>
        ),
      )}
    </>
  );
}

// ---------- section label ----------

const useLabelStyles = makeStyles({
  label: {
    fontSize: tokens.fontSizeBase200,
    lineHeight: tokens.lineHeightBase200,
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground3,
    margin: 0,
  },
});

export function SectionLabel({ children, as: Tag = 'h2' }: { children: ReactNode; as?: 'h2' | 'h3' | 'p' }) {
  const s = useLabelStyles();
  return <Tag className={s.label}>{children}</Tag>;
}

// ---------- clock ----------

export function useNow(active: boolean, intervalMs = 1000): number {
  const [now, setNow] = useState(() => Date.now());
  useEffect(() => {
    if (!active) return;
    setNow(Date.now());
    const t = setInterval(() => setNow(Date.now()), intervalMs);
    return () => clearInterval(t);
  }, [active, intervalMs]);
  return now;
}

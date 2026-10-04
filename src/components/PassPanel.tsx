import { Caption1, Divider, Text, makeStyles, tokens } from '@fluentui/react-components';
import { MASTERY_PASSES, formatDuration, type PassOutcome } from '../engine/progress';
import type { CheckReport } from '../engine/types';
import { MasteryDots } from './bits';

const draw = { from: { strokeDashoffset: 48 }, to: { strokeDashoffset: 0 } };
const ring = { from: { strokeDashoffset: 151 }, to: { strokeDashoffset: 0 } };
const rise = { from: { opacity: 0, transform: 'translateY(6px)' }, to: { opacity: 1, transform: 'translateY(0)' } };
const reduce = { '@media (prefers-reduced-motion: reduce)': { animationDuration: '1ms', animationDelay: '0ms' } };

const useStyles = makeStyles({
  panel: {
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    gap: tokens.spacingVerticalS,
    padding: `${tokens.spacingVerticalL} ${tokens.spacingHorizontalM}`,
    borderRadius: tokens.borderRadiusLarge,
    backgroundColor: tokens.colorNeutralBackground2,
    textAlign: 'center',
  },
  ring: { strokeDasharray: 151, animationName: ring, animationDuration: '520ms', animationTimingFunction: tokens.curveDecelerateMax, animationFillMode: 'both', ...reduce },
  check: {
    strokeDasharray: 48,
    animationName: draw,
    animationDuration: '320ms',
    animationDelay: '360ms',
    animationTimingFunction: tokens.curveDecelerateMid,
    animationFillMode: 'both',
    ...reduce,
  },
  title: { animationName: rise, animationDuration: '300ms', animationDelay: '200ms', animationFillMode: 'both', ...reduce },
  sub: { color: tokens.colorNeutralForeground3 },
  divider: { width: '100%', margin: `${tokens.spacingVerticalXS} 0` },
  mastery: { display: 'flex', alignItems: 'center', justifyContent: 'space-between', width: '100%' },
  caption: { color: tokens.colorNeutralForeground2, textAlign: 'left', alignSelf: 'stretch' },
});

function Mark() {
  const s = useStyles();
  return (
    <svg width="56" height="56" viewBox="0 0 56 56" aria-hidden="true">
      <circle cx="28" cy="28" r="24" fill="none" stroke={tokens.colorNeutralStroke2} strokeWidth="3" />
      <circle className={s.ring} cx="28" cy="28" r="24" fill="none" stroke={tokens.colorBrandStroke1} strokeWidth="3" strokeLinecap="round" transform="rotate(-90 28 28)" />
      <path className={s.check} d="M18 28.5l6.8 6.8L38.5 21" fill="none" stroke={tokens.colorBrandForeground1} strokeWidth="3.4" strokeLinecap="round" strokeLinejoin="round" />
    </svg>
  );
}

interface Props {
  outcome?: PassOutcome;
  report: CheckReport;
  revealed: boolean;
  passes: number;
  mastered: boolean;
}

export function PassPanel({ outcome, report, revealed, passes, mastered }: Props) {
  const s = useStyles();
  const checks = report.items.filter((i) => i.status === 'pass').length;

  const sub = revealed
    ? 'The answer was shown, so this rep is practice only.'
    : outcome?.personalBest && outcome.passesTowardMastery > 1
      ? 'New personal best.'
      : outcome?.clean && (outcome.state.cleanStreak ?? 0) >= 2
        ? `${outcome.state.cleanStreak} clean passes in a row: first try, no hints.`
        : outcome?.clean
          ? 'First try, no hints.'
          : `All ${checks} checks passed.`;

  const remaining = MASTERY_PASSES - passes;
  const masteryCaption = outcome?.newlyMastered
    ? 'Mastered. It comes back for a quick review in 2 days, so it sticks.'
    : outcome?.wasReview
      ? 'Review done. The next one is further out.'
      : mastered
        ? 'Already mastered. Extra reps keep it sharp.'
        : `${remaining} more ${remaining === 1 ? 'pass' : 'passes'} on new data to master this skill.`;

  return (
    <section className={s.panel} role="status" aria-live="polite">
      <Mark />
      <div className={s.title}>
        <Text as="p" weight="semibold" size={500} block>
          Passed{outcome ? ` in ${formatDuration(outcome.elapsedMs)}` : ''}
        </Text>
        <Caption1 className={s.sub}>{sub}</Caption1>
      </div>
      <Divider className={s.divider} />
      <div className={s.mastery}>
        <Caption1>Mastery</Caption1>
        <MasteryDots passes={passes} fresh={outcome && !revealed ? passes - 1 : -1} />
      </div>
      <Caption1 className={s.caption}>{masteryCaption}</Caption1>
    </section>
  );
}

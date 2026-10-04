import { Caption1, ProgressBar, Text, makeStyles, mergeClasses, tokens } from '@fluentui/react-components';
import { History16Regular } from '@fluentui/react-icons';
import type { CSSProperties } from 'react';
import type { ProgressState, WeekSummary as WeekStats } from '../engine/progress';

const MINUTE = 60_000;
const INITIALS = ['S', 'M', 'T', 'W', 'T', 'F', 'S'];

const grow = { from: { transform: 'scaleY(0)' }, to: { transform: 'scaleY(1)' } };

const useStyles = makeStyles({
  card: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalM,
    padding: `${tokens.spacingVerticalM} ${tokens.spacingHorizontalM}`,
    borderRadius: tokens.borderRadiusLarge,
    backgroundColor: tokens.colorNeutralBackground2,
  },
  top: { display: 'grid', gridTemplateColumns: 'minmax(0, 1fr) auto', alignItems: 'end', gap: tokens.spacingHorizontalM },
  stat: { display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalXXS, minWidth: 0 },
  label: { color: tokens.colorNeutralForeground3 },
  figure: { fontVariantNumeric: 'tabular-nums' },
  bars: { display: 'grid', gridTemplateColumns: 'repeat(7, 12px)', columnGap: '3px' },
  day: { display: 'flex', flexDirection: 'column', alignItems: 'center', gap: tokens.spacingVerticalXXS },
  track: { display: 'flex', alignItems: 'flex-end', height: '32px' },
  bar: {
    width: '6px',
    height: '2px',
    borderRadius: `${tokens.borderRadiusMedium} ${tokens.borderRadiusMedium} 1px 1px`,
    backgroundColor: tokens.colorNeutralStroke2,
    transformOrigin: 'bottom',
    animationName: grow,
    animationDuration: tokens.durationSlower,
    animationTimingFunction: tokens.curveDecelerateMid,
    animationDelay: 'calc(var(--day) * 30ms)',
    animationFillMode: 'both',
    '@media (prefers-reduced-motion: reduce)': { animationDuration: '1ms', animationDelay: '0ms' },
  },
  // At least 3:1 on the card in both themes; the empty-day stub above stays a faint baseline.
  practiced: { minHeight: '4px', backgroundColor: tokens.colorNeutralStrokeAccessible },
  todayBar: { backgroundColor: tokens.colorBrandForeground1 },
  initial: { fontSize: tokens.fontSizeBase100, lineHeight: tokens.lineHeightBase100, color: tokens.colorNeutralForeground3 },
  todayInitial: { color: tokens.colorBrandForeground1, fontWeight: tokens.fontWeightSemibold },
  tiles: {
    display: 'grid',
    gridTemplateColumns: 'repeat(2, minmax(0, 1fr))',
    paddingTop: tokens.spacingVerticalM,
    borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  tileLead: { paddingRight: tokens.spacingHorizontalM },
  tileSplit: { paddingLeft: tokens.spacingHorizontalM, borderLeft: `1px solid ${tokens.colorNeutralStroke2}` },
  value: { display: 'flex', flexWrap: 'wrap', alignItems: 'baseline', columnGap: tokens.spacingHorizontalSNudge, fontVariantNumeric: 'tabular-nums' },
  of: { color: tokens.colorNeutralForeground3 },
  delta: { color: tokens.colorBrandForeground1, fontWeight: tokens.fontWeightSemibold },
  dueIcon: { color: tokens.colorPaletteMarigoldForeground1, display: 'inline-flex', alignSelf: 'center' },
  meter: { marginTop: tokens.spacingVerticalXS },
});

/** "1 h 25 min", "35 min", "Under a minute". `long` spells the units out for screen readers. */
export function formatPracticeTime(ms: number, long = false): string {
  const minutes = Math.floor(Math.max(0, ms) / MINUTE);
  if (ms > 0 && minutes === 0) return long ? 'under a minute' : 'Under a minute';
  const h = Math.floor(minutes / 60);
  const m = minutes % 60;
  if (!long) return h === 0 ? `${m} min` : m === 0 ? `${h} h` : `${h} h ${m} min`;
  const unit = (n: number, word: string) => `${n} ${word}${n === 1 ? '' : 's'}`;
  return h === 0 ? unit(m, 'minute') : m === 0 ? unit(h, 'hour') : `${unit(h, 'hour')} ${unit(m, 'minute')}`;
}

/** The whole card in one sentence, for the region's accessible name. */
export function weekSummaryLabel(summary: WeekStats, total: number): string {
  const time = summary.practicedMs > 0 ? `You practiced ${formatPracticeTime(summary.practicedMs, true)} this week` : 'You haven’t practiced this week';
  const fresh = summary.masteredThisWeek > 0 ? ` (${summary.masteredThisWeek} this week)` : '';
  const due = summary.reviewsDue === 0 ? 'no reviews' : `${summary.reviewsDue} ${summary.reviewsDue === 1 ? 'review' : 'reviews'}`;
  return `${time}, have mastered ${summary.masteredTotal} of ${total} skills${fresh}, and have ${due} due.`;
}

/** The card appears once there's something to sum up: a pass, a finished mission or drill run, or practice time. */
export function hasProgress(state: ProgressState): boolean {
  return (
    Object.values(state.exercises).some((p) => p.lastPassAt !== undefined || p.passSeeds.length > 0) ||
    Object.values(state.missions ?? {}).some((m) => m.completions > 0) ||
    Object.values(state.drills ?? {}).some((d) => d.runs > 0) ||
    Object.values(state.practice ?? {}).some((ms) => ms > 0)
  );
}

/** "2026-10-04" → the local date it names. */
function dateOf(key: string): Date {
  const [y, m, d] = key.split('-').map(Number);
  return new Date(y, m - 1, d);
}

function dayTitle(key: string, ms: number, today: boolean): string {
  const day = today ? 'Today' : dateOf(key).toLocaleDateString('en-US', { weekday: 'short', month: 'short', day: 'numeric' });
  return `${day} · ${formatPracticeTime(ms)}`;
}

interface Props {
  summary: WeekStats;
  /** Skills in the curriculum, for "12 of 36". */
  total: number;
}

/** Home screen card: practice this week with a bar per day, skills mastered, and reviews due. */
export function WeekSummary({ summary, total }: Props) {
  const s = useStyles();
  const max = Math.max(...summary.days.map((d) => d.ms));
  const last = summary.days.length - 1;

  return (
    <section className={s.card} aria-label={weekSummaryLabel(summary, total)}>
      <div className={s.top}>
        <div className={s.stat}>
          <Caption1 className={s.label}>Practiced this week</Caption1>
          <Text size={500} weight="semibold" className={s.figure}>
            {formatPracticeTime(summary.practicedMs)}
          </Text>
        </div>
        <div className={s.bars} aria-hidden="true">
          {summary.days.map((d, i) => (
            <span key={i} className={s.day} title={dayTitle(d.key, d.ms, i === last)}>
              <span className={s.track}>
                <span
                  className={mergeClasses(s.bar, d.ms > 0 && s.practiced, i === last && s.todayBar)}
                  style={{ height: d.ms > 0 ? `${(d.ms / max) * 100}%` : undefined, '--day': i } as CSSProperties}
                />
              </span>
              <span className={mergeClasses(s.initial, i === last && s.todayInitial)}>{INITIALS[dateOf(d.key).getDay()]}</span>
            </span>
          ))}
        </div>
      </div>

      <div className={s.tiles}>
        <div className={mergeClasses(s.stat, s.tileLead)}>
          <Caption1 className={s.label}>Skills mastered</Caption1>
          <span className={s.value}>
            <span>
              <Text size={400} weight="semibold">
                {summary.masteredTotal}
              </Text>
              <Text className={s.of}> of {total}</Text>
            </span>
            {summary.masteredThisWeek > 0 && <Caption1 className={s.delta}>+{summary.masteredThisWeek} this week</Caption1>}
          </span>
          <ProgressBar className={s.meter} value={total ? summary.masteredTotal / total : 0} thickness="medium" aria-hidden="true" />
        </div>
        <div className={mergeClasses(s.stat, s.tileSplit)}>
          <Caption1 className={s.label}>Reviews due</Caption1>
          <span className={s.value}>
            {summary.reviewsDue > 0 && <History16Regular className={s.dueIcon} aria-hidden="true" />}
            <Text size={400} weight="semibold">
              {summary.reviewsDue}
            </Text>
          </span>
        </div>
      </div>
    </section>
  );
}

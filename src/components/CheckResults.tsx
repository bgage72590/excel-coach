import { Caption1, Link, Text, makeStyles, mergeClasses, tokens } from '@fluentui/react-components';
import { CheckmarkCircle16Filled, Circle16Regular, DismissCircle16Filled, Warning20Filled } from '@fluentui/react-icons';
import type { CheckItem, CheckReport } from '../engine/types';

const slideIn = { from: { opacity: 0, transform: 'translateY(4px)' }, to: { opacity: 1, transform: 'translateY(0)' } };

const useStyles = makeStyles({
  box: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalS,
    padding: tokens.spacingHorizontalM,
    borderRadius: tokens.borderRadiusLarge,
    backgroundColor: tokens.colorPaletteMarigoldBackground1,
    border: `1px solid ${tokens.colorPaletteMarigoldBorder1}`,
    animationName: slideIn,
    animationDuration: '220ms',
    animationTimingFunction: tokens.curveDecelerateMid,
    '@media (prefers-reduced-motion: reduce)': { animationDuration: '1ms' },
  },
  head: { display: 'flex', alignItems: 'center', gap: tokens.spacingHorizontalS },
  headIcon: { color: tokens.colorPaletteMarigoldForeground1, display: 'inline-flex' },
  list: { listStyle: 'none', margin: 0, padding: 0, display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalS },
  item: { display: 'grid', gridTemplateColumns: '16px minmax(0, 1fr)', gap: tokens.spacingHorizontalS, alignItems: 'start' },
  icon: { display: 'inline-flex', marginTop: '2px' },
  pass: { color: tokens.colorPaletteGreenForeground1 },
  fail: { color: tokens.colorPaletteRedForeground1 },
  skip: { color: tokens.colorNeutralForeground4 },
  body: { display: 'flex', flexDirection: 'column', gap: '2px', minWidth: 0 },
  label: { fontSize: tokens.fontSizeBase200, lineHeight: tokens.lineHeightBase200 },
  failLabel: { fontWeight: tokens.fontWeightSemibold },
  detail: { color: tokens.colorNeutralForeground2 },
  skipText: { color: tokens.colorNeutralForeground3 },
});

const ICON = { pass: CheckmarkCircle16Filled, fail: DismissCircle16Filled, skip: Circle16Regular };

function Item({ item, onGoTo }: { item: CheckItem; onGoTo(address: string): void }) {
  const s = useStyles();
  const Icon = ICON[item.status];
  return (
    <li className={s.item}>
      <span className={mergeClasses(s.icon, s[item.status])} aria-label={item.status === 'pass' ? 'Passed' : item.status === 'fail' ? 'Failed' : 'Not checked yet'} role="img">
        <Icon />
      </span>
      <div className={s.body}>
        <Text className={mergeClasses(s.label, item.status === 'fail' && s.failLabel, item.status === 'skip' && s.skipText)}>{item.label}</Text>
        {item.status === 'fail' && item.detail && <Caption1 className={s.detail}>{item.detail}</Caption1>}
        {item.status === 'skip' && item.detail && <Caption1 className={s.skipText}>{item.detail}</Caption1>}
        {item.status === 'fail' && item.focus && (
          <Link as="button" inline={false} onClick={() => onGoTo(item.focus!)} style={{ alignSelf: 'flex-start', fontSize: tokens.fontSizeBase200 }}>
            Go to {item.focus}
          </Link>
        )}
      </div>
    </li>
  );
}

export function CheckResults({ report, onGoTo }: { report: CheckReport; onGoTo(address: string): void }) {
  const s = useStyles();
  const counted = report.items.filter((i) => i.status !== 'skip');
  const passed = counted.filter((i) => i.status === 'pass').length;
  // Skips with their own reason (an API this Excel lacks) explain themselves; the rest wait on the values.
  const waiting = report.items.filter((i) => i.status === 'skip' && !i.detail).length;

  return (
    <section className={s.box} role="status" aria-live="polite">
      <div className={s.head}>
        <span className={s.headIcon} aria-hidden="true">
          <Warning20Filled />
        </span>
        <Text weight="semibold">Not yet</Text>
        <Caption1 className={s.detail}>
          {passed} of {counted.length} checks passed
        </Caption1>
      </div>
      <ul className={s.list}>
        {report.items.map((item) => (
          <Item key={item.id} item={item} onGoTo={onGoTo} />
        ))}
      </ul>
      {waiting > 0 && <Caption1 className={s.skipText}>The remaining checks run once the values are right.</Caption1>}
    </section>
  );
}

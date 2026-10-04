import { Button, Menu, MenuItem, MenuList, MenuPopover, MenuTrigger, Text, makeStyles, tokens } from '@fluentui/react-components';
import { ArrowReset20Regular, MoreHorizontal20Regular } from '@fluentui/react-icons';
import { useState } from 'react';
import { emptyProgress } from '../engine/progress';
import { useCoach } from '../taskpane/context';
import { Logo } from './bits';
import { ConfirmDialog } from './ConfirmDialog';

const useStyles = makeStyles({
  header: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    height: '44px',
    padding: `0 ${tokens.spacingHorizontalS} 0 ${tokens.spacingHorizontalL}`,
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    flexShrink: 0,
  },
  brand: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
    background: 'none',
    border: 'none',
    padding: 0,
    cursor: 'pointer',
    color: tokens.colorNeutralForeground1,
    borderRadius: tokens.borderRadiusMedium,
    ':focus-visible': { outline: `2px solid ${tokens.colorStrokeFocus2}`, outlineOffset: '2px' },
  },
});

export function Header() {
  const s = useStyles();
  const { goHome, setProgress } = useCoach();
  const [confirmReset, setConfirmReset] = useState(false);

  return (
    <header className={s.header}>
      <button className={s.brand} onClick={goHome} aria-label="Excel Coach home">
        <Logo />
        <Text weight="semibold" size={300}>
          Excel Coach
        </Text>
      </button>
      <Menu positioning="below-end">
        <MenuTrigger disableButtonEnhancement>
          <Button appearance="subtle" icon={<MoreHorizontal20Regular />} aria-label="More options" />
        </MenuTrigger>
        <MenuPopover>
          <MenuList>
            <MenuItem icon={<ArrowReset20Regular />} onClick={() => setConfirmReset(true)}>
              Reset progress
            </MenuItem>
          </MenuList>
        </MenuPopover>
      </Menu>
      <ConfirmDialog
        open={confirmReset}
        title="Reset your progress?"
        body="This clears mastery, personal bests and review dates for every skill. Practice sheets in your workbook stay where they are."
        confirmLabel="Reset progress"
        onCancel={() => setConfirmReset(false)}
        onConfirm={() => {
          setProgress(() => emptyProgress());
          setConfirmReset(false);
          goHome();
        }}
      />
    </header>
  );
}

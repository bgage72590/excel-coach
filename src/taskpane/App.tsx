import { FluentProvider, makeStyles, tokens } from '@fluentui/react-components';
import { useCallback, useEffect, useMemo, useState } from 'react';
import { localize as localizeFor } from '../engine/platform';
import { loadProgress, saveProgress, type ProgressState } from '../engine/progress';
import type { Platform } from '../engine/types';
import { ExcelHost } from '../excel/excelHost';
import type { CoachHost } from '../excel/host';
import { MockHost } from '../excel/mockHost';
import { DrillView } from '../components/DrillView';
import { ExerciseView } from '../components/ExerciseView';
import { FixView } from '../components/FixView';
import { Header } from '../components/Header';
import { Home } from '../components/Home';
import { MissionView } from '../components/MissionView';
import { ScanView } from '../components/ScanView';
import { CoachContext, type CoachContextValue, type FixTarget } from './context';
import { darkTheme, lightTheme, prefersDark } from './theme';

export interface Env {
  inExcel: boolean;
  platform: Platform;
  mock: boolean;
  excelApi113: boolean;
}

type View =
  | { name: 'home' }
  | { name: 'exercise'; id: string }
  | { name: 'mission'; id: string }
  | { name: 'scan' }
  | { name: 'fix'; target: FixTarget }
  | { name: 'drills' };

// Note: FluentProvider copies its className onto portal layers (tooltips, menus),
// so page-level styles live on the inner shell, not on the provider.
const useStyles = makeStyles({
  shell: { display: 'flex', flexDirection: 'column', height: '100%', backgroundColor: tokens.colorNeutralBackground1, color: tokens.colorNeutralForeground1 },
  scroll: { flex: 1, overflowY: 'auto', overflowX: 'hidden' },
});

/** Remounts the page when the view changes, so each one starts fresh. */
function viewKey(view: View): string {
  if (view.name === 'exercise' || view.name === 'mission') return `${view.name}:${view.id}`;
  if (view.name === 'fix') return `fix:${view.target.sheet.sheet}:${view.target.finding.id}`;
  return view.name;
}

function useDarkMode(): boolean {
  const [dark, setDark] = useState(prefersDark);
  useEffect(() => {
    const mq = matchMedia('(prefers-color-scheme: dark)');
    const onChange = () => setDark(prefersDark());
    mq.addEventListener('change', onChange);
    return () => mq.removeEventListener('change', onChange);
  }, []);
  return dark;
}

export function App({ env }: { env: Env }) {
  const styles = useStyles();
  const dark = useDarkMode();
  const [progress, setProgressState] = useState<ProgressState>(loadProgress);
  const [view, setView] = useState<View>({ name: 'home' });

  const localize = useCallback((text: string) => localizeFor(text, env.platform), [env.platform]);

  const host = useMemo<CoachHost | null>(() => {
    if (env.inExcel) return new ExcelHost(localize);
    if (env.mock) return new MockHost();
    return null;
  }, [env.inExcel, env.mock, localize]);

  const hostNotice = !host
    ? 'Open Excel Coach from the Home tab in Excel to set up practice sheets.'
    : env.inExcel && !env.excelApi113
      ? 'This version of Excel is missing features the coach uses. Update Excel to check your work.'
      : undefined;

  const setProgress = useCallback((update: (prev: ProgressState) => ProgressState) => {
    setProgressState((prev) => {
      const next = update(prev);
      saveProgress(next);
      return next;
    });
  }, []);

  const value: CoachContextValue = {
    host,
    platform: env.platform,
    hostNotice,
    progress,
    setProgress,
    localize,
    openExercise: (id) => setView({ name: 'exercise', id }),
    openMission: (id) => setView({ name: 'mission', id }),
    openScan: () => setView({ name: 'scan' }),
    openFix: (target) => setView({ name: 'fix', target }),
    openDrills: () => setView({ name: 'drills' }),
    goHome: () => setView({ name: 'home' }),
  };

  return (
    <FluentProvider theme={dark ? darkTheme : lightTheme} style={{ height: '100%' }}>
      <CoachContext.Provider value={value}>
        <div className={styles.shell}>
          <Header />
          <main className={styles.scroll} key={viewKey(view)}>
            {view.name === 'home' && <Home />}
            {view.name === 'exercise' && <ExerciseView id={view.id} />}
            {view.name === 'mission' && <MissionView id={view.id} />}
            {view.name === 'scan' && <ScanView />}
            {view.name === 'fix' && <FixView target={view.target} />}
            {view.name === 'drills' && <DrillView />}
          </main>
        </div>
      </CoachContext.Provider>
    </FluentProvider>
  );
}

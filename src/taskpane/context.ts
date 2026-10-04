import { createContext, useContext } from 'react';
import type { ProgressState } from '../engine/progress';
import type { Finding, SheetFormulas } from '../engine/scan';
import type { Platform } from '../engine/types';
import type { CoachHost } from '../excel/host';

/** What "Fix it in My Work" needs: the finding and the sheet it was found on, as scanned. */
export interface FixTarget {
  finding: Finding;
  sheet: SheetFormulas;
  workbook: string;
}

export interface CoachContextValue {
  host: CoachHost | null;
  platform: Platform;
  /** Why setup/check are unavailable, if they are. */
  hostNotice?: string;
  progress: ProgressState;
  setProgress(update: (prev: ProgressState) => ProgressState): void;
  localize(text: string): string;
  openExercise(id: string): void;
  openMission(id: string): void;
  openScan(): void;
  openFix(target: FixTarget): void;
  openDrills(): void;
  goHome(): void;
}

export const CoachContext = createContext<CoachContextValue | null>(null);

export function useCoach(): CoachContextValue {
  const ctx = useContext(CoachContext);
  if (!ctx) throw new Error('useCoach must be used inside CoachContext');
  return ctx;
}

import type { Exercise } from '../engine/types';

/**
 * A timed drill: a short run of single actions ("Freeze the top row", "Sort by Amount, largest
 * first"). Each item is a small Exercise (module 'drills') that the host sets up and checks like
 * any other, so the result is checked, never the keystrokes. The timer runs from the moment an
 * item's sheet is ready until its check passes; setup time doesn't count.
 */
export interface DrillItem {
  exercise: Exercise<any>;
  /** The fastest route per platform, shown after the item passes, e.g. { mac: '⌘⇧F', windows: 'Ctrl+Shift+L' }. */
  shortcut: { mac: string; windows: string };
}

export interface Drill {
  id: string;
  title: string;
  /** One line under the title. */
  blurb: string;
  items: DrillItem[];
  /** A good time for the whole run, in seconds, shown as the target. */
  parSeconds: number;
}

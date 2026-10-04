import type { Rng } from '../engine/rng';
import type { AnswerArea, Block, ExpectedGrid, InputWrite, Inspection, Rules, Variant } from '../engine/types';

/**
 * A mission is a realistic job task on one dataset, done in steps. Each step is checked like an
 * exercise (same grader, same variants), against the shared data generated from one seed.
 */
export interface MissionStep<D> {
  /** Short name shown in the step list, e.g. "Stock value by warehouse". */
  title: string;
  /** What to do, with `backticked` cell references. */
  task(d: D): string;
  hints: string[];
  solution(d: D): string;
  answer(d: D): AnswerArea;
  expected(d: D): ExpectedGrid;
  variants: Variant<D>[];
  rules?: Rules;
  inspections?(d: D): Inspection[];
  /**
   * Writes applied once, when this step begins: next week's rows arrive, a figure is corrected.
   * A mission that uses onStart keeps `inputs` empty and its steps variant-free, because restoring
   * after a variant would undo these writes. expected(d) for this step and later ones should
   * describe the data after the writes.
   */
  onStart?(d: D): InputWrite[];
  /** Shown above the task once onStart has run, e.g. "Six new shipments were added to the Shipments Table." */
  startNote?(d: D): string;
}

export interface MissionBrief {
  from: string;
  subject: string;
  /** A short email from a manager. Plain text; `backticks` render as cell references. */
  body: string;
}

export interface Mission<D = unknown> {
  id: string;
  title: string;
  role: 'ops' | 'finance' | 'sales';
  /** One line: the real deliverable this rebuilds. */
  summary: string;
  minutes: number;
  /** Skills it draws on, by exercise id, for "practice this first" links. */
  skills: string[];
  brief(d: D): MissionBrief;
  make(rng: Rng): D;
  /** All data and labels for every step. Step answer areas are styled automatically. */
  blocks(d: D): Block[];
  /** Every input region the checker may rewrite (shared by all steps). */
  inputs(d: D): InputWrite[];
  steps: MissionStep<D>[];
  ownsNames?: string[];
}

export function defineMission<D>(m: Mission<D>): Mission<D> {
  return m;
}

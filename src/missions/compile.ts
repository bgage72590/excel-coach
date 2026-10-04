import type { Exercise } from '../engine/types';
import type { Mission } from './types';

/** Turns step `index` of a mission into an Exercise the host can set up and check. */
export function missionStepExercise<D>(m: Mission<D>, index: number): Exercise<D> {
  const step = m.steps[index];
  return {
    id: `${m.id}#${index + 1}`,
    module: 'missions',
    title: step.title,
    replaces: m.summary,
    minutes: m.minutes,
    task: step.task,
    concept: { summary: m.summary, syntax: '' },
    hints: step.hints,
    solution: step.solution,
    make: m.make,
    layout: (d) => ({
      blocks: m.blocks(d),
      answer: step.answer(d),
      alsoStyle: m.steps.filter((_, i) => i !== index).map((s) => s.answer(d)),
    }),
    expected: step.expected,
    inputs: m.inputs,
    variants: step.variants,
    rules: step.rules,
    inspections: step.inspections,
    ownsNames: m.ownsNames,
  };
}

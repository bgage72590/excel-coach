import { describe, expect, it } from 'vitest';
import { EXERCISES, MODULES } from '../src/exercises';
import { exerciseSuite } from './helpers/suite';

describe('curriculum', () => {
  it('has unique ids and known modules', () => {
    const ids = EXERCISES.map((e) => e.id);
    expect(new Set(ids).size).toBe(ids.length);
    for (const e of EXERCISES) expect(MODULES.some((m) => m.id === e.module)).toBe(true);
  });
});

exerciseSuite(EXERCISES);

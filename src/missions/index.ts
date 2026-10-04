import { apAging, carrierCostReview, weeklyRefresh } from './extra';
import { FINANCE_MISSIONS } from './finance';
import { FORECAST_MISSIONS } from './forecast';
import { OPS_MISSIONS } from './ops';
import { SALES_MISSIONS } from './sales';
import type { Mission } from './types';

/** Missions in the order they're offered, grouped by role: ops, then finance, then sales. */
export const MISSIONS: Mission<any>[] = [
  ...OPS_MISSIONS,
  carrierCostReview,
  weeklyRefresh,
  ...FINANCE_MISSIONS,
  apAging,
  ...FORECAST_MISSIONS,
  ...SALES_MISSIONS,
];

export function getMission(id: string): Mission<any> | undefined {
  return MISSIONS.find((m) => m.id === id);
}

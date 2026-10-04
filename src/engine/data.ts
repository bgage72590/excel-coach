import { round, type Rng } from './rng';

// ---------- vocabulary (ops, finance, sales) ----------

export const WAREHOUSES = ['Reno', 'Dallas', 'Atlanta', 'Columbus'] as const;
export const CARRIERS = ['Northline', 'Coastal', 'Ridge Freight', 'Summit Haul'] as const;
export const VENDORS = ['Apex Supply', 'Birchwood Co', 'Cobalt Parts', 'Delta Packaging', 'Evergreen Mfg', 'Foundry Direct'] as const;
export const EXTRA_VENDORS = ['Granite Tools', 'Harbor Industrial', 'Ironclad Safety'] as const;
export const CATEGORIES = ['Packaging', 'Hardware', 'Electrical', 'Safety', 'Janitorial'] as const;
export const REGIONS = ['North', 'South', 'East', 'West'] as const;
export const PRODUCTS = ['Basic', 'Plus', 'Pro', 'Max'] as const;

export const REPS = [
  'Ava Patel', 'Ben Ortiz', 'Chloe Kim', 'Dev Shah', 'Elena Rossi', 'Felix Wu', 'Grace Lee', 'Hugo Silva',
] as const;
export const EXTRA_REPS = ['Iris Novak', 'Jonah Reyes'] as const;

export interface ItemDef {
  item: string;
  category: (typeof CATEGORIES)[number];
  cost: number;
}

export const ITEMS: readonly ItemDef[] = [
  { item: 'Pallet wrap', category: 'Packaging', cost: 18.5 },
  { item: 'Corner board', category: 'Packaging', cost: 0.42 },
  { item: 'Packing tape', category: 'Packaging', cost: 2.1 },
  { item: 'Shrink film', category: 'Packaging', cost: 24.0 },
  { item: 'Bubble mailer', category: 'Packaging', cost: 0.38 },
  { item: 'Kraft paper roll', category: 'Packaging', cost: 31.0 },
  { item: 'Air pillows', category: 'Packaging', cost: 46.0 },
  { item: 'Steel strapping', category: 'Packaging', cost: 52.0 },
  { item: 'Hex bolt M8', category: 'Hardware', cost: 0.16 },
  { item: 'Lock washer M8', category: 'Hardware', cost: 0.04 },
  { item: 'Shelf bracket', category: 'Hardware', cost: 3.4 },
  { item: 'Caster wheel', category: 'Hardware', cost: 7.9 },
  { item: 'Bin divider', category: 'Hardware', cost: 1.25 },
  { item: 'Hand truck', category: 'Hardware', cost: 89.0 },
  { item: 'Cable tie 8 in', category: 'Electrical', cost: 0.05 },
  { item: 'Extension cord', category: 'Electrical', cost: 14.75 },
  { item: 'LED shop light', category: 'Electrical', cost: 27.5 },
  { item: 'Scanner battery', category: 'Electrical', cost: 19.0 },
  { item: 'Label printer roll', category: 'Electrical', cost: 9.6 },
  { item: 'Nitrile gloves', category: 'Safety', cost: 11.2 },
  { item: 'Safety vest', category: 'Safety', cost: 6.8 },
  { item: 'Hard hat', category: 'Safety', cost: 15.4 },
  { item: 'Ear plugs', category: 'Safety', cost: 0.22 },
  { item: 'Spill kit', category: 'Safety', cost: 64.0 },
  { item: 'Floor tape', category: 'Safety', cost: 8.3 },
  { item: 'Mop head', category: 'Janitorial', cost: 5.6 },
  { item: 'Trash liner', category: 'Janitorial', cost: 0.31 },
  { item: 'Degreaser', category: 'Janitorial', cost: 12.9 },
  { item: 'Paper towels', category: 'Janitorial', cost: 22.0 },
  { item: 'Broom', category: 'Janitorial', cost: 13.5 },
];

export const GL_ACCOUNTS: readonly { no: number; name: string; type: string }[] = [
  { no: 4000, name: 'Product revenue', type: 'Revenue' },
  { no: 4100, name: 'Service revenue', type: 'Revenue' },
  { no: 5000, name: 'Cost of goods sold', type: 'COGS' },
  { no: 5100, name: 'Freight in', type: 'COGS' },
  { no: 5200, name: 'Inventory adjustments', type: 'COGS' },
  { no: 6000, name: 'Salaries and wages', type: 'Opex' },
  { no: 6100, name: 'Payroll taxes', type: 'Opex' },
  { no: 6200, name: 'Rent', type: 'Opex' },
  { no: 6300, name: 'Utilities', type: 'Opex' },
  { no: 6400, name: 'Repairs and maintenance', type: 'Opex' },
  { no: 6500, name: 'Software subscriptions', type: 'Opex' },
  { no: 6600, name: 'Travel', type: 'Opex' },
  { no: 6700, name: 'Office supplies', type: 'Opex' },
  { no: 6800, name: 'Freight out', type: 'Opex' },
  { no: 6900, name: 'Insurance', type: 'Opex' },
];

export const DEPARTMENTS = ['Operations', 'Warehouse', 'Procurement', 'Finance', 'Sales', 'Customer success', 'IT', 'HR'] as const;

export const FIRST_NAMES = [
  'Maya', 'Liam', 'Noah', 'Emma', 'Olivia', 'Lucas', 'Sofia', 'Ethan', 'Zoe', 'Mason', 'Aria', 'Leo', 'Nora', 'Owen',
  'Ruby', 'Caleb', 'Ivy', 'Jack', 'Hazel', 'Miles', 'Stella', 'Wyatt', 'Clara', 'Theo',
] as const;
export const LAST_NAMES = [
  'Novak', 'Garcia', 'Chen', 'Okafor', 'Murphy', 'Larsen', 'Haddad', 'Moreau', 'Tanaka', 'Silva', 'Kowalski', 'Bennett',
  'Duarte', 'Fischer', 'Nguyen', 'Rahman', 'Walsh', 'Ibrahim', 'Costa', 'Lindqvist',
] as const;

// ---------- Excel dates ----------

const EPOCH = Date.UTC(1899, 11, 30);
const DAY = 86_400_000;

/** Excel serial date for a calendar date (month is 1-12). */
export function serial(year: number, month: number, day: number): number {
  return Math.round((Date.UTC(year, month - 1, day) - EPOCH) / DAY);
}

export function fromSerial(s: number): { year: number; month: number; day: number } {
  const d = new Date(EPOCH + s * DAY);
  return { year: d.getUTCFullYear(), month: d.getUTCMonth() + 1, day: d.getUTCDate() };
}

/** Same as Excel's EOMONTH(date, 0). */
export function eomonth(s: number): number {
  const { year, month } = fromSerial(s);
  return serial(year, month + 1, 0);
}

export function monthStart(s: number): number {
  const { year, month } = fromSerial(s);
  return serial(year, month, 1);
}

export function monthKey(s: number): string {
  const { year, month } = fromSerial(s);
  return `${year}-${String(month).padStart(2, '0')}`;
}

export function randomDate(rng: Rng, from: number, to: number): number {
  return rng.int(from, to);
}

// ---------- Excel text semantics ----------

/** Excel TRIM: strips leading/trailing spaces and collapses inner runs of spaces. */
export function excelTrim(s: string): string {
  return s.replace(/ +/g, ' ').replace(/^ | $/g, '');
}

/** Excel PROPER: capitalizes each letter that follows a non-letter, lowercases the rest. */
export function excelProper(s: string): string {
  let out = '';
  let prevLetter = false;
  for (const ch of s) {
    const isLetter = /\p{L}/u.test(ch);
    out += isLetter ? (prevLetter ? ch.toLowerCase() : ch.toUpperCase()) : ch;
    prevLetter = isLetter;
  }
  return out;
}

// ---------- small helpers ----------

export function sum(values: readonly number[]): number {
  return values.reduce((a, b) => a + b, 0);
}

export function money(rng: Rng, min: number, max: number): number {
  return rng.float(min, max, 2);
}

export function skuCode(n: number): string {
  return `SKU-${1000 + n}`;
}

export function scale(values: readonly number[], rng: Rng, min = 0.6, max = 1.6, decimals = 2): number[] {
  return values.map((v) => round(v * rng.float(min, max, 3), decimals));
}

/** Compare text the way Excel's SORT does for plain ASCII names. */
export function excelTextCompare(a: string, b: string): number {
  const x = a.toLowerCase();
  const y = b.toLowerCase();
  return x < y ? -1 : x > y ? 1 : 0;
}

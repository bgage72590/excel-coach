import { FIRST_NAMES, ITEMS, LAST_NAMES, excelProper, excelTrim } from '../engine/data';
import type { Rng } from '../engine/rng';
import { cells, column, defineExercise, rangeWrite } from './common';

// ---------- Fix a messy name export ----------

interface Person {
  first: string;
  last: string;
  raw: string;
}

const casing = [(s: string) => s.toUpperCase(), (s: string) => s.toLowerCase(), (s: string) => s];
const spaces = (rng: Rng, max: number) => ' '.repeat(rng.int(0, max));

function person(rng: Rng): Person {
  const first = rng.pick(FIRST_NAMES);
  const last = rng.pick(LAST_NAMES);
  const raw = `${spaces(rng, 2)}${rng.pick(casing)(last)},${spaces(rng, 3)}${rng.pick(casing)(first)}${spaces(rng, 2)}`;
  return { first, last, raw };
}

interface NameData {
  people: Person[];
}

const people = (rng: Rng) => Array.from({ length: 20 }, () => person(rng));

export const cleanNames = defineExercise<NameData>({
  id: 'clean-names',
  module: 'cleaning',
  title: 'Fix a messy name export',
  replaces: 'Retyping names from a system export by hand',
  minutes: 5,
  task: () =>
    'Column `A` holds names exported as “LAST, first” with stray spaces and random capitals. In `B2:C21`, return a clean first name and last name for each, capitalized normally (“Maya”, “Novak”).',
  concept: {
    summary:
      'TEXTBEFORE and TEXTAFTER split text at a delimiter. TRIM removes extra spaces, and PROPER fixes the capitals. Nest them from the inside out.',
    syntax: '=PROPER(TRIM(TEXTAFTER(A2, ",")))',
    example: 'TEXTAFTER("  NOVAK,  maya ", ",") gives "  maya ". TRIM makes it "maya", and PROPER makes it "Maya".',
  },
  hints: [
    'The first name is everything after the comma: TEXTAFTER(A2, ",").',
    'Clean the result with TRIM, then fix the capitals with PROPER.',
    'The last name works the same way with TEXTBEFORE.',
  ],
  solution: () => 'B2: =PROPER(TRIM(TEXTAFTER(A2,",")))    C2: =PROPER(TRIM(TEXTBEFORE(A2,",")))',
  make: (rng) => ({ people: people(rng) }),
  layout: (d) => ({
    blocks: [
      cells('A1', [['Employee (export)', 'First name', 'Last name']], 'header'),
      cells('A2', column(d.people.map((p) => p.raw)), 'input'),
    ],
    answer: { kind: 'cells', range: 'B2:C21', consistency: 'columns' },
  }),
  expected: (d) =>
    d.people.map((p) => {
      const [before, after] = [p.raw.slice(0, p.raw.indexOf(',')), p.raw.slice(p.raw.indexOf(',') + 1)];
      return [excelProper(excelTrim(after)), excelProper(excelTrim(before))];
    }),
  inputs: (d) => [rangeWrite('A2:A21', column(d.people.map((p) => p.raw)))],
  variants: [{ label: 'a new export arrives', apply: (_d, rng) => ({ people: people(rng) }) }],
});

// ---------- Split item codes into columns ----------

interface Code {
  sku: string;
  item: string;
  bin: string;
}

function codes(rng: Rng): Code[] {
  return Array.from({ length: 20 }, () => ({
    sku: `SKU-${rng.int(100, 999)}`,
    item: rng.pick(ITEMS).item,
    bin: `R${rng.int(1, 24)}-${rng.pick(['A', 'B', 'C', 'D'])}${rng.int(1, 6)}`,
  }));
}

const raw = (c: Code) => `${c.sku}|${c.item}|${c.bin}`;

export const cleanSplit = defineExercise<{ codes: Code[] }>({
  id: 'clean-split',
  module: 'cleaning',
  title: 'Split item codes into columns',
  replaces: 'Text to Columns, run again every time the export refreshes',
  minutes: 3,
  task: () => 'Each code in column `A` packs a SKU, a description and a bin location, separated by `|`. Split them into `B2:D21`.',
  concept: {
    summary: 'TEXTSPLIT breaks text apart at a delimiter and spills the pieces across the row. Unlike Text to Columns, it updates when the source changes.',
    syntax: '=TEXTSPLIT(text, column_delimiter)',
    example: '=TEXTSPLIT("SKU-104|Pallet wrap|R3-B2", "|") returns SKU-104, Pallet wrap and R3-B2 in three cells.',
  },
  hints: ['One formula per row is enough; the result spills to the right.', '=TEXTSPLIT(A2, "|"), then fill down.'],
  solution: () => '=TEXTSPLIT(A2,"|")',
  make: (rng) => ({ codes: codes(rng) }),
  layout: (d) => ({
    blocks: [cells('A1', [['Item code', 'SKU', 'Description', 'Bin']], 'header'), cells('A2', column(d.codes.map(raw)), 'input')],
    answer: { kind: 'cells', range: 'B2:D21', consistency: 'columns', spillOk: true },
  }),
  expected: (d) => d.codes.map((c) => [c.sku, c.item, c.bin]),
  inputs: (d) => [rangeWrite('A2:A21', column(d.codes.map(raw)))],
  variants: [{ label: 'the export refreshes with new codes', apply: (_d, rng) => ({ codes: codes(rng) }) }],
});

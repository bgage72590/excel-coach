import { FIRST_NAMES, ITEMS, LAST_NAMES, excelProper, excelTrim } from '../engine/data';
import type { Rng } from '../engine/rng';
import { cells, column, defineExercise, rangeWrite } from './common';
import { checkStep, fillStep, part, raw as glue, typeStep } from './guides';

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
  guide: (d) => {
    const exported = d.people[0].raw;
    const cut = exported.indexOf(',');
    const after = excelTrim(exported.slice(cut + 1));
    const first = excelProper(after);
    const last = excelProper(excelTrim(exported.slice(0, cut)));
    const looks = excelTrim(exported);
    const extra = looks !== exported ? ', plus some stray spaces' : '';
    const fixed = after === first ? `, so “${first}” stays as it is (other rows need the fix)` : `, so “${after}” becomes “${first}”`;
    return [
      {
        do: 'Meet the data. Column `A` holds names the way the system exported them: the last name, a comma, then the first name.',
        why: 'The export also added stray spaces (at the ends and after the comma) and random capitals. Every name has exactly one comma, so the comma is the place to cut.',
        show: [
          { label: 'The export', at: 'A2:A21', note: 'Twenty names, all in the same “LAST, first” pattern.' },
          { label: 'Where the results go', at: 'B2:C21', note: 'First names go in column `B`, last names in column `C`.' },
        ],
      },
      {
        do: `See what row 2 should become: ${first} in \`B2\` and ${last} in \`C2\`.`,
        why: 'Each name needs three fixes: cut it at the comma, trim the spaces, then fix the capitals. One formula does all three, working from the inside out.',
        show: [
          {
            label: 'Look at A2',
            at: 'A2',
            note: `\`A2\` reads “${looks}”${extra}. The part before the comma is the last name; the part after it is the first name.`,
          },
        ],
      },
      typeStep({
        cell: 'B2',
        formula: [
          part('=PROPER(', `Fix 3, done last: capitals. Each word gets a capital first letter and lowercase after it${fixed}.`),
          part('TRIM(', 'Fix 2: spaces. Removes the spaces at the start and end, and squeezes any run of spaces inside down to one.'),
          part('TEXTAFTER(', 'Fix 1, done first: the cut. Returns the text after a delimiter, the character that marks where to cut.'),
          part('A2', 'The messy name in `A2`.', 'A2'),
          glue(', '),
          part('","', 'The delimiter: a comma, in double quotes because it’s text. Everything after it is the first name.'),
          glue(')))'),
        ],
        why: 'Excel works from the innermost brackets out: TEXTAFTER cuts, TRIM cleans what it cut, and PROPER fixes the capitals. The three `)` at the end close the three functions.',
      }),
      typeStep({
        cell: 'C2',
        do: 'Click `C2` and type this formula, then press {enter}. It’s the `B2` formula with TEXTBEFORE in place of TEXTAFTER.',
        formula: [
          part('=PROPER(TRIM(', 'The same two fixes as in `B2`: trim the spaces, then fix the capitals.'),
          part('TEXTBEFORE(', 'This time keep the text before the delimiter: the last name.'),
          part('A2', 'The same messy name in `A2`.', 'A2'),
          glue(', '),
          part('","', 'Cut at the comma again.'),
          glue(')))'),
        ],
        why: `\`C2\` should show ${last}.`,
      }),
      {
        ...fillStep({
          from: 'B2',
          range: 'B2:C21',
          direction: 'down',
          why: 'Fill Down copies both formulas at once: `B2` down column `B` and `C2` down column `C`. Excel moves `A2` along to `A3`, `A4` and on down, so each row cleans its own name.',
        }),
        do: 'Select `B2:C21` (click `B2`, then Shift-click `C21`) and press {fillDown}.',
      },
      checkStep('The coach swaps in a fresh export behind the scenes, with new names, spaces and capitals, then puts yours back. Your formulas clean whatever arrives.'),
    ];
  },
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
  guide: (d) => {
    const c = d.codes[0];
    return [
      {
        do: 'Meet the data. Each code in column `A` packs three things into one cell: a SKU, a description and a bin location.',
        why: 'The pieces are separated by `|`, the vertical bar. That character is the delimiter: the marker that says where one piece ends and the next begins.',
        show: [
          { label: 'The codes', at: 'A2:A21', note: 'Twenty codes, each in the same SKU|description|bin pattern.' },
          { label: 'Where the pieces go', at: 'B1:D1', note: 'SKU, Description and Bin: one column per piece.' },
        ],
      },
      {
        do: `See what row 2 should become: ${c.sku}, ${c.item} and ${c.bin} in \`B2\`, \`C2\` and \`D2\`.`,
        why: 'Text to Columns would do this once. TEXTSPLIT does it with a formula, so when the export refreshes, the pieces update too.',
        show: [
          { label: 'Look at A2', at: 'A2', note: `\`A2\` holds \`${raw(c)}\`. Cut it at each \`|\` and you get the three pieces.` },
          { label: 'Where row 2 goes', at: 'B2:D2', note: 'One formula in `B2` fills all three of these cells.' },
        ],
      },
      {
        ...typeStep({
          cell: 'B2',
          formula: [
            part('=TEXTSPLIT(', 'Cuts text at a delimiter and puts each piece in its own cell, left to right.'),
            part('A2', 'The code to split: `A2`.', 'A2'),
            glue(', '),
            part('"|"', 'The delimiter: the vertical bar, in double quotes because it’s text. On a US keyboard, type it with Shift and the backslash key (`\\`).'),
            glue(')'),
          ],
          why: 'Press {enter} once. The three pieces spill into `B2:D2`, meaning Excel fills the cells to the right by itself. Keep `C2` and `D2` empty: anything typed there blocks the spill, and `B2` shows a spill error instead.',
        }),
        done: { kind: 'answer', cells: 'B2:D2' },
      },
      {
        ...fillStep({
          from: 'B2',
          range: 'B2:B21',
          direction: 'down',
          why: 'Each row’s formula spills its own three pieces into columns `C` and `D`. Leave those columns out of the selection: they fill themselves, and anything copied into them blocks the spills.',
        }),
        do: 'Select `B2:B21` (click `B2`, then Shift-click `B21`) and press {fillDown}.',
      },
      checkStep('The coach swaps in a refreshed export behind the scenes, then puts yours back. Text to Columns would leave the old pieces behind; your formulas split the new codes.'),
    ];
  },
  make: (rng) => ({ codes: codes(rng) }),
  layout: (d) => ({
    blocks: [cells('A1', [['Item code', 'SKU', 'Description', 'Bin']], 'header'), cells('A2', column(d.codes.map(raw)), 'input')],
    answer: { kind: 'cells', range: 'B2:D21', consistency: 'columns', spillOk: true },
  }),
  expected: (d) => d.codes.map((c) => [c.sku, c.item, c.bin]),
  inputs: (d) => [rangeWrite('A2:A21', column(d.codes.map(raw)))],
  variants: [{ label: 'the export refreshes with new codes', apply: (_d, rng) => ({ codes: codes(rng) }) }],
});

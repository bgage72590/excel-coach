import { describe, expect, it } from 'vitest';
import { colToNumber, numberToCol, parseCell, parseRange } from '../src/engine/address';
import { FIXABLE, MAX_FIX_ROWS, copyReferences, gradeFix, planFix, relativeKey, scanEdgeBelow, type FixPlan, type RangeRead } from '../src/engine/fix';
import { scanWorkbook, type Finding, type SheetFormulas } from '../src/engine/scan';
import type { Cell, CheckItem, Grid } from '../src/engine/types';

// ---------- Test helpers (the same synthetic sheets as the scanner's tests) ----------

type Cells = Record<string, Cell>;

/** Converts the A1 references in a formula to R1C1 as seen from (row, col), the way Excel reports .formulasR1C1. */
function toR1C1(formula: string, row: number, col: number): string {
  const rowPart = (abs: string, r: number) => (abs ? `R${r}` : r === row ? 'R' : `R[${r - row}]`);
  const colPart = (abs: string, c: number) => (abs ? `C${c}` : c === col ? 'C' : `C[${c - col}]`);
  const convert = (segment: string) =>
    segment.replace(
      /(?<![A-Za-z0-9_.$\]])(?:(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})|(\$?)([A-Za-z]{1,3})(\$?)(\d+))(?![A-Za-z0-9_(])/g,
      (_m, a1, c1, a2, c2, ac, c, ar, r) =>
        c1 !== undefined
          ? `${colPart(a1, colToNumber(c1))}:${colPart(a2, colToNumber(c2))}`
          : `${rowPart(ar, Number(r))}${colPart(ac, colToNumber(c))}`,
    );
  let out = '';
  let plain = '';
  for (let i = 0; i < formula.length; i++) {
    const ch = formula[i];
    if (ch === '"' || ch === "'" || ch === '[') {
      out += convert(plain);
      plain = '';
      let j = i + 1;
      if (ch === '[') {
        let depth = 1;
        while (j < formula.length && depth > 0) {
          if (formula[j] === '[') depth++;
          if (formula[j] === ']') depth--;
          j++;
        }
      } else {
        while (j < formula.length && formula[j] !== ch) j++;
        j++;
      }
      out += formula.slice(i, j);
      i = j - 1;
    } else {
      plain += ch;
    }
  }
  return out + convert(plain);
}

/** Builds the used range of a sheet from a map of cell → value or formula. */
function sheet(name: string, cells: Cells): SheetFormulas {
  const refs = Object.keys(cells).map((a) => ({ a, ...parseCell(a) }));
  const top = Math.min(...refs.map((r) => r.row));
  const left = Math.min(...refs.map((r) => r.col));
  const bottom = Math.max(...refs.map((r) => r.row));
  const right = Math.max(...refs.map((r) => r.col));
  const formulas: Grid = [];
  const r1c1: Grid = [];
  for (let r = top; r <= bottom; r++) {
    formulas.push(new Array(right - left + 1).fill(''));
    r1c1.push(new Array(right - left + 1).fill(''));
  }
  for (const ref of refs) {
    const v = cells[ref.a];
    formulas[ref.row - top][ref.col - left] = v;
    r1c1[ref.row - top][ref.col - left] = typeof v === 'string' && v.startsWith('=') ? toR1C1(v, ref.row, ref.col) : v;
  }
  return { sheet: name, address: `${numberToCol(left)}${top}:${numberToCol(right)}${bottom}`, formulas, r1c1 };
}

/** Fills a column from row `from` to `to` (inclusive). */
function fill(col: string, from: number, to: number, f: (r: number) => Cell): Cells {
  const out: Cells = {};
  for (let r = from; r <= to; r++) out[`${col}${r}`] = f(r);
  return out;
}

/** Order lines: A = SKU, B = qty, C = unit price, rows 2..(n+1). */
function orders(n = 20): Cells {
  return {
    A1: 'SKU',
    B1: 'Qty',
    C1: 'Unit price',
    ...fill('A', 2, n + 1, (r) => `SKU-${1000 + r}`),
    ...fill('B', 2, n + 1, (r) => (r * 7) % 40),
    ...fill('C', 2, n + 1, (r) => 3 + (r % 9) * 1.25),
  };
}

/** An item master in H:J (SKU, Vendor, Lead days) for lookups, rows 2..40. */
function items(): Cells {
  return {
    H1: 'SKU',
    I1: 'Vendor',
    J1: 'Lead days',
    ...fill('H', 2, 40, (r) => `SKU-${1000 + r}`),
    ...fill('I', 2, 40, (r) => ['Apex Supply', 'Birchwood Co', 'Cobalt Parts'][r % 3]),
    ...fill('J', 2, 40, (r) => 3 + (r % 11)),
  };
}

function findingFor(s: SheetFormulas, id: string, ...more: SheetFormulas[]): Finding {
  const f = scanWorkbook([s, ...more]).find((x) => x.id === id);
  if (!f) throw new Error(`The scanner found no ${id} on ${s.sheet}`);
  return f;
}

function planFor(s: SheetFormulas, id: string): FixPlan {
  const p = planFix(findingFor(s, id), s);
  if (!p) throw new Error(`No plan for ${id}`);
  return p;
}

/** A hand-made finding, for pointing at a specific cell. */
function finding(id: string, sheetName: string, ...addresses: string[]): Finding {
  return { id, severity: 'medium', title: 't', why: 'w', fix: 'f', count: addresses.length, cells: addresses.map((address) => ({ sheet: sheetName, address })) };
}

/** Reads a plan's range from a cell map: values from `values`, formulas from `cells` (typed values read as themselves). */
function readRange(range: string, cells: Cells, values: (address: string) => Cell): RangeRead {
  const r = parseRange(range);
  const formulas: Grid = [];
  const vals: Grid = [];
  for (let row = r.start.row; row <= r.end.row; row++) {
    const f: Cell[] = [];
    const v: Cell[] = [];
    for (let col = r.start.col; col <= r.end.col; col++) {
      const address = `${numberToCol(col)}${row}`;
      f.push(cells[address] ?? '');
      v.push(values(address));
    }
    formulas.push(f);
    vals.push(v);
  }
  return { address: range, values: vals, formulas };
}

const byId = (items: CheckItem[], id: string) => items.find((i) => i.id === id);
const passed = (items: CheckItem[]) => items.every((i) => i.status === 'pass');

/** Copy rules for a check's details. Error values (#REF!) and sheet prefixes in formulas are syntax, not prose. */
function expectPlainDetails(items: CheckItem[]) {
  for (const item of items) {
    const prose = `${item.label} ${item.detail ?? ''}`.replace(/#[A-Z/0]+[!?A]?/g, '').replace(/'(?:[^']|'')+'!/g, '');
    expect(prose, item.detail).not.toMatch(/\b(please|simply|just|easy|easily|successfully|leverage|seamless)\b|!/i);
    expect(prose, item.detail).not.toMatch(/'/);
  }
}

/** Copy rules from the brief, checked on every plan the tests build. */
function expectHouseStyle(p: FixPlan) {
  // Formula syntax is exempt: code in backticks, and quoted sheet names inside a formula in check advice.
  const outsideCode = (text: string) => text.replace(/`[^`]*`/g, '').replace(/'(?:[^']|'')+'!/g, '');
  expect(p.steps.length, p.title).toBeGreaterThanOrEqual(3);
  expect(p.steps.length, p.title).toBeLessThanOrEqual(5);
  expect(p.title).toMatch(/^[A-Z0-9]/);
  expect(p.title).not.toMatch(/\.$/);
  for (const text of [p.title, p.intro, ...p.steps, ...(p.copyBack ? [p.copyBack] : []), ...[...p.forbid, ...p.require].flatMap((r) => [r.label, r.advice])]) {
    const prose = outsideCode(text);
    expect(prose, text).not.toMatch(/\b(please|simply|just|easy|easily|successfully|leverage|seamless)\b|!/i);
    // Typographic apostrophes only.
    expect(prose, text).not.toMatch(/'/);
  }
  for (const text of [p.intro, ...p.steps]) {
    expect(text, text).toMatch(/[.:]$/);
    // Cell references sit in backticks.
    expect(outsideCode(text), text).not.toMatch(/(?<![A-Za-z])\$?[A-Z]{1,3}\$?\d+\b/);
    // Backticks pair up.
    expect((text.match(/`/g) ?? []).length % 2, text).toBe(0);
  }
}

// ---------- Fixtures: realistic sheets with real scan findings ----------

/** Orders with a VLOOKUP column (E), an IFERROR lookup column (F) and the item master beside them. */
function lookupSheet(): SheetFormulas {
  return sheet('Orders', {
    ...orders(),
    E1: 'Lead days',
    ...fill('E', 2, 21, (r) => `=VLOOKUP(A${r},$H$2:$J$40,3,FALSE)`),
    F1: 'Vendor',
    ...fill('F', 2, 21, (r) => `=IFERROR(VLOOKUP(A${r},$H$2:$J$40,2,FALSE),"Missing")`),
    ...items(),
  });
}

// ---------- relativeKey ----------

describe('relativeKey', () => {
  it('gives every copy of one filled formula the same key', () => {
    expect(relativeKey('=B2*C2', 'D2')).toBe(relativeKey('=B9*C9', 'D9'));
    expect(relativeKey('=B9*C8', 'D9')).not.toBe(relativeKey('=B9*C9', 'D9'));
    expect(relativeKey('=VLOOKUP(A2,$H$2:$J$40,3,FALSE)', 'E2')).toBe(relativeKey('=vlookup( A7, $H$2:$J$40, 3, false )', 'E7'));
  });

  it('keeps absolute parts fixed and text literals exact', () => {
    expect(relativeKey('=B2*$H$1', 'D2')).toBe('R[0]C[-2]*R1C8');
    expect(relativeKey('=SUM($C:$C)+SUM(B:B)', 'D2')).toBe('SUM(C3:C3)+SUM(C[-2]:C[-2])');
    expect(relativeKey('="Unit B2 "&A2', 'D2')).toBe('"Unit B2 "&R[0]C[-3]');
    expect(relativeKey("='Q1 Ops'!B2+Items[Cost]", 'D2')).toBe("'Q1OPS'!R[0]C[-2]+ITEMS[COST]");
    expect(relativeKey('=LOG10(B2)', 'D2')).toBe('LOG10(R[0]C[-2])');
  });
});

// ---------- Plans, one per detector ----------

describe('planFix: lookups with a typed column number', () => {
  it('rewrites the column’s own VLOOKUP as XLOOKUP', () => {
    const p = planFor(lookupSheet(), 'vlookup-column-number');
    expect(p).toMatchObject({ findingId: 'vlookup-column-number', sheet: 'Orders', range: 'E2:E21', title: 'Replace VLOOKUP with XLOOKUP', valuesMayChange: [], allFormulas: false });
    expect(p.steps[0]).toBe('On the copy, select `E2`. Its VLOOKUP returns column 3 of `$H$2:$J$40`, counted by position.');
    expect(p.steps[1]).toContain('`=XLOOKUP(A2,$H$2:$H$40,$J$2:$J$40)`');
    expect(p.steps[2]).toBe('Fill it down: copy `E2`, select `E3:E21`, and paste.');
    expect(p.forbid.map((r) => r.label)).toEqual(['No VLOOKUP or HLOOKUP left']);
    expect(p.require.map((r) => r.label)).toEqual(['Uses XLOOKUP or INDEX and MATCH']);
    expectHouseStyle(p);
  });

  it('handles HLOOKUP, other sheets and lookups inside a larger formula', () => {
    const s = sheet('Plan', {
      A1: 'Month',
      ...fill('A', 2, 6, (r) => ['Jan', 'Feb', 'Mar', 'Apr', 'May'][r - 2]),
      B1: 'Rate',
      ...fill('B', 2, 6, (r) => `=ROUND(HLOOKUP(A${r},'Rate card'!$B$1:$M$9,4,0)*1.1,2)`),
    });
    const p = planFor(s, 'vlookup-column-number');
    expect(p.title).toBe('Replace HLOOKUP with XLOOKUP');
    expect(p.steps[1]).toContain("`=ROUND(XLOOKUP(A2,'Rate card'!$B$1:$M$1,'Rate card'!$B$4:$M$4)*1.1,2)`");
    expectHouseStyle(p);
  });

  it('falls back to general steps when the table isn’t a plain range', () => {
    const s = sheet('Orders', { ...orders(5), E1: 'Cost', ...fill('E', 2, 6, (r) => `=VLOOKUP(A${r},Items,4,FALSE)`) });
    const p = planFor(s, 'vlookup-column-number');
    expect(p.steps[1]).toMatch(/^Rewrite each VLOOKUP as `XLOOKUP\(lookup_value, lookup_array, return_array\)`/);
    expectHouseStyle(p);
  });

  it('skips approximate lookups but plans a later exact one on the sheet', () => {
    const s = sheet('Orders', {
      ...orders(),
      E1: 'Band',
      ...fill('E', 2, 21, (r) => `=VLOOKUP(B${r},$H$2:$J$40,2)`),
      F1: 'Lead days',
      ...fill('F', 2, 21, (r) => `=VLOOKUP(A${r},$H$2:$J$40,3,FALSE)`),
      ...items(),
    });
    const f = findingFor(s, 'vlookup-column-number');
    expect(f.cells[0].address).toBe('E2');
    expect(planFix(f, s)?.range).toBe('F2:F21');
    expect(planFix({ ...f, cells: f.cells.filter((c) => c.address.startsWith('E')) }, s)).toBeNull();
  });
});

describe('planFix: IFERROR around a lookup', () => {
  it('moves the fallback into XLOOKUP', () => {
    const p = planFor(lookupSheet(), 'iferror-lookup');
    expect(p.range).toBe('F2:F21');
    expect(p.steps[0]).toBe('On the copy, select `F2`. IFERROR shows `"Missing"` for any error, not only a missing key.');
    expect(p.steps[1]).toContain('`=XLOOKUP(A2,$H$2:$H$40,$I$2:$I$40,"Missing")`');
    expect(p.forbid).toHaveLength(1);
    const [rule] = p.forbid;
    expect(rule.label).toBe('No IFERROR around the lookup');
    expect(rule.test!('=IFERROR(VLOOKUP(A2,$H$2:$J$40,2,FALSE),"Missing")')).toBe(true);
    expect(rule.test!('=IFNA(VLOOKUP(A2,$H$2:$J$40,2,FALSE),"Missing")')).toBe(false);
    expect(rule.test!('=IFERROR(B2/C2,0)')).toBe(false);
    expect(rule.test!('=XLOOKUP(A2,$H$2:$H$40,$I$2:$I$40,"IFERROR(VLOOKUP(")')).toBe(false);
    expectHouseStyle(p);
  });

  it('uses IFNA when XLOOKUP can’t take the fallback, and XLOOKUP for an IFNA wrapper', () => {
    const indexMatch = sheet('Orders', { ...orders(5), E1: 'Vendor', ...fill('E', 2, 6, (r) => `=IFERROR(INDEX($I$2:$I$40,MATCH(A${r},$H$2:$H$40,0)),"")`) });
    const a = planFor(indexMatch, 'iferror-lookup');
    expect(a.steps[1]).toContain('`=IFNA(INDEX($I$2:$I$40,MATCH(A2,$H$2:$H$40,0)),"")`');
    expect(a.steps[0]).toContain('IFERROR shows a blank for any error');

    const ifna = sheet('Orders', { ...orders(5), E1: 'Cost', ...fill('E', 2, 6, (r) => `=IFNA(XLOOKUP(A${r},Items[SKU],Items[Cost]),0)`) });
    const b = planFor(ifna, 'iferror-lookup');
    expect(b.title).toBe('Move the fallback into XLOOKUP');
    expect(b.steps[1]).toContain('`=XLOOKUP(A2,Items[SKU],Items[Cost],0)`');
    expect(b.forbid[0].label).toBe('No IFERROR or IFNA around the lookup');
    expect(b.forbid[0].test!('=IFNA(XLOOKUP(A2,Items[SKU],Items[Cost]),0)')).toBe(true);
    expectHouseStyle(a);
    expectHouseStyle(b);
  });
});

describe('planFix: nested IFs', () => {
  const tiers = (formula: (r: number) => string) => sheet('Sales', { ...orders(), D1: 'Tier', ...fill('D', 2, 21, formula) });

  it('rewrites an else-chain as IFS', () => {
    const p = planFor(tiers((r) => `=IF(B${r}>30,"A",IF(B${r}>20,"B",IF(B${r}>10,"C","D")))`), 'nested-if');
    expect(p.range).toBe('D2:D21');
    expect(p.steps[1]).toContain('`=IFS(B2>30,"A",B2>20,"B",B2>10,"C",TRUE,"D")`');
    expectHouseStyle(p);
  });

  it('keeps Excel’s results for a missing or empty last branch', () => {
    const missing = planFor(tiers((r) => `=IF(B${r}>30,0.05,IF(B${r}>20,0.03,IF(B${r}>10,0.01)))`), 'nested-if');
    expect(missing.steps[1]).toContain('`=IFS(B2>30,0.05,B2>20,0.03,B2>10,0.01,TRUE,FALSE)`');
    const empty = planFor(tiers((r) => `=IF(B${r}>30,0.05,IF(B${r}>20,0.03,IF(B${r}>10,0.01,)))`), 'nested-if');
    expect(empty.steps[1]).toContain('TRUE,0)`');
  });

  it('counts depth without looking inside text', () => {
    const p = planFor(tiers((r) => `=IF(B${r}>30,"A",IF(B${r}>20,"B",IF(B${r}>10,"C","D")))`), 'nested-if');
    const [rule] = p.forbid;
    expect(rule.test!('=IF(B2>30,"A",IF(B2>20,"B",IF(B2>10,"C","D")))')).toBe(true);
    expect(rule.test!('=IFS(B2>30,"IF(IF(IF(",TRUE,"D")')).toBe(false);
    expect(rule.test!('=IF(B2>20,"B",IF(B2>10,"C","D"))')).toBe(false);
  });

  it('gives general steps for nesting it can’t turn into IFS', () => {
    const p = planFor(tiers((r) => `=IF(B${r}>0,IF(C${r}>5,IF(A${r}<>"","Y","N"),"N"),"N")`), 'nested-if');
    expect(p.steps[1]).toMatch(/^Rewrite the chain as `IFS\(/);
    expectHouseStyle(p);
  });
});

describe('planFix: SUMPRODUCT conditions', () => {
  const summary = (cells: Cells) => sheet('Summary', { ...orders(), G1: 'Region', G2: 'SKU-1004', H1: 'Min qty', H2: 10, ...cells });

  it('rewrites to SUMIFS with criteria built from comparisons', () => {
    const s = summary({ I1: 'Total', I2: '=SUMPRODUCT(($A$2:$A$21=G2)*($B$2:$B$21>=$H$2)*$C$2:$C$21)' });
    const p = planFor(s, 'sumproduct-conditions');
    expect(p.title).toBe('Replace SUMPRODUCT with SUMIFS');
    expect(p.steps[1]).toContain('`=SUMIFS($C$2:$C$21,$A$2:$A$21,G2,$B$2:$B$21,">="&$H$2)`');
    expectHouseStyle(p);
  });

  it('counts with COUNTIFS and flips reversed comparisons', () => {
    const s = summary({ I1: 'Lines', I2: '=SUMPRODUCT(--("SKU-1004"=$A$2:$A$21),--(25<$B$2:$B$21))' });
    const p = planFor(s, 'sumproduct-conditions');
    expect(p.title).toBe('Replace SUMPRODUCT with COUNTIFS');
    expect(p.steps[1]).toContain('`=COUNTIFS($A$2:$A$21,"SKU-1004",$B$2:$B$21,">25")`');
  });
});

describe('planFix: INDIRECT and OFFSET', () => {
  it('turns a dynamic OFFSET range into a range ending in INDEX', () => {
    const s = sheet('Ops', { ...orders(), E1: 'Total', E2: '=SUM(OFFSET($C$2,0,0,COUNTA($C:$C)-1,1))' });
    const p = planFor(s, 'volatile-references');
    expect(p.title).toBe('Replace OFFSET with INDEX');
    expect(p.steps[1]).toContain('`=SUM($C$2:INDEX($C:$C,COUNTA($C:$C)))`');
    expect(p.forbid.map((r) => r.label)).toEqual(['No OFFSET or INDIRECT left']);
    expectHouseStyle(p);
  });

  it('turns INDIRECT("B"&n) into INDEX', () => {
    const s = sheet('Ops', { ...orders(), F1: 'Row', ...fill('F', 2, 6, (r) => r + 3), G1: 'Qty', ...fill('G', 2, 6, (r) => `=INDIRECT("B"&F${r})`) });
    const p = planFor(s, 'volatile-references');
    expect(p.title).toBe('Replace INDIRECT with INDEX');
    expect(p.steps[1]).toContain('`=INDEX($B:$B,F2)`');
    expect(p.steps[2]).toBe('Fill it down: copy `G2`, select `G3:G6`, and paste.');
  });
});

describe('planFix: typed over a formula', () => {
  const extended = (r: number) => `=B${r}*C${r}`;

  it('takes the column through the typed value and allows only it to change', () => {
    const s = sheet('Ops', { ...orders(), D1: 'Extended', ...fill('D', 2, 21, extended), D9: 412.5 });
    const p = planFor(s, 'typed-over-formula');
    expect(p).toMatchObject({ range: 'D2:D21', valuesMayChange: ['D9'], allFormulas: true, forbid: [], require: [] });
    expect(p.restore).toMatchObject({ key: relativeKey('=B2*C2', 'D2'), label: 'Typed values replaced with the column’s formula' });
    expect(p.steps[0]).toBe('On the copy, select `D9`. It holds the typed number `412.5` where the rest of the Extended column has a formula.');
    expect(p.steps[2]).toBe('Copy `D8`, which has the column’s formula, and paste it into `D9`. Excel adjusts the references for each row.');
    expectHouseStyle(p);
  });

  it('includes a number typed over the last formula, and lists several', () => {
    const s = sheet('Ops', { ...orders(), D1: 'Extended', ...fill('D', 2, 21, extended), D7: 88, D12: 90.25, D21: 14 });
    const p = planFor(s, 'typed-over-formula');
    expect(p.range).toBe('D2:D21');
    expect(p.valuesMayChange).toEqual(['D7', 'D12', 'D21']);
    expect(p.steps[0]).toBe('On the copy, `D7`, `D12` and `D21` hold typed numbers where the rest of the Extended column has a formula.');
  });

  it('lets the rows below change in a running column', () => {
    const s = sheet('Cash', { A1: 'Day', B1: 'Net', ...fill('B', 2, 21, (r) => (r % 5) * 10 - 15), C1: 'Balance', C2: '=B2', ...fill('C', 3, 21, (r) => `=C${r - 1}+B${r}`), C10: 500 });
    const p = planFor(s, 'typed-over-formula');
    expect(p.valuesMayChange[0]).toBe('C10');
    expect(p.valuesMayChange).toContain('C21');
    expect(p.valuesMayChange).not.toContain('C9');
    expect(p.steps[3]).toContain('and so can the rows that build on them');
  });
});

describe('planFix: formulas that break their column’s pattern', () => {
  const extended = (r: number) => `=B${r}*C${r}`;

  it('asks for the majority formula in the odd cells only', () => {
    const s = sheet('Ops', { ...orders(), D1: 'Extended', ...fill('D', 2, 21, extended), D12: '=B12*C11', D16: '=B16*1.08' });
    const p = planFor(s, 'inconsistent-column');
    expect(p).toMatchObject({ range: 'D2:D21', valuesMayChange: ['D12', 'D16'], allFormulas: true });
    expect(p.restore?.key).toBe(relativeKey('=B2*C2', 'D2'));
    expect(p.steps[0]).toBe('On the copy, compare `D12` and `D16` with `D11`, which has the formula the rest of the column shares: `=B11*C11`.');
    expect(p.steps[2]).toBe('Copy `D11` and paste it into `D12` and `D16`. Excel adjusts the references for each row.');
    expectHouseStyle(p);
  });

  it('leaves a running balance’s seed and a total alone', () => {
    const s = sheet('Cash', {
      A1: 'Day',
      B1: 'Net',
      ...fill('B', 2, 21, (r) => (r % 5) * 10 - 15),
      C1: 'Balance',
      C2: '=B2',
      ...fill('C', 3, 21, (r) => `=C${r - 1}+B${r}`),
      C15: '=C13+B15',
      C22: '=MIN(C2:C21)',
    });
    const p = planFor(s, 'inconsistent-column');
    expect(p.range).toBe('C2:C21');
    expect(p.valuesMayChange[0]).toBe('C15');
    expect(p.valuesMayChange).not.toContain('C2');
    expect(p.valuesMayChange).toContain('C16');
  });
});

describe('planFix: a number typed into many formulas', () => {
  const taxed = () =>
    sheet('Ops', {
      ...orders(9),
      E1: 'With tax',
      ...fill('E', 2, 10, (r) => `=C${r}*1.0825`),
      G1: 'Total',
      G2: '=SUM(C2:C10)*1.0825',
      H1: 'Pre-tax',
      H2: '=B2/1.0825',
    });

  it('moves the number into a labeled input cell past the data', () => {
    const p = planFor(taxed(), 'hardcoded-constants');
    expect(p.title).toBe('Move 1.0825 into an input cell');
    expect(p.range).toBe('E2:E10');
    expect(p.steps[0]).toBe('On the copy, type `1.0825` in `K1`, with a short label in `J1` that says what it is.');
    expect(p.steps[1]).toContain('`=C2*$K$1`');
    expect(p.steps[1]).toContain('{absKey}');
    expect(p.copyBack).toContain('`J1:K1`');
    expect(p.forbid.map((r) => r.label)).toEqual(['No 1.0825 typed into the formulas']);
    expectHouseStyle(p);
  });

  it('forbids the number however it’s written, but not other numbers', () => {
    const [rule] = planFor(taxed(), 'hardcoded-constants').forbid;
    expect(rule.test!('=C2*1.0825')).toBe(true);
    expect(rule.test!('=C2*108.25%')).toBe(true);
    expect(rule.test!('=C2*$K$1')).toBe(false);
    expect(rule.test!('=ROUND(C2*$K$1,4)')).toBe(false);
    expect(rule.test!('=C2&" at 1.0825"')).toBe(false);
  });

  it('picks the flagged number when a formula types in several', () => {
    // 5 is a single digit filled down one formula, which the scanner treats as structure.
    const s = sheet('Ops', {
      ...orders(9),
      E1: 'Commission',
      ...fill('E', 2, 10, (r) => `=IF(B${r}>5,C${r}*0.05,0)`),
      G2: '=SUM(C2:C10)*0.05',
      H2: '=B2*0.05',
    });
    const p = planFor(s, 'hardcoded-constants');
    expect(p.title).toBe('Move 0.05 into an input cell');
    expect(p.forbid).toHaveLength(1);
    expect(p.steps[1]).toContain('`=IF(B2>5,C2*$K$1,0)`');
  });

  it('targets only the flagged number in a weekend-rate fill', () => {
    const s = sheet('Ops', {
      ...orders(5),
      E1: 'Pay',
      ...fill('E', 2, 6, (r) => `=IF(WEEKDAY(A${r},2)>5,B${r}*1.5,B${r})`),
    });
    const p = planFor(s, 'hardcoded-constants');
    expect(findingFor(s, 'hardcoded-constants').numbers).toEqual([1.5]);
    expect(p.title).toBe('Move 1.5 into an input cell');
    expect(p.forbid).toHaveLength(1);
  });

  it('targets only the rate in a rounding fill', () => {
    const s = sheet('Ops', {
      ...orders(9),
      E1: 'Billed',
      ...fill('E', 2, 10, (r) => `=INT(C${r}*1.0825+0.5)`),
    });
    expect(planFor(s, 'hardcoded-constants').title).toBe('Move 1.0825 into an input cell');
  });

  it('moves several flagged numbers into input cells, the most repeated first', () => {
    const s = sheet('Ops', {
      ...orders(9),
      E1: 'Commission',
      ...fill('E', 2, 10, (r) => `=IF(C${r}>5000,C${r}*0.05,0)`),
      G2: '=SUM(C2:C10)*0.05',
      H2: '=B2*0.05',
    });
    const p = planFor(s, 'hardcoded-constants');
    expect(p.title).toBe('Move the typed numbers into input cells');
    expect(p.forbid.map((r) => r.label)).toEqual(['No 0.05 typed into the formulas', 'No 5000 typed into the formulas']);
    expect(p.steps[0]).toBe('On the copy, type `0.05` in `K1` and `5000` in `K2`, each with a short label to its left that says what it is.');
    expect(p.steps[1]).toContain('`=IF(C2>$K$2,C2*$K$1,0)`');
    expect(p.copyBack).toContain('`J1:K2`');
    expectHouseStyle(p);
  });
});

describe('planFix: fixed ranges over long lists', () => {
  it('converts the data on the sheet to a Table and names its columns', () => {
    const s = sheet('Ops', { ...orders(39), E1: 'Total', E2: '=SUM($C$2:$C$40)', E3: '=SUMIFS($C$2:$C$40,$B$2:$B$40,">20")' });
    const p = planFor(s, 'fixed-long-ranges');
    expect(p.range).toBe('E2:E3');
    expect(p.steps[0]).toBe('On the copy, select `A1:C40` and press {tableKey}. Keep My table has headers selected.');
    expect(p.steps[1]).toBe('On the Table tab (Table Design on Windows), type `OpsData` in the Table Name box and press {enter}.');
    expect(p.steps[2]).toContain('`=SUM(OpsData[Unit price])`');
    expect(p.steps[3]).toBe('Rewrite the other formulas with fixed ranges in `E2:E3` the same way.');
    expect(p.require.map((r) => r.label)).toEqual(['Refers to the Table’s columns by name']);
    expect(p.copyBack).toContain('`A1:C40`');
    expectHouseStyle(p);
  });

  it('keeps a blank row inside the data inside the Table', () => {
    const cells: Cells = { ...orders(39), E1: 'Total', E2: '=SUM($C$2:$C$40)' };
    delete cells.A20;
    delete cells.B20;
    delete cells.C20;
    const p = planFor(sheet('Ops', cells), 'fixed-long-ranges');
    expect(p.steps[0]).toBe('On the copy, select `A1:C40` and press {tableKey}. Keep My table has headers selected.');
    expect(p.steps[2]).toContain('`=SUM(OpsData[Unit price])`');
    expect(p.copyBack).toContain('`A1:C40`');
  });

  it('ends the Table at the last row of data the ranges reach', () => {
    // Rows 41 to 60 are blank, and a separate block starts after a gap below the data.
    const s = sheet('Ops', { ...orders(39), A62: 'Notes', A63: 'Checked', E1: 'Total', E2: '=SUM($C$2:$C$60)' });
    expect(planFor(s, 'fixed-long-ranges').steps[0]).toContain('`A1:C40`');
  });

  it('has no plan when the data is on another sheet or runs past the range', () => {
    const summary = sheet('Summary', { A1: 'Region', A2: 'East', B1: 'Spend', B2: '=SUMIFS(Data!$C$2:$C$200,Data!$A$2:$A$200,A2)' });
    expect(planFix(findingFor(summary, 'fixed-long-ranges'), summary)).toBeNull();
    const short = sheet('Ops', { ...orders(45), E1: 'Total', E2: '=SUM($C$2:$C$40)' });
    expect(planFix(findingFor(short, 'fixed-long-ranges'), short)).toBeNull();
    // A blank row inside the data doesn't hide rows the range leaves out.
    const gap: Cells = { ...orders(45), E1: 'Total', E2: '=SUM($C$2:$C$40)' };
    delete gap.A20;
    delete gap.B20;
    delete gap.C20;
    expect(planFix(findingFor(sheet('Ops', gap), 'fixed-long-ranges'), sheet('Ops', gap))).toBeNull();
    // Nor do rows past a shorter range.
    const mixed = sheet('Ops', { ...orders(45), E1: 'Total', E2: '=SUMIFS($C$2:$C$46,$B$2:$B$30,">20")' });
    expect(planFix(findingFor(mixed, 'fixed-long-ranges'), mixed)).toBeNull();
  });
});

describe('planFix: findings without a guided fix', () => {
  it('lists the fixable detectors', () => {
    expect([...FIXABLE].sort()).toEqual([
      'fixed-long-ranges',
      'hardcoded-constants',
      'iferror-lookup',
      'inconsistent-column',
      'nested-if',
      'sumproduct-conditions',
      'typed-over-formula',
      'vlookup-column-number',
      'volatile-references',
    ]);
  });

  it('returns null for approximate lookups, joined keys and other sheets', () => {
    const s = sheet('Ops', { ...orders(), E1: 'Key', ...fill('E', 2, 21, (r) => `=A${r}&"|"&B${r}`), F1: 'Band', ...fill('F', 2, 21, (r) => `=VLOOKUP(B${r},$H$2:$J$40,2)`), ...items() });
    expect(planFix(findingFor(s, 'concatenated-keys'), s)).toBeNull();
    expect(planFix(findingFor(s, 'vlookup-approximate'), s)).toBeNull();
    const elsewhere = { ...findingFor(lookupSheet(), 'vlookup-column-number'), cells: [{ sheet: 'Other', address: 'E2' }] };
    expect(planFix(elsewhere, lookupSheet())).toBeNull();
  });
});

// ---------- The block to fix ----------

describe('planFix: the block it covers', () => {
  const lookup = (r: number) => `=VLOOKUP(A${r},$H$2:$J$40,3,FALSE)`;

  it('starts below the header and stops at a blank row', () => {
    const s = sheet('Orders', { ...orders(), E1: 'Lead days', ...fill('E', 2, 9, lookup), ...fill('E', 11, 21, lookup), ...items() });
    expect(planFix(finding('vlookup-column-number', 'Orders', 'E4'), s)?.range).toBe('E2:E9');
    expect(planFix(finding('vlookup-column-number', 'Orders', 'E15'), s)?.range).toBe('E11:E21');
  });

  it('stops at a typed value, except when fixing typed values', () => {
    const s = sheet('Orders', { ...orders(), E1: 'Lead days', ...fill('E', 2, 21, lookup), E9: 7, ...items() });
    expect(planFix(finding('vlookup-column-number', 'Orders', 'E4'), s)?.range).toBe('E2:E8');
    expect(planFix(finding('vlookup-column-number', 'Orders', 'E15'), s)?.range).toBe('E10:E21');
    expect(planFix(finding('typed-over-formula', 'Orders', 'E9'), s)?.range).toBe('E2:E21');
  });

  it('leaves out a total under the column', () => {
    const s = sheet('Orders', { ...orders(), E1: 'Lead days', ...fill('E', 2, 21, lookup), E22: '=SUM(E2:E21)', ...items() });
    expect(planFix(finding('vlookup-column-number', 'Orders', 'E2'), s)?.range).toBe('E2:E21');
  });

  it('keeps a column that reads whole Table columns, with or without a typed value, and leaves out its total', () => {
    const cost = (r: number) => `=IFNA(XLOOKUP(A${r},Items[SKU],Items[Cost]),0)`;
    const s = sheet('Orders', { ...orders(), E1: 'Cost', ...fill('E', 2, 21, cost), E22: '=SUBTOTAL(9,E2:E21)' });
    expect(planFix(finding('iferror-lookup', 'Orders', 'E5'), s)?.range).toBe('E2:E21');
    const typed = sheet('Orders', { ...orders(), E1: 'Cost', ...fill('E', 2, 21, cost), E9: 6.5 });
    expect(planFix(finding('typed-over-formula', 'Orders', 'E9'), typed)?.range).toBe('E2:E21');
  });

  it('works when the used range doesn’t start at A1', () => {
    const s = sheet('Report', { C5: 'SKU', D5: 'Lead days', ...fill('C', 6, 12, (r) => `SKU-${r}`), ...fill('D', 6, 12, (r) => `=VLOOKUP(C${r},$H$6:$J$40,3,FALSE)`), H6: 'x' });
    expect(s.address).toBe('C5:H12');
    expect(planFor(s, 'vlookup-column-number').range).toBe('D6:D12');
  });

  it(`caps a fix at ${MAX_FIX_ROWS.toLocaleString('en-US')} rows, keeping the finding’s cell inside`, () => {
    const rows = 6000;
    const s = sheet('Big', { A1: 'SKU', B1: 'Lead days', ...fill('A', 2, rows + 1, (r) => `SKU-${r}`), ...fill('B', 2, rows + 1, (r) => `=VLOOKUP(A${r},$H$2:$J$40,3,FALSE)`) });
    const top = planFix(finding('vlookup-column-number', 'Big', 'B2'), s)!;
    expect(top.range).toBe(`B2:B${MAX_FIX_ROWS + 1}`);
    expect(top.continues).toBe('The column continues below `B5001`. Fill the new formula to the end of it.');
    expect(top.steps[2]).toBe('Fill it down: copy `B2`, select `B3:B5001`, and paste. The column continues below `B5001`. Fill the new formula to the end of it.');
    const bottom = planFix(finding('vlookup-column-number', 'Big', 'B5500'), s)!;
    expect(bottom.range).toBe(`B5500:B${rows + 1}`);
    expect(bottom.continues).toBe('The column continues above `B5500`. Fill the new formula up to the top of it.');
    expectHouseStyle(top);
  });

  it('says the column goes on when the sheet was scanned in part', () => {
    const cells: Cells = { ...orders(), E1: 'Lead days', ...fill('E', 2, 21, (r) => `=VLOOKUP(A${r},$H$2:$J$21,3,FALSE)`), H1: 'SKU', I1: 'Vendor', J1: 'Lead days' };
    const s = sheet('Orders', cells);
    const whole = planFor(s, 'vlookup-column-number');
    expect(whole.continues).toBeUndefined();
    expect(scanEdgeBelow(whole, s)).toBe('E22');
    const part = planFix(findingFor(s, 'vlookup-column-number'), s, { partial: true })!;
    expect(part.continues).toBe('The column continues below `E21`. Fill the new formula to the end of it.');
    expect(part.steps[2]).toContain('continues below `E21`');
    // A block that stops above the last scanned row can’t go on past it.
    expect(scanEdgeBelow(planFor(lookupSheet(), 'vlookup-column-number'), lookupSheet())).toBeNull();
  });

  it('adds nothing about the rest of the column to a typed-over fix', () => {
    const s = sheet('Ops', { ...orders(), D1: 'Extended', ...fill('D', 2, 21, (r) => `=B${r}*C${r}`), D9: 412.5 });
    expect(planFix(findingFor(s, 'typed-over-formula'), s, { partial: true })?.continues).toBeUndefined();
  });
});

// ---------- Grading ----------

describe('gradeFix', () => {
  const s = lookupSheet();
  const p = planFor(s, 'vlookup-column-number');
  const original: Cells = fill('E', 2, 21, (r) => `=VLOOKUP(A${r},$H$2:$J$40,3,FALSE)`);
  const fixed: Cells = fill('E', 2, 21, (r) => `=XLOOKUP(A${r},$H$2:$H$40,$J$2:$J$40)`);
  // Row r's SKU sits in row r of the item master, so its lead days are J's value for that row.
  const leadDays = (address: string) => 3 + (parseCell(address).row % 11);
  const before = readRange(p.range, original, leadDays);

  it('passes when the copy matches and the rules hold', () => {
    const items = gradeFix(p, before, readRange(p.range, fixed, leadDays));
    expect(items.map((i) => i.id)).toEqual(['fix-values', 'fix-forbid-0', 'fix-require-0']);
    expect(passed(items)).toBe(true);
  });

  it('names the first changed cell and counts the rest', () => {
    const items = gradeFix(p, before, readRange(p.range, fixed, (a) => (a === 'E7' || a === 'E12' ? 99 : leadDays(a))));
    expect(byId(items, 'fix-values')).toMatchObject({ status: 'fail', detail: '2 cells changed. E7 was 10, now 99.', focus: 'E7' });
    const one = gradeFix(p, before, readRange(p.range, fixed, (a) => (a === 'E7' ? '' : leadDays(a))));
    expect(byId(one, 'fix-values')?.detail).toBe('E7 was 10, now blank.');
  });

  it('explains an error that appears, and shows enough digits to tell values apart', () => {
    const error = gradeFix(p, before, readRange(p.range, fixed, (a) => (a === 'E5' ? '#N/A' : leadDays(a))));
    expect(byId(error, 'fix-values')?.detail).toBe('E5 was 8, now #N/A. A lookup didn’t find a match.');
    const close = gradeFix(p, before, readRange(p.range, fixed, (a) => (a === 'E5' ? 8.00001 : leadDays(a))));
    expect(byId(close, 'fix-values')?.detail).toBe('E5 was 8, now 8.00001.');
    const float = gradeFix(p, before, readRange(p.range, fixed, (a) => (a === 'E5' ? 8 + 1e-12 : leadDays(a))));
    expect(byId(float, 'fix-values')?.status).toBe('pass');
  });

  it('fails forbid and require rules at the first offending cell', () => {
    const half = { ...fixed, E7: original.E7, E9: original.E9 };
    const items = gradeFix(p, before, readRange(p.range, half, leadDays));
    expect(byId(items, 'fix-values')?.status).toBe('pass');
    expect(byId(items, 'fix-forbid-0')).toMatchObject({ status: 'fail', focus: 'E7', detail: 'Found in E7 and E9. Replace each one with XLOOKUP, pointing at the key column and the return column.' });
    expect(byId(items, 'fix-require-0')).toMatchObject({ status: 'fail', focus: 'E7' });
    expect(byId(items, 'fix-require-0')?.detail).toMatch(/^Not yet in E7 and E9\. For example, E2 becomes =XLOOKUP\(A2,\$H\$2:\$H\$40,\$J\$2:\$J\$40\)\.$/);
  });

  it('accepts INDEX and MATCH as well as XLOOKUP', () => {
    const indexMatch = fill('E', 2, 21, (r) => `=INDEX($J$2:$J$40,MATCH(A${r},$H$2:$H$40,0))`);
    expect(passed(gradeFix(p, before, readRange(p.range, indexMatch, leadDays)))).toBe(true);
  });

  it('fails clearly on a size mismatch', () => {
    const after = readRange('E2:E20', fixed, leadDays);
    const items = gradeFix(p, before, after);
    expect(items).toHaveLength(1);
    expect(items[0]).toMatchObject({ id: 'fix-values', status: 'fail' });
    expect(items[0].detail).toContain('E2:E21');
  });

  describe('typed-over fixes', () => {
    const extended = (r: number) => `=B${r}*C${r}`;
    const cells: Cells = { ...orders(), D1: 'Extended', ...fill('D', 2, 21, extended), D9: 412.5 };
    const typed = sheet('Ops', cells);
    const plan = planFor(typed, 'typed-over-formula');
    const qty = (r: number) => (r * 7) % 40;
    const price = (r: number) => 3 + (r % 9) * 1.25;
    const live = (a: string) => qty(parseCell(a).row) * price(parseCell(a).row);
    const original = readRange(plan.range, cells, (a) => (a === 'D9' ? 412.5 : live(a)));
    const restored = { ...cells, D9: extended(9) };

    it('lets the typed cell change and checks it holds the column’s formula', () => {
      const items = gradeFix(plan, original, readRange(plan.range, restored, live));
      expect(items.map((i) => i.id)).toEqual(['fix-values', 'fix-formulas', 'fix-restore']);
      expect(passed(items)).toBe(true);
    });

    it('fails while the typed value is still there', () => {
      const items = gradeFix(plan, original, original);
      expect(byId(items, 'fix-values')?.status).toBe('pass');
      expect(byId(items, 'fix-formulas')).toMatchObject({ status: 'fail', detail: 'D9 has a typed value.', focus: 'D9', label: 'Every cell in D2:D21 has a formula' });
      expect(byId(items, 'fix-restore')).toMatchObject({ status: 'fail', focus: 'D9', label: 'Typed values replaced with the column’s formula' });
      expect(byId(items, 'fix-restore')?.detail).toBe('D9 doesn’t hold the column’s formula yet. Copy D8 and paste it over each typed value.');
    });

    it('rejects a different formula in the typed cell, and an emptied cell', () => {
      const wrong = gradeFix(plan, original, readRange(plan.range, { ...cells, D9: '=412.5' }, (a) => (a === 'D9' ? 412.5 : live(a))));
      expect(byId(wrong, 'fix-formulas')?.status).toBe('pass');
      expect(byId(wrong, 'fix-restore')?.status).toBe('fail');
      const emptied = gradeFix(plan, original, readRange(plan.range, { ...cells, D9: '', D10: '' }, (a) => (a === 'D9' || a === 'D10' ? '' : live(a))));
      expect(byId(emptied, 'fix-formulas')?.detail).toBe('2 cells are empty or have typed values, starting at D9.');
      expect(byId(emptied, 'fix-values')).toMatchObject({ status: 'fail', focus: 'D10' });
    });
  });

  describe('running columns', () => {
    const net = (r: number) => (r % 5) * 10 - 15;
    /** A running balance in C with a deliberate fees section in C17:C19 (fees in D). */
    const base = (): Cells => ({
      A1: 'Day',
      B1: 'Net',
      ...fill('B', 2, 21, net),
      C1: 'Balance',
      C2: '=B2',
      ...fill('C', 3, 21, (r) => `=C${r - 1}+B${r}`),
      ...fill('C', 17, 19, (r) => `=C${r - 1}+B${r}-D${r}`),
      D1: 'Fees',
      ...fill('D', 17, 19, (r) => 4 + (r % 3)),
    });
    /** What Excel shows in C for these formulas. */
    const balances = (cells: Cells) => {
      const v: Record<string, number> = {};
      for (let r = 2; r <= 21; r++) {
        const f = cells[`C${r}`];
        if (typeof f === 'number') {
          v[`C${r}`] = f;
          continue;
        }
        const m = /^=(?:C(\d+)\+)?B(\d+)(?:-D(\d+))?$/.exec(String(f))!;
        v[`C${r}`] = (m[1] ? v[`C${m[1]}`] : 0) + net(Number(m[2])) - (m[3] ? Number(cells[`D${m[3]}`]) : 0);
      }
      return (a: string): Cell => v[a] ?? '';
    };
    const grade = (plan: FixPlan, original: Cells, copy: Cells) => gradeFix(plan, readRange(plan.range, original, balances(original)), readRange(plan.range, copy, balances(copy)));
    const flatten = (cells: Cells, from: number, to: number): Cells => ({ ...cells, ...fill('C', from, to, (r) => `=C${r - 1}+B${r}`) });

    it('restores a typed cell above a deliberate section and leaves the section alone', () => {
      const cells = { ...base(), C8: 500 };
      const plan = planFor(sheet('Cash', cells), 'typed-over-formula');
      expect(plan.restore).toMatchObject({ cells: ['C8'], model: 'C7' });
      expect(plan.valuesMayChange).toContain('C17');
      const restored = { ...cells, C8: '=C7+B8' };
      const items = grade(plan, cells, restored);
      expect(items.map((i) => i.id)).toEqual(['fix-values', 'fix-formulas', 'fix-restore', 'fix-builders']);
      expect(passed(items)).toBe(true);
    });

    it('fails when the section below is overwritten with the column’s formula', () => {
      const cells = { ...base(), C8: 500 };
      const plan = planFor(sheet('Cash', cells), 'typed-over-formula');
      const items = grade(plan, cells, flatten({ ...cells, C8: '=C7+B8' }, 17, 19));
      expect(byId(items, 'fix-values')?.status).toBe('pass');
      expect(byId(items, 'fix-restore')?.status).toBe('pass');
      expectPlainDetails(items);
      expect(byId(items, 'fix-builders')).toMatchObject({
        status: 'fail',
        focus: 'C17',
        label: 'Rows that build on the repaired cells keep their formulas',
        detail: 'C17 and 2 other cells no longer have their formulas from your sheet. Only the typed cells needed the column’s formula, so copy the others back from your sheet.',
      });
    });

    it('keeps a deliberate last row', () => {
      const cells: Cells = { ...base(), ...fill('C', 17, 19, (r) => `=C${r - 1}+B${r}`), D17: '', D18: '', D19: '', C21: '=C20+B21-D21', D21: 7, C8: 500 };
      const plan = planFor(sheet('Cash', cells), 'typed-over-formula');
      expect(plan.valuesMayChange).toContain('C21');
      expect(passed(grade(plan, cells, { ...cells, C8: '=C7+B8' }))).toBe(true);
      const items = grade(plan, cells, flatten({ ...cells, C8: '=C7+B8' }, 21, 21));
      expect(byId(items, 'fix-builders')?.detail).toBe('C21 no longer has its formula from your sheet, =C20+B21-D21. Only the typed cells needed the column’s formula, so put it back.');
    });

    it('repairs an odd cell above a deliberate section', () => {
      const cells = { ...base(), C8: '=C6+B8' };
      const plan = planFor(sheet('Cash', cells), 'inconsistent-column');
      expect(plan.restore).toMatchObject({ cells: ['C8'], model: 'C7' });
      expect(passed(grade(plan, cells, { ...cells, C8: '=C7+B8' }))).toBe(true);
      const items = grade(plan, cells, flatten({ ...cells, C8: '=C7+B8' }, 17, 19));
      expect(byId(items, 'fix-restore')?.status).toBe('pass');
      expect(byId(items, 'fix-builders')).toMatchObject({ status: 'fail', focus: 'C17' });
      expect(byId(items, 'fix-builders')?.detail).toContain('Only the cells that differed needed the column’s formula');
    });
  });

  describe('formulas that point at the copy', () => {
    const copy = { name: 'Fix-Orders', original: s };

    it('fails a reference to the copy’s own sheet, and tells how to remove it', () => {
      const selfRef = fill('E', 2, 21, (r) => `=XLOOKUP(A${r},'Fix-Orders'!$H$2:$H$40,'Fix-Orders'!$J$2:$J$40)`);
      const items = gradeFix(p, before, readRange(p.range, selfRef, leadDays), copy);
      expect(byId(items, 'fix-values')?.status).toBe('pass');
      expect(byId(items, 'fix-copy-refs')).toMatchObject({
        status: 'fail',
        focus: 'E2',
        label: 'Formulas don’t depend on the copy',
        detail: "E2 and 19 other cells name the copy’s own sheet, so they’d show #REF! once you delete the copy. Select E2:E21 on the copy, choose Home › Find & Select › Replace, and replace 'Fix-Orders'! with nothing.",
      });
      expectPlainDetails(items);
    });

    it('adds nothing when the formulas name only the learner’s own sheets and Tables', () => {
      const own = fill('E', 2, 21, (r) => `=XLOOKUP(A${r},Orders!$H$2:$H$40,Orders!$J$2:$J$40)`);
      expect(gradeFix(p, before, readRange(p.range, own, leadDays), copy).map((i) => i.id)).toEqual(['fix-values', 'fix-forbid-0', 'fix-require-0']);
    });

    describe('a Table Excel renamed on the copy', () => {
      const cost = (r: number) => `=IFNA(XLOOKUP(A${r},Items[SKU],Items[Cost]),0)`;
      const cells: Cells = { ...orders(), E1: 'Cost', ...fill('E', 2, 21, cost), E9: 6.5 };
      const typed = sheet('Orders', cells);
      const plan = planFix(finding('typed-over-formula', 'Orders', 'E9'), typed)!;
      const price = (a: string) => 2 + (parseCell(a).row % 4);
      const original = readRange(plan.range, cells, (a) => (a === 'E9' ? 6.5 : price(a)));
      // As Excel copies the sheet: its Items Table becomes Items2, and so do the formulas on the copy.
      const onCopy: Cells = Object.fromEntries(Object.entries(cells).map(([a, v]) => [a, typeof v === 'string' ? v.replace(/Items\[/g, 'Items2[') : v]));
      const restored = readRange(plan.range, { ...onCopy, E9: cost(9).replace(/Items\[/g, 'Items2[') }, price);

      it('takes the column’s formula from the copy itself', () => {
        expect(passed(gradeFix(plan, original, restored))).toBe(true);
      });

      it('asks for the original Table’s name before the formulas go back', () => {
        const items = gradeFix(plan, original, restored, { name: 'Fix-Orders', original: typed });
        expect(byId(items, 'fix-restore')?.status).toBe('pass');
        expect(byId(items, 'fix-copy-refs')).toMatchObject({
          status: 'fail',
          focus: 'E2',
          detail: 'E2 and 19 other cells refer to Items2, the copy’s own Table, so they’d show #REF! once you delete the copy. Change Items2 to Items, your sheet’s Table. Select E2:E21 on the copy, choose Home › Find & Select › Replace, and replace Items2[ with Items[.',
        });
        expectPlainDetails(items);
        const renamedBack = gradeFix(plan, original, readRange(plan.range, { ...cells, E9: cost(9) }, price), { name: 'Fix-Orders', original: typed });
        expect(renamedBack.map((i) => i.id)).toEqual(['fix-values', 'fix-formulas', 'fix-restore']);
        expect(passed(renamedBack)).toBe(true);
      });

      it('compares the rows below with the original through the rename', () => {
        const fee = (table: string) => (r: number) => `=C${r - 1}+B${r}-INDEX(${table}[Fee],1)`;
        const cash: Cells = { A1: 'Day', B1: 'Net', ...fill('B', 2, 21, (r) => r), C1: 'Balance', C2: '=B2', ...fill('C', 3, 21, fee('Rates')), C8: 500 };
        const running = planFor(sheet('Cash', cash), 'typed-over-formula');
        // C8 restored on the copy, where every formula names Rates2.
        const fixed = { ...cash, ...fill('C', 3, 21, fee('Rates2')) };
        const items = gradeFix(running, readRange(running.range, cash, () => 1), readRange(running.range, fixed, () => 1), { name: 'Fix-Cash', original: sheet('Cash', cash) });
        expect(byId(items, 'fix-restore')?.status).toBe('pass');
        expect(byId(items, 'fix-builders')?.status).toBe('pass');
        expect(byId(items, 'fix-copy-refs')?.detail).toMatch(/^C3 and 18 other cells refer to Rates2, the copy’s own Table/);
      });
    });

    it('lists Tables it can’t place, except the one a Table fix makes', () => {
      const other = fill('E', 2, 21, (r) => `=XLOOKUP(A${r},Lookup[SKU],Lookup[Lead days])`);
      expect(copyReferences(p, readRange(p.range, other, leadDays), copy)).toEqual({ sheet: undefined, renamed: [], unknown: ['Lookup'] });
      const table = planFor(sheet('Ops', { ...orders(39), E1: 'Total', E2: '=SUM($C$2:$C$40)' }), 'fixed-long-ranges');
      expect(copyReferences(table, readRange(table.range, { E2: '=SUM(OpsData[Unit price])' }, () => 1), { name: 'Fix-Ops' }).unknown).toEqual([]);
    });
  });

  describe('IFERROR fixes', () => {
    const plan = planFor(s, 'iferror-lookup');
    const vendor = (a: string) => ['Apex Supply', 'Birchwood Co', 'Cobalt Parts'][parseCell(a).row % 3];
    const original = readRange(plan.range, {}, (a) => (a === 'F7' || a === 'F12' ? 'Missing' : vendor(a)));
    const xlookup = fill('F', 2, 21, (r) => `=XLOOKUP(A${r},$H$2:$H$40,$I$2:$I$40,"Missing")`);
    const showing = (shown: Record<string, Cell>) => readRange(plan.range, xlookup, (a) => (a in shown ? shown[a] : a === 'F7' || a === 'F12' ? 'Missing' : vendor(a)));

    it('keeps the fallback when it’s a plain value', () => {
      expect(plan.fallback).toBe('Missing');
      const blank = planFor(sheet('Orders', { ...orders(5), E1: 'Vendor', ...fill('E', 2, 6, (r) => `=IFERROR(INDEX($I$2:$I$40,MATCH(A${r},$H$2:$H$40,0)),"")`) }), 'iferror-lookup');
      expect(blank.fallback).toBe('');
      const ifna = planFor(sheet('Orders', { ...orders(5), E1: 'Cost', ...fill('E', 2, 6, (r) => `=IFNA(XLOOKUP(A${r},Items[SKU],Items[Cost]),0)`) }), 'iferror-lookup');
      expect(ifna.fallback).toBeUndefined();
    });

    it('explains an error IFERROR was hiding', () => {
      expect(byId(gradeFix(plan, original, showing({ F7: '#VALUE!' })), 'fix-values')).toMatchObject({
        status: 'fail',
        focus: 'F7',
        detail: 'F7 now shows #VALUE!, which IFERROR was hiding. A value has the wrong type, often text where a number is expected. Fix that on your original sheet, then start over with a fresh copy.',
      });
      expectPlainDetails(gradeFix(plan, original, showing({ F7: '#VALUE!' })));
      expect(byId(gradeFix(plan, original, showing({ F7: '#DIV/0!', F12: '#REF!' })), 'fix-values')?.detail).toBe(
        'F7 and 1 other cell now show errors IFERROR was hiding, starting with #DIV/0! in F7. Fix them on your original sheet, then start over with a fresh copy.',
      );
    });

    it('treats a missing fallback and a broken formula as ordinary changes', () => {
      expect(byId(gradeFix(plan, original, showing({ F7: '#N/A' })), 'fix-values')?.detail).toBe('F7 was “Missing”, now #N/A. A lookup didn’t find a match.');
      const broken = gradeFix(plan, original, readRange(plan.range, xlookup, () => '#VALUE!'));
      expect(byId(broken, 'fix-values')?.detail).toMatch(/^20 cells changed\. F2 was “Cobalt Parts”, now #VALUE!\./);
    });
  });

  it('grades a nested-IF fix on its depth, ignoring text', () => {
    const cells: Cells = { ...orders(), D1: 'Tier', ...fill('D', 2, 21, (r) => `=IF(B${r}>30,"A",IF(B${r}>20,"B",IF(B${r}>10,"C","D")))`) };
    const plan = planFor(sheet('Sales', cells), 'nested-if');
    const tier = (a: string) => {
      const q = (parseCell(a).row * 7) % 40;
      return q > 30 ? 'A' : q > 20 ? 'B' : q > 10 ? 'C' : 'D';
    };
    const before = readRange(plan.range, cells, tier);
    // LEFT(…,0) adds nothing to the result; the IFs inside its text mustn't count as nesting.
    const ifs = fill('D', 2, 21, (r) => `=IFS(B${r}>30,"A",B${r}>20,"B",B${r}>10,"C",TRUE,"D")&LEFT("IF(IF(IF(",0)`);
    expect(passed(gradeFix(plan, before, readRange(plan.range, ifs, tier)))).toBe(true);
    const items = gradeFix(plan, before, before);
    expect(byId(items, 'fix-forbid-0')).toMatchObject({ status: 'fail', focus: 'D2', detail: expect.stringMatching(/^Found in D2 and 19 other cells\./) });
  });

  it('runs the require rule only when the whole block had the pattern', () => {
    const mixed = sheet('Orders', { ...orders(), E1: 'Lead days', E2: '=B2*2', ...fill('E', 3, 21, (r) => `=VLOOKUP(A${r},$H$2:$J$40,3,FALSE)`), ...items() });
    const plan = planFix(finding('vlookup-column-number', 'Orders', 'E3'), mixed)!;
    expect(plan.range).toBe('E2:E21');
    expect(plan.require).toEqual([]);
    expect(plan.steps[0]).toContain('`E3`');
  });

  it('reports a require rule with no formulas left as a failure', () => {
    const cleared = readRange(p.range, {}, leadDays);
    expect(byId(gradeFix(p, before, cleared), 'fix-require-0')).toMatchObject({ status: 'fail', detail: expect.stringMatching(/^E2:E21 has no formulas yet\./) });
  });
});

// ---------- Speed ----------

describe('fix speed', () => {
  it(`plans and grades a ${MAX_FIX_ROWS.toLocaleString('en-US')}-row column well under a second`, () => {
    const rows = MAX_FIX_ROWS;
    const cells: Cells = { A1: 'SKU', B1: 'Cost', ...fill('A', 2, rows + 1, (r) => `SKU-${r}`), ...fill('B', 2, rows + 1, (r) => `=IFERROR(VLOOKUP(A${r},$H$2:$J$400,3,FALSE),0)`), H1: 'x' };
    const s = sheet('Big', cells);
    const started = performance.now();
    const plan = planFix(finding('iferror-lookup', 'Big', 'B2'), s)!;
    const fixed = fill('B', 2, rows + 1, (r) => `=XLOOKUP(A${r},$H$2:$H$400,$J$2:$J$400,0)`);
    const items = gradeFix(plan, readRange(plan.range, cells, () => 4), readRange(plan.range, fixed, () => 4), { name: 'Fix-Big', original: s });
    const elapsed = performance.now() - started;
    expect(plan.range).toBe(`B2:B${rows + 1}`);
    expect(passed(items)).toBe(true);
    expect(elapsed).toBeLessThan(1000);
  });
});

import { describe, expect, it, vi } from 'vitest';
import { colToNumber, numberToCol, parseCell } from '../src/engine/address';
import { scanWorkbook, type Finding, type SheetFormulas } from '../src/engine/scan';
import { analyzeFormula, arithmeticConstants, isRepeatedConstant, maskFormula, runRefFacts, splitArgs } from '../src/engine/scanDetectors';
import type { Cell, Grid } from '../src/engine/types';
import { MockHost } from '../src/excel/mockHost';

// ---------- Test helpers ----------

type Cells = Record<string, Cell>;

/** Converts the A1 references in a formula to R1C1 as seen from (row, col), the way Excel reports .formulasR1C1. */
function toR1C1(formula: string, row: number, col: number): string {
  const rowPart = (abs: string, r: number) => (abs ? `R${r}` : r === row ? 'R' : `R[${r - row}]`);
  const colPart = (abs: string, c: number) => (abs ? `C${c}` : c === col ? 'C' : `C[${c - col}]`);
  // One pass, so a converted whole-column reference (C4:C4) is never re-read as a cell.
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

const scan = (...sheets: SheetFormulas[]) => scanWorkbook(sheets);
const finding = (findings: Finding[], id: string) => findings.find((f) => f.id === id);
const one = (formula: string, id: string) => finding(scan(sheet('Ops', { A1: 'x', E2: formula })), id);
const addresses = (f: Finding | undefined) => f?.cells.map((c) => c.address) ?? [];

// ---------- Formula helpers ----------

describe('formula text helpers', () => {
  it('splits arguments only at top-level commas', () => {
    expect(splitArgs('A2,"Smith, J",{1,2;3,4},MAX(B1,C1),Sales[[#This Row],[Qty]]')).toEqual([
      'A2',
      '"Smith, J"',
      '{1,2;3,4}',
      'MAX(B1,C1)',
      'Sales[[#This Row],[Qty]]',
    ]);
    expect(splitArgs('A2,F:G,2,')).toEqual(['A2', 'F:G', '2', '']);
    expect(splitArgs("'Q1, East'!A2:A9,(B1,C1)")).toEqual(["'Q1, East'!A2:A9", '(B1,C1)']);
  });

  it('masks text, sheet names and structured references without moving anything', () => {
    const body = `"a,b"&'My Sheet'!A1&Sales[[#This Row],[Qty]]`;
    const masked = maskFormula(body);
    expect(masked).toHaveLength(body.length);
    expect(masked).not.toContain('a,b');
    expect(masked).not.toContain('My Sheet');
    expect(masked).not.toContain('Qty');
    expect(masked).toContain('!A1&Sales[');
  });

  it('finds numbers used in arithmetic but not structural arguments', () => {
    expect(arithmeticConstants('IF(C2>5000,C2*0.05,0)')).toEqual([5000, 0.05]);
    expect(arithmeticConstants('ROUND(B2*1.0825,4)')).toEqual([1.0825]);
    expect(arithmeticConstants('LEFT(A2,3)&DATE(2026,1,15)')).toEqual([]);
    expect(arithmeticConstants('EDATE(A2,-3)')).toEqual([]);
    expect(arithmeticConstants('A2-30')).toEqual([30]);
    expect(arithmeticConstants('B2*-0.5')).toEqual([0.5]);
    expect(arithmeticConstants('B2*8%')).toEqual([0.08]);
    expect(arithmeticConstants('SUM(2:5)*3')).toEqual([3]);
    expect(arithmeticConstants('B2/12+C2*100')).toEqual([]);
    expect(arithmeticConstants(maskFormula('COUNTIF(A:A,">5000")'))).toEqual([]);
  });

  it('measures IF nesting through other functions', () => {
    expect(analyzeFormula('=IF(A2>1,1,IF(A2>2,2,IF(A2>3,3,4)))').ifDepth).toBe(3);
    expect(analyzeFormula('=IF(A1,1,0)+IF(B1,1,0)+IF(C1,1,0)').ifDepth).toBe(1);
    expect(analyzeFormula('=IF(A2="",0,ROUND(IF(B2>0,IF(C2>0,1,2),3),0))').ifDepth).toBe(3);
    expect(analyzeFormula('=IF(A2="(",1,IF(A2=")",2,0))').ifDepth).toBe(2);
  });

  it('reads R1C1 offsets and Table references for column checks', () => {
    // Column D is column 4. Offsets in brackets must survive masking.
    expect(runRefFacts('=RC[-2]*R[-1]C[-1]', 4)).toMatchObject({ rowAbove: true, ownCell: false, shape: 'X*X' });
    expect(runRefFacts('=SUM(R[-20]C:R[-1]C)', 4)).toMatchObject({ ownRange: true });
    expect(runRefFacts('=SUM(R2C4:R21C4)', 4)).toMatchObject({ ownRange: true });
    expect(runRefFacts('=SUMIFS(Data!R2C4:R500C4,Data!R2C1:R500C1,RC1)', 4)).toMatchObject({ ownRange: false, summary: false });
    expect(runRefFacts('=R[-3]C-R[-1]C', 4)).toMatchObject({ ownCell: true });
    expect(runRefFacts('=SUBTOTAL(109,[Extended])', 4)).toMatchObject({ summary: true });
    expect(runRefFacts('=[@Qty]*[@[Unit price]]', 4)).toMatchObject({ summary: false, hasRef: true, shape: 'X*X' });
    expect(runRefFacts('=""', 4)).toMatchObject({ hasRef: false });
    expect(runRefFacts('=TaxRate*RC[-1]', 4)).toMatchObject({ hasRef: true });
  });

  it('measures the longest explicit range', () => {
    expect(analyzeFormula('=SUMIFS(Data!$C$2:$C$500,Data!$A$2:$A$500,A2)').longestRange).toBe(499);
    expect(analyzeFormula('=SUM(C:C)+SUM(2:2)').longestRange).toBe(0);
    expect(analyzeFormula('=SUM(\'Q1 2026\'!B2:D41)').longestRange).toBe(40);
  });
});

// ---------- Detectors ----------

describe('vlookup-column-number', () => {
  it('flags VLOOKUP and HLOOKUP with a typed column index', () => {
    const f = finding(scan(sheet('Ops', { ...orders(), ...fill('E', 2, 21, (r) => `=VLOOKUP(A${r},$H$2:$J$40,3,FALSE)`) })), 'vlookup-column-number');
    expect(f).toMatchObject({ severity: 'medium', exerciseId: 'xlookup-lead-time', count: 20 });
    expect(f!.cells).toHaveLength(8);
    expect(f!.cells[0]).toEqual({ sheet: 'Ops', address: 'E2' });
    expect(one('=HLOOKUP("Mar",$B$1:$M$9,4,FALSE)', 'vlookup-column-number')).toBeDefined();
    expect(one('=_xlfn.VLOOKUP(A2,Items,{2,3},FALSE)', 'vlookup-column-number')).toBeDefined();
  });

  it('leaves XLOOKUP and computed column numbers alone', () => {
    expect(one('=XLOOKUP(A2,$H$2:$H$40,$J$2:$J$40)', 'vlookup-column-number')).toBeUndefined();
    expect(one('=VLOOKUP(A2,$H$2:$J$40,MATCH("Cost",$H$1:$J$1,0),FALSE)', 'vlookup-column-number')).toBeUndefined();
    expect(one('=INDEX($J$2:$J$40,MATCH(A2,$H$2:$H$40,0))', 'vlookup-column-number')).toBeUndefined();
  });
});

describe('vlookup-approximate', () => {
  it('flags a missing, TRUE or 1 range_lookup', () => {
    for (const formula of [
      '=VLOOKUP(A2,$H$2:$J$40,2)',
      '=VLOOKUP(A2,$H$2:$J$40,2,TRUE)',
      '=VLOOKUP(A2,$H$2:$J$40,2,1)',
      '=HLOOKUP(A2,$B$1:$M$3,2)',
      '=VLOOKUP(A2,{0,"Low";500,"Mid";2000,"High"},2)',
      '=IFERROR(VLOOKUP(A2,Rates,2,true),0)',
    ]) {
      expect(one(formula, 'vlookup-approximate'), formula).toMatchObject({ severity: 'high', exerciseId: 'xlookup-tiered' });
    }
  });

  it('treats FALSE, 0 and an empty last argument as exact', () => {
    for (const formula of [
      '=VLOOKUP(A2,$H$2:$J$40,2,FALSE)',
      '=VLOOKUP(A2,$H$2:$J$40,2,0)',
      '=VLOOKUP(A2,$H$2:$J$40,2,)',
      '=VLOOKUP("Smith, J",$H$2:$J$40,2,FALSE)',
      '=VLOOKUP(IF(A2="",B2,A2),$H$2:$J$40,MAX(1,2),FALSE)',
      '=XLOOKUP(A2,$H$2:$H$40,$J$2:$J$40,,-1)',
      '=VLOOKUP(A2,$H$2:$J$40,2,$Z$1)',
    ]) {
      expect(one(formula, 'vlookup-approximate'), formula).toBeUndefined();
    }
  });
});

describe('iferror-lookup', () => {
  it('flags IFERROR or IFNA around a lookup', () => {
    for (const formula of [
      '=IFERROR(VLOOKUP(A2,$H$2:$J$40,2,FALSE),0)',
      '=IFNA(XLOOKUP(A2,Items[SKU],Items[Cost]),"")',
      '=IFERROR(INDEX($J$2:$J$40,MATCH(A2,$H$2:$H$40,0)),"Missing")',
      '=ROUND(IFERROR(VLOOKUP(A2,Rates,2,FALSE)*B2,0),2)',
    ]) {
      expect(one(formula, 'iferror-lookup'), formula).toMatchObject({ severity: 'medium', exerciseId: 'xlookup-not-found' });
    }
  });

  it('leaves other IFERROR uses and XLOOKUP fallbacks alone', () => {
    for (const formula of [
      '=IFERROR(B2/C2,0)',
      '=XLOOKUP(A2,Items[SKU],Items[Cost],"Not found")',
      '=IFERROR(B2/C2,VLOOKUP(A2,$H$2:$J$40,2,FALSE))',
      '=VLOOKUP(A2,$H$2:$J$40,2,FALSE)',
    ]) {
      expect(one(formula, 'iferror-lookup'), formula).toBeUndefined();
    }
  });
});

describe('nested-if', () => {
  it('flags three or more levels', () => {
    const f = one('=IF(B2>=10000,0.05,IF(B2>=5000,0.03,IF(B2>=1000,0.01,0)))', 'nested-if');
    expect(f).toMatchObject({ severity: 'medium', exerciseId: 'xlookup-tiered', count: 1 });
  });

  it('leaves two levels, IFS and side-by-side IFs alone', () => {
    expect(one('=IF(B2>=5000,0.03,IF(B2>=1000,0.01,0))', 'nested-if')).toBeUndefined();
    expect(one('=IFS(B2>=10000,0.05,B2>=5000,0.03,B2>=1000,0.01,TRUE,0)', 'nested-if')).toBeUndefined();
    expect(one('=IF(B2>1,1,0)+IF(C2>1,1,0)+IF(D2>1,1,0)', 'nested-if')).toBeUndefined();
  });
});

describe('sumproduct-conditions', () => {
  it('flags SUMPRODUCT that multiplies comparisons SUMIFS could handle', () => {
    for (const formula of [
      '=SUMPRODUCT(($A$2:$A$60="East")*($C$2:$C$60))',
      '=SUMPRODUCT(--($A$2:$A$60=G2),--($B$2:$B$60>=H$1),$C$2:$C$60)',
      '=SUMPRODUCT((Sales[Region]=$G2)*(Sales[Month]=H$1)*Sales[Amount])',
      '=SUMPRODUCT(($B$2:$B$60>=DATE(2026,1,1))*$C$2:$C$60)',
      '=SUMPRODUCT(("East"=$A$2:$A$60)*1)',
    ]) {
      expect(one(formula, 'sumproduct-conditions'), formula).toMatchObject({ severity: 'low', exerciseId: 'sumifs-warehouse' });
    }
  });

  it('leaves weighted sums and logic SUMIFS cannot express alone', () => {
    for (const formula of [
      '=SUMPRODUCT($B$2:$B$60,$C$2:$C$60)',
      '=SUMPRODUCT((MONTH($A$2:$A$60)=3)*$C$2:$C$60)',
      '=SUMPRODUCT((($A$2:$A$60="East")+($A$2:$A$60="West"))*$C$2:$C$60)',
      '=SUMPRODUCT(($A$2:$A$60=$B$2:$B$60)*$C$2:$C$60)',
      '=SUMPRODUCT(($A$2:$A$60="East")*$B$2:$B$60*$C$2:$C$60)',
      '=SUMPRODUCT(--ISNUMBER(SEARCH("pallet",$A$2:$A$60)),$C$2:$C$60)',
      '=SUMIFS($C$2:$C$60,$A$2:$A$60,"East")',
    ]) {
      expect(one(formula, 'sumproduct-conditions'), formula).toBeUndefined();
    }
  });
});

describe('concatenated-keys', () => {
  it('flags a column of joined keys filled down', () => {
    const f = finding(scan(sheet('Ops', { ...orders(10), ...fill('E', 2, 11, (r) => `=A${r}&B${r}`) })), 'concatenated-keys');
    expect(f).toMatchObject({ severity: 'low', exerciseId: 'xlookup-two-keys', count: 10 });
    expect(f!.cells[0]).toEqual({ sheet: 'Ops', address: 'E2' });
  });

  it('recognises separators, Table rows and CONCAT/TEXTJOIN', () => {
    const forms: Array<(r: number) => string> = [
      (r) => `=A${r}&"|"&B${r}`,
      () => '=[@Region]&[@SKU]',
      (r) => `=CONCAT(A${r},"-",B${r})`,
      (r) => `=CONCATENATE($A${r},"_",B${r})`,
      (r) => `=TEXTJOIN("|",TRUE,A${r},B${r},C${r})`,
    ];
    for (const form of forms) {
      const f = finding(scan(sheet('Ops', { ...orders(5), ...fill('E', 2, 6, form) })), 'concatenated-keys');
      expect(f?.count, form(2)).toBe(5);
    }
  });

  it('needs two references and a fill of three or more', () => {
    const negatives: Array<[string, Cells]> = [
      ['two cells', fill('E', 2, 3, (r) => `=A${r}&B${r}`)],
      ['one reference', fill('E', 2, 11, (r) => `="Q"&B${r}`)],
      ['arithmetic', fill('E', 2, 11, (r) => `=B${r}+C${r}`)],
      ['wrapped', fill('E', 2, 11, (r) => `=TRIM(A${r})&TEXT(B${r},"0")`)],
      ['name with a space', fill('E', 2, 11, (r) => `=A${r}&" "&B${r}`)],
      ['TEXTJOIN with a comma and space', fill('E', 2, 11, (r) => `=TEXTJOIN(", ",TRUE,A${r},B${r})`)],
    ];
    for (const [label, cells] of negatives) {
      expect(finding(scan(sheet('Ops', { ...orders(10), ...cells })), 'concatenated-keys'), label).toBeUndefined();
    }
  });
});

describe('fixed-long-ranges', () => {
  it('flags references to 20+ rows on a sheet without Table references', () => {
    const f = finding(scan(sheet('Ops', { ...orders(39), E2: '=SUM(C2:C40)', E3: '=SUM(B2:B21)' })), 'fixed-long-ranges');
    expect(f).toMatchObject({ severity: 'medium', exerciseId: 'tables-convert', count: 2 });
    const summary = sheet('Summary', { A1: 'Region', A2: 'East', B2: '=SUMIFS(Data!$C$2:$C$200,Data!$A$2:$A$200,A2)' });
    expect(finding(scan(summary), 'fixed-long-ranges')?.cells).toEqual([{ sheet: 'Summary', address: 'B2' }]);
  });

  it('leaves short ranges, whole columns and sheets that use Tables alone', () => {
    expect(one('=SUM(C2:C20)', 'fixed-long-ranges')).toBeUndefined();
    expect(one('=SUMIFS(C:C,A:A,F2)', 'fixed-long-ranges')).toBeUndefined();
    expect(one('=SUM(Sales[Amount])', 'fixed-long-ranges')).toBeUndefined();
    const mixed = sheet('Ops', { ...orders(39), E2: '=SUM(C2:C40)', E4: '=SUM(Sales[Amount])' });
    expect(finding(scan(mixed), 'fixed-long-ranges')).toBeUndefined();
    const external = sheet('Ops', { ...orders(39), E2: '=SUM(C2:C40)', E4: '=[1]Budget!B4*2' });
    expect(finding(scan(external), 'fixed-long-ranges')).toBeDefined();
  });
});

describe('inconsistent-column', () => {
  const extended = (r: number) => `=B${r}*C${r}`;

  it('flags the cell that breaks a filled column', () => {
    const cells = { ...orders(), D1: 'Extended', ...fill('D', 2, 21, extended), D12: '=B12*C11' };
    const f = finding(scan(sheet('Ops', cells)), 'inconsistent-column');
    expect(f).toMatchObject({ severity: 'high', exerciseId: 'sumifs-grid', count: 1 });
    expect(addresses(f)).toEqual(['D12']);
  });

  it('flags a broken last cell in a short run', () => {
    const cells = { ...orders(3), ...fill('D', 2, 4, extended), D4: '=B4*C4*2' };
    expect(addresses(finding(scan(sheet('Ops', cells)), 'inconsistent-column'))).toEqual(['D4']);
  });

  it('leaves a properly filled column alone', () => {
    expect(finding(scan(sheet('Ops', { ...orders(), ...fill('D', 2, 21, extended) })), 'inconsistent-column')).toBeUndefined();
  });

  it('allows totals, subtotals and running-balance seeds', () => {
    for (const sum of ['=SUM(D2:D21)', '=SUM($D$2:$D$21)', '=SUM(D$2:D21)', '=AVERAGE(D2:D21)*1']) {
      const total = { ...orders(), ...fill('D', 2, 21, extended), D22: sum };
      expect(finding(scan(sheet('Ops', total)), 'inconsistent-column'), sum).toBeUndefined();
    }
    // A footer that works from other columns (an average price, say) is a deliberate line, not a broken fill.
    const otherColumn = { ...orders(), ...fill('D', 2, 21, extended), D22: '=SUM($C$2:$C$21)/COUNT($C$2:$C$21)' };
    expect(finding(scan(sheet('Ops', otherColumn)), 'inconsistent-column')).toBeUndefined();

    const subtotals = {
      ...orders(),
      ...fill('D', 2, 10, extended),
      D11: '=SUBTOTAL(9,D2:D10)',
      ...fill('D', 12, 20, extended),
      D21: '=SUBTOTAL(9,D12:D20)',
      D22: '=D11+D21',
    };
    expect(finding(scan(sheet('Ops', subtotals)), 'inconsistent-column')).toBeUndefined();

    const balance = { ...orders(), E2: '=B2', ...fill('E', 3, 21, (r) => `=E${r - 1}+B${r}`) };
    expect(finding(scan(sheet('Ops', balance)), 'inconsistent-column')).toBeUndefined();
  });

  it('leaves deliberate sections and columns with no common formula alone', () => {
    const sections = { ...orders(), ...fill('D', 2, 14, extended), ...fill('D', 15, 21, (r) => `=B${r}*C${r}*1.1`) };
    expect(finding(scan(sheet('Ops', sections)), 'inconsistent-column')).toBeUndefined();
    const summary = { A1: 'x', B10: '=SUM(B2:B9)', B11: '=B10*0.1', B12: '=B10+B11' };
    expect(finding(scan(sheet('Ops', summary)), 'inconsistent-column')).toBeUndefined();
  });

  it('flags a hard-coded rate or result in one cell of a fill', () => {
    const rate = (r: number) => `=B${r}*$H$1`;
    const typedRate = { ...orders(), ...fill('D', 2, 21, rate), D7: '=B7*1.08' };
    expect(addresses(finding(scan(sheet('Ops', typedRate)), 'inconsistent-column'))).toEqual(['D7']);
    const typedResult = { ...orders(), ...fill('D', 2, 21, rate), D7: '=412.5' };
    expect(addresses(finding(scan(sheet('Ops', typedResult)), 'inconsistent-column'))).toEqual(['D7']);
    const rounded = { ...orders(), ...fill('D', 2, 21, rate), D7: '=ROUND(B7*$H$1,2)' };
    expect(addresses(finding(scan(sheet('Ops', rounded)), 'inconsistent-column'))).toEqual(['D7']);
  });

  it('still catches a broken fill in the data rows after a subtotal', () => {
    const cells = {
      ...orders(),
      ...fill('D', 2, 10, extended),
      D11: '=SUBTOTAL(9,D2:D10)',
      ...fill('D', 12, 21, extended),
      D16: '=B16*C15',
    };
    expect(addresses(finding(scan(sheet('Ops', cells)), 'inconsistent-column'))).toEqual(['D16']);
  });

  it('needs R1C1 to compare fills', () => {
    const s = sheet('Ops', { ...orders(), ...fill('D', 2, 21, extended), D12: '=B12*C11' });
    expect(finding(scan({ ...s, r1c1: [] }), 'inconsistent-column')).toBeUndefined();
  });
});

describe('typed-over-formula', () => {
  const extended = (r: number) => `=B${r}*C${r}`;

  it('flags a number typed among filled-down formulas', () => {
    const cells = { ...orders(), D1: 'Extended', ...fill('D', 2, 21, extended), D9: 412.5 };
    const findings = scan(sheet('Ops', cells));
    const f = finding(findings, 'typed-over-formula');
    expect(f).toMatchObject({ severity: 'high', count: 1 });
    expect(f!.exerciseId).toBeUndefined();
    expect(addresses(f)).toEqual(['D9']);
    expect(finding(findings, 'inconsistent-column')).toBeUndefined();
  });

  it('flags a number typed over the last formula of the column', () => {
    const cells = { ...orders(), D1: 'Extended', ...fill('D', 2, 21, extended), D21: 88 };
    expect(addresses(finding(scan(sheet('Ops', cells)), 'typed-over-formula'))).toEqual(['D21']);
  });

  it('ignores headers, spilled results and columns that are mostly values', () => {
    const header = { ...orders(), D1: 'Extended', ...fill('D', 2, 21, extended) };
    expect(finding(scan(sheet('Ops', header)), 'typed-over-formula')).toBeUndefined();

    // Excel reports spilled cells as plain values below the anchor formula.
    const spill = { ...orders(), E1: 'Sorted', E2: '=SORT(C2:C21)', ...fill('E', 3, 21, (r) => r * 1.5) };
    expect(finding(scan(sheet('Ops', spill)), 'typed-over-formula')).toBeUndefined();

    const mostlyValues = { ...orders(), ...fill('D', 2, 21, (r) => r * 3), D5: extended(5), D10: extended(10), D15: extended(15) };
    expect(finding(scan(sheet('Ops', mostlyValues)), 'typed-over-formula')).toBeUndefined();

    const seed = { ...orders(), E1: 'Balance', E2: 1500, ...fill('E', 3, 21, (r) => `=E${r - 1}+B${r}`) };
    expect(finding(scan(sheet('Ops', seed)), 'typed-over-formula')).toBeUndefined();
  });
});

describe('volatile-references', () => {
  it('flags INDIRECT and OFFSET', () => {
    expect(one(`=INDIRECT("'"&A2&"'!B5")`, 'volatile-references')).toMatchObject({ severity: 'medium', exerciseId: 'tables-convert' });
    expect(one('=SUM(OFFSET($B$2,0,0,COUNTA($B:$B)-1,1))', 'volatile-references')).toBeDefined();
  });

  it('ignores the words inside text', () => {
    expect(one('="Avoid OFFSET( and INDIRECT( "&A2', 'volatile-references')).toBeUndefined();
    expect(one('=INDEX($B$2:$B$40,ROWS($B$2:B2))', 'volatile-references')).toBeUndefined();
  });
});

describe('hardcoded-constants', () => {
  const rate = (r: number) => `=C${r}*1.0825`;
  const constants = (cells: Cells) => finding(scan(sheet('Ops', { ...orders(), ...cells })), 'hardcoded-constants');

  it('flags a number typed into three or more different formulas', () => {
    const cells = {
      ...orders(9),
      ...fill('E', 2, 10, rate),
      G2: '=SUM(C2:C10)*1.0825',
      H2: '=B2/1.0825',
    };
    const f = finding(scan(sheet('Ops', cells)), 'hardcoded-constants');
    expect(f).toMatchObject({ severity: 'low', count: 11 });
    expect(f!.exerciseId).toBeUndefined();
    expect(f!.why).toBe('1.0825 is typed into 11 cells across three formulas. When it changes, every copy has to be found and edited.');
    expect(addresses(f).slice(0, 3)).toEqual(['E2', 'G2', 'H2']);
  });

  it('counts the same formula on different sheets separately', () => {
    const tab = (name: string) => sheet(name, { ...orders(5), E2: '=SUM(C2:C6)*0.0825' });
    const f = finding(scan(tab('Jan'), tab('Feb'), tab('Mar')), 'hardcoded-constants');
    expect(f?.count).toBe(3);
    expect(f?.why).toBe('0.0825 is typed into three formulas. When it changes, every copy has to be found and edited.');
    const filled = (name: string, rows: number) => sheet(name, { ...orders(5), ...fill('E', 2, 1 + rows, rate) });
    expect(finding(scan(filled('Jan', 3), filled('Feb', 2)), 'hardcoded-constants')?.why).toBe(
      '1.0825 is typed into five cells across two formulas. When it changes, every copy has to be found and edited.',
    );
  });

  it('lists several numbers when more than one repeats', () => {
    const cells = { ...orders(5), E2: '=B2*0.15', E3: '=C3*0.15+250', E4: '=B4/0.15+250', F2: '=C2+250' };
    expect(finding(scan(sheet('Ops', cells)), 'hardcoded-constants')?.why).toBe('0.15 and 250 are each typed into three or more formulas. When one changes, every copy has to be found and edited.');
  });

  it('flags a rate in one formula filled down a column', () => {
    // Regression: real Excel showed =C2*D2*1.0825 filled through 29 rows as one R1C1 formula, and the scan missed it.
    const cells = { D1: 'Pack size', ...fill('D', 2, 30, (r) => 3 + (r % 4)), ...fill('F', 2, 30, (r) => `=C${r}*D${r}*1.0825`) };
    const f = finding(scan(sheet('Ops', { ...orders(29), ...cells })), 'hardcoded-constants');
    expect(f).toMatchObject({ severity: 'low', count: 29 });
    expect(f!.why).toBe('1.0825 is typed into 29 cells. When it changes, every copy has to be found and edited.');
    expect(addresses(f)).toEqual(['F2', 'F3', 'F4', 'F5', 'F6', 'F7', 'F8', 'F9']);
  });

  it('flags the filled rate in the mock workbook and skips the typed-over cell', async () => {
    vi.useFakeTimers();
    try {
      const reading = new MockHost().readWorkbook();
      await vi.runAllTimersAsync();
      const findings = scanWorkbook((await reading).sheets);
      const f = finding(findings, 'hardcoded-constants');
      // Rows 2 to 30 of Total hold the formula, except F17, where 412.5 is typed over it.
      expect(f).toMatchObject({ count: 28 });
      expect(f!.why).toBe('1.0825 is typed into 28 cells. When it changes, every copy has to be found and edited.');
      expect(addresses(f)).not.toContain('F17');
      expect(addresses(finding(findings, 'typed-over-formula'))).toEqual(['F17']);
    } finally {
      vi.useRealTimers();
    }
  });

  it('needs five cells before one filled formula counts', () => {
    expect(constants(fill('E', 2, 5, rate)), '4 cells').toBeUndefined();
    expect(constants(fill('E', 2, 6, rate))?.count, '5 cells').toBe(5);
    // Two different formulas count their cells together.
    expect(constants({ ...fill('E', 2, 3, rate), ...fill('F', 2, 3, (r) => `=B${r}/1.0825`) }), '2 + 2 cells').toBeUndefined();
    const split = constants({ ...fill('E', 2, 4, rate), ...fill('F', 2, 3, (r) => `=B${r}/1.0825`) });
    expect(split?.count, '3 + 2 cells').toBe(5);
    expect(split?.why).toBe('1.0825 is typed into five cells across two formulas. When it changes, every copy has to be found and edited.');
  });

  it('flags rates, thresholds and day offsets in long fills', () => {
    for (const [label, form, n] of [
      ['rate', (r: number) => `=ROUND(B${r}*C${r}*0.0825,2)`, '0.0825'],
      ['markup', (r: number) => `=C${r}*(1+15%)`, '0.15'],
      ['threshold', (r: number) => `=IF(B${r}*C${r}>5000,"Review","")`, '5000'],
      ['due date', (r: number) => `=D${r}+30`, '30'],
    ] as const) {
      expect(constants(fill('E', 2, 21, form))?.why, label).toBe(`${n} is typed into 20 cells. When it changes, every copy has to be found and edited.`);
    }
    const terms = constants({ ...fill('E', 2, 21, (r) => `=D${r}+30`), ...fill('F', 2, 21, (r) => `=D${r}+45`) });
    expect(terms).toMatchObject({ count: 40 });
    expect(terms!.why).toBe('30 and 45 are each typed into several cells. When one changes, every copy has to be found and edited.');
  });

  it('writes small counts in words, so a Fix plan can’t mistake one for a typed number', () => {
    // Fix plans look for the flagged numbers in this sentence. "1.5 is typed into 5 cells" also sent
    // them after the 5 in WEEKDAY(D2,2)>5, a single digit the scan deliberately leaves alone.
    const overtime = constants(fill('E', 2, 6, (r) => `=IF(WEEKDAY(D${r},2)>5,C${r}*1.5,C${r})`));
    expect(overtime?.why).toBe('1.5 is typed into five cells. When it changes, every copy has to be found and edited.');
    const bonus = constants(fill('E', 2, 6, (r) => `=IF(B${r}>5,C${r}*0.05,0)`));
    expect(bonus?.why).toBe('0.05 is typed into five cells. When it changes, every copy has to be found and edited.');
    const average = constants({ ...fill('E', 2, 3, rate), G2: '=SUM(E2:E3)/3*1.0825', H2: '=B2/1.0825' });
    expect(average?.why).toBe('1.0825 is typed into four cells across three formulas. When it changes, every copy has to be found and edited.');
    const many = constants({
      ...fill('E', 2, 21, rate),
      ...fill('F', 2, 21, (r) => `=C${r}*0.15`),
      ...fill('G', 2, 21, (r) => `=C${r}*0.05`),
      ...fill('H', 2, 21, (r) => `=D${r}+30`),
      ...fill('I', 2, 21, (r) => `=D${r}+45`),
    });
    expect(many?.why).toBe('0.05, 0.15, 1.0825 and two other numbers are each typed into several cells. When one changes, every copy has to be found and edited.');
  });

  it('groups the digits of large counts', () => {
    expect(constants(fill('E', 2, 1201, rate))?.why).toBe('1.0825 is typed into 1,200 cells. When it changes, every copy has to be found and edited.');
  });

  it('says "several cells" when a filled number joins one typed into different formulas', () => {
    const cells = { ...fill('E', 2, 21, rate), G2: '=B2*0.15', G3: '=C3*0.15', G4: '=B4/0.15' };
    expect(constants(cells)?.why).toBe('0.15 and 1.0825 are each typed into several cells. When one changes, every copy has to be found and edited.');
  });

  it('leaves common numbers and structural arguments alone', () => {
    const negatives: Array<[string, Cells]> = [
      ['months and percents', { E2: '=B2/12', E3: '=C3*12', E4: '=B4*100', F5: '=C5/100' }],
      ['ROUND digits', { E2: '=ROUND(B2,4)', E3: '=ROUND(C3,4)', E4: '=ROUND(B4*C4,4)' }],
      ['text positions', { E2: '=LEFT(A2,4)', E3: '=MID(A3,4,3)', E4: '=RIGHT(A4,4)&"-"&B4' }],
    ];
    for (const [label, cells] of negatives) expect(constants(cells), label).toBeUndefined();
  });

  it('leaves units, single digits, halves and quarters in long fills alone', () => {
    const forms: Array<[string, (r: number) => string]> = [
      ['thousands', (r) => `=C${r}/1000`],
      ['next row', (r) => `=B${r}+1`],
      ['hours and minutes', (r) => `=(D${r}-C${r})*24*60`],
      ['ROUND digits', (r) => `=ROUND(B${r}*C${r},2)`],
      ['age in years', (r) => `=INT((TODAY()-D${r})/365.25)`],
      ['quarter', (r) => `="Q"&ROUNDUP(MONTH(D${r})/3,0)`],
      ['weekend', (r) => `=IF(WEEKDAY(D${r},2)>5,"Weekend","")`],
      ['trimmed code', (r) => `=LEFT(A${r},LEN(A${r})-3)`],
      ['row number', () => '=ROW()-5'],
      ['due soon', (r) => `=IF(D${r}-TODAY()<=3,"Due soon","")`],
      ['square root', (r) => `=B${r}^0.5`],
      ['round half up', (r) => `=INT(B${r}+0.5)`],
      ['midpoint', (r) => `=(B${r}+C${r})*0.5`],
      ['noon', (r) => `=D${r}+0.5`],
      ['quarter steps', (r) => `=TRUNC(B${r}/0.25)*0.25`],
    ];
    for (const [label, form] of forms) expect(constants(fill('E', 2, 21, form)), label).toBeUndefined();
  });

  it('still counts single digits, halves and quarters typed into different formulas', () => {
    const cells = { E2: '=B2*3', E3: '=C3*3+B3', E4: '=B4/3' };
    expect(constants(cells)?.why).toBe('3 is typed into three formulas. When it changes, every copy has to be found and edited.');
    const deposit = { E2: '=B2*0.5', E3: '=C3*50%', E4: '=B4*C4*0.5' };
    expect(constants(deposit)?.why).toBe('0.5 is typed into three formulas. When it changes, every copy has to be found and edited.');
    const markup = { E2: '=C2*(1+25%)', E3: '=C3*0.25+C3', E4: '=B4*C4*0.25' };
    expect(constants(markup)?.why).toBe('0.25 is typed into three formulas. When it changes, every copy has to be found and edited.');
  });

  it('shares one rule for which numbers count as hard-coded', () => {
    expect(isRepeatedConstant(1.0825, 1, 4), 'one formula, 4 cells').toBe(false);
    expect(isRepeatedConstant(1.0825, 1, 5), 'one formula, 5 cells').toBe(true);
    expect(isRepeatedConstant(1.0825, 2, 4), 'two formulas, 4 cells').toBe(false);
    expect(isRepeatedConstant(1.0825, 3, 3), 'three formulas').toBe(true);
    for (const n of [5, 0.5, 0.25]) {
      expect(isRepeatedConstant(n, 2, 500), `${n} in long fills`).toBe(false);
      expect(isRepeatedConstant(n, 3, 3), `${n} in three formulas`).toBe(true);
    }
  });
});

// ---------- Workbook behaviour ----------

/** One sheet that trips every detector. */
function kitchenSink(): SheetFormulas[] {
  const ops = sheet('Ops', {
    ...orders(),
    D1: 'Extended',
    ...fill('D', 2, 21, (r) => `=B${r}*C${r}`),
    D9: 99,
    D15: '=B15*C14',
    E1: 'Key',
    ...fill('E', 2, 21, (r) => `=A${r}&"|"&B${r}`),
    F1: 'Cost',
    ...fill('F', 2, 21, (r) => `=IFERROR(VLOOKUP(A${r},$H$2:$J$40,3),0)`),
    G1: 'Tier',
    ...fill('G', 2, 21, (r) => `=IF(B${r}>30,"A",IF(B${r}>20,"B",IF(B${r}>10,"C","D")))`),
    L2: '=SUMPRODUCT(($A$2:$A$21="SKU-1004")*$C$2:$C$21)',
    L3: '=SUM(OFFSET($C$2,0,0,20,1))',
    L4: '=SUM(C2:C21)*1.0825',
    L5: '=B2*1.0825',
    L6: '=C3/1.0825',
  });
  return [ops];
}

describe('scanWorkbook', () => {
  it('runs every detector and sorts by severity, then count', () => {
    const findings = scanWorkbook(kitchenSink());
    expect(new Set(findings.map((f) => f.id))).toEqual(
      new Set([
        'vlookup-column-number',
        'vlookup-approximate',
        'iferror-lookup',
        'nested-if',
        'sumproduct-conditions',
        'concatenated-keys',
        'fixed-long-ranges',
        'inconsistent-column',
        'typed-over-formula',
        'volatile-references',
        'hardcoded-constants',
      ]),
    );
    const rank = { high: 0, medium: 1, low: 2 };
    for (let i = 1; i < findings.length; i++) {
      const [a, b] = [findings[i - 1], findings[i]];
      expect(rank[a.severity] < rank[b.severity] || (a.severity === b.severity && a.count >= b.count)).toBe(true);
    }
    expect(findings[0].severity).toBe('high');
    expect(findings[findings.length - 1].severity).toBe('low');
    expect(addresses(finding(findings, 'typed-over-formula'))).toEqual(['D9']);
    expect(addresses(finding(findings, 'inconsistent-column'))).toEqual(['D15']);
  });

  it('writes copy that follows the house rules', () => {
    for (const f of scanWorkbook(kitchenSink())) {
      for (const text of [f.title, f.why, f.fix]) {
        expect(text, f.id).not.toMatch(/\b(please|simply|just|easy|easily|successfully)\b|!/i);
        expect(text, f.id).toMatch(/^[A-Z0-9]/);
      }
      expect(f.title, f.id).not.toMatch(/\.$/);
      expect(f.why, f.id).toMatch(/\.$/);
      expect(f.fix, f.id).toMatch(/\.$/);
      // Sentence case: after the first word, only acronyms and function names are capitalised.
      for (const word of f.title.split(' ').slice(1)) expect(word, f.title).toMatch(/^([a-z0-9-]+|[A-Z0-9]{2,})$/);
      expect(f.cells.length).toBeGreaterThan(0);
      expect(f.cells.length).toBeLessThanOrEqual(8);
    }
  });

  it('skips the coach’s own sheets', () => {
    const practice = sheet('Coach-lookups', { A1: 'x', E2: '=VLOOKUP(A2,$H$2:$J$40,2)' });
    expect(scanWorkbook([practice], { skipPrefix: 'Coach-' })).toEqual([]);
    expect(scanWorkbook([practice])).not.toEqual([]);
    expect(scanWorkbook([practice], { skipPrefix: '' })).not.toEqual([]);
  });

  it('reports absolute addresses when the used range does not start at A1', () => {
    const s = sheet('Report', { C5: 'Header', D7: '=VLOOKUP(C7,$H$2:$J$40,2)' });
    expect(s.address).toBe('C5:D7');
    expect(finding(scan(s), 'vlookup-approximate')?.cells).toEqual([{ sheet: 'Report', address: 'D7' }]);
  });

  it('spreads examples across places and caps them at eight', () => {
    const cells: Cells = { A1: 'x' };
    for (let c = 2; c <= 11; c++) Object.assign(cells, fill(numberToCol(c), 2, 6, (r) => `=VLOOKUP($A${r},$P$2:$Z$40,${c},FALSE)`));
    const f = finding(scan(sheet('Ops', cells)), 'vlookup-column-number');
    expect(f?.count).toBe(50);
    expect(addresses(f)).toEqual(['B2', 'C2', 'D2', 'E2', 'F2', 'G2', 'H2', 'I2']);
  });

  it('copes with empty and odd input', () => {
    expect(scanWorkbook([])).toEqual([]);
    expect(scanWorkbook([{ sheet: 'Blank', address: 'A1', formulas: [['']], r1c1: [['']] }])).toEqual([]);
    expect(scanWorkbook([{ sheet: 'Blank', address: 'A1', formulas: [], r1c1: [] }])).toEqual([]);
    const ragged: SheetFormulas = {
      sheet: 'Odd',
      address: 'not an address',
      formulas: [['=VLOOKUP(A1,B:C,2)'], [], [null, true, 4]],
      r1c1: [],
    };
    expect(finding(scanWorkbook([ragged]), 'vlookup-approximate')?.cells).toEqual([{ sheet: 'Odd', address: 'A1' }]);
  });
});

/** Fast on a quiet Mac (well under 0.5 s); the budget leaves room for parallel runs and CI. */
const SPEED_BUDGET_MS = 2500;

describe('scan speed', () => {
  it('scans a 50,000-cell sheet well under a second', () => {
    const rows = 5000;
    const formulas: Grid = [];
    const r1c1: Grid = [];
    for (let r = 1; r <= rows + 1; r++) {
      const row: Cell[] =
        r === 1
          ? ['SKU', 'Qty', 'Price', 'Extended', 'Key', 'Cost', 'Tier', 'Share', 'Tax', 'Flag']
          : [
              `SKU-${r}`,
              r % 37,
              (r % 11) + 0.5,
              `=B${r}*C${r}`,
              `=A${r}&"|"&B${r}`,
              `=IFERROR(VLOOKUP(A${r},$L$2:$N$400,3,FALSE),0)`,
              `=IF(B${r}>30,"A",IF(B${r}>20,"B",IF(B${r}>10,"C","D")))`,
              `=D${r}/SUM($D$2:$D$5001)`,
              r % 500 === 0 ? 12.5 : `=D${r}*0.0825`,
              `=SUMPRODUCT(($A$2:$A$5001=A${r})*$B$2:$B$5001)`,
            ];
      formulas.push(row);
      r1c1.push(row.map((v, c) => (typeof v === 'string' && v.startsWith('=') ? toR1C1(v, r, c + 1) : v)));
    }
    const input: SheetFormulas = { sheet: 'Big', address: `A1:J${rows + 1}`, formulas, r1c1 };
    const started = performance.now();
    const findings = scanWorkbook([input]);
    const elapsed = performance.now() - started;
    expect(finding(findings, 'typed-over-formula')?.count).toBe(10);
    expect(finding(findings, 'nested-if')?.count).toBe(rows);
    // Generous for busy machines and CI; a quadratic slip would take tens of seconds.
    expect(elapsed).toBeLessThan(SPEED_BUDGET_MS);
  });

  it('stays fast when every formula is different', () => {
    const rows = 5000;
    const formulas: Grid = [];
    for (let r = 1; r <= rows; r++) {
      const row: Cell[] = [];
      for (let c = 0; c < 10; c++) row.push(`=IFERROR(VLOOKUP(A${r}&"-${c}",$L$2:$N$${r + 40},${(c % 3) + 2}),0)*${r}.${c + 1}+IF(B${r}>${r},1,0)`);
      formulas.push(row);
    }
    const started = performance.now();
    const findings = scanWorkbook([{ sheet: 'Unique', address: `A1:J${rows}`, formulas, r1c1: formulas }]);
    const elapsed = performance.now() - started;
    expect(finding(findings, 'vlookup-approximate')).toBeDefined();
    expect(elapsed).toBeLessThan(SPEED_BUDGET_MS * 2);
  });
});

// ---------- Good workbooks: nothing here should be flagged ----------

describe('realistic good workbooks', () => {
  const ids = (s: SheetFormulas) => scan(s).map((f) => f.id);

  it('a P&L column of SUMIFS lines with subtotal and net rows', () => {
    const gl = (r: number) => `=SUMIFS(GL!$D:$D,GL!$A:$A,$A${r},GL!$B:$B,C$1)`;
    const cells: Cells = {
      A1: 'Line', C1: 'Jan', D1: 'Feb',
      A2: 'Revenue', C2: gl(2), D2: gl(2).replace(/C\$1/, 'D$1').replace(/\$A2/, '$A2'),
      A3: 'COGS', C3: gl(3),
      A4: 'Gross profit', C4: '=C2-C3',
      A5: 'Rent', C5: gl(5),
      A6: 'Payroll', C6: gl(6),
      A7: 'Software', C7: gl(7),
      A8: 'Travel', C8: gl(8),
      A9: 'Total opex', C9: '=SUM(C5:C8)',
      A10: 'Operating income', C10: '=C4-C9',
      A11: 'Margin', C11: '=C10/C2',
    };
    expect(ids(sheet('P&L', cells))).not.toContain('inconsistent-column');
  });

  it('a P&L line pulled from another sheet among SUMIFS lines', () => {
    const gl = (r: number) => `=SUMIFS(GL!$D:$D,GL!$A:$A,$A${r})`;
    const cells: Cells = { A1: 'Line', ...fill('C', 2, 7, gl), C5: '=Payroll!F40' };
    expect(ids(sheet('P&L', cells))).not.toContain('inconsistent-column');
  });

  it('a Table calculated column with a totals row', () => {
    const cells: Cells = {
      A1: 'SKU', B1: 'Qty', C1: 'Price', D1: 'Extended',
      ...orders(20),
      ...fill('D', 2, 21, () => '=[@Qty]*[@[Unit price]]'),
      A22: 'Total', D22: '=SUBTOTAL(109,[Extended])',
    };
    const found = ids(sheet('Orders', cells));
    expect(found).not.toContain('inconsistent-column');
    expect(found).not.toContain('typed-over-formula');
  });

  it('a percent-change column whose first row is blank text', () => {
    const cells: Cells = { ...orders(20), E2: '=""', ...fill('E', 3, 21, (r) => `=(C${r}-C${r - 1})/C${r - 1}`) };
    expect(ids(sheet('Trend', cells))).not.toContain('inconsistent-column');
  });

  it('a footer block under a filled column', () => {
    const cells: Cells = {
      ...orders(20),
      ...fill('D', 2, 21, (r) => `=B${r}*C${r}`),
      D22: '=SUM(D2:D21)',
      D23: '=D22*$H$1',
      D24: '=D22+D23',
    };
    expect(ids(sheet('Invoice', cells))).not.toContain('inconsistent-column');
  });

  it('display text joined from two fields', () => {
    const cells: Cells = {
      ...orders(10),
      ...fill('E', 2, 11, (r) => `=A${r}&" "&B${r}`),
      ...fill('F', 2, 11, (r) => `=A${r}&", "&C${r}`),
      ...fill('G', 2, 11, (r) => `="Order "&A${r}&" / "&B${r}`),
    };
    expect(ids(sheet('Labels', cells))).not.toContain('concatenated-keys');
  });

  it('a two-way SUMPRODUCT against month headers', () => {
    expect(one('=SUMPRODUCT(($A$2:$A$60=$P2)*($B$1:$M$1=Q$1)*$B$2:$M$60)', 'sumproduct-conditions')).toBeUndefined();
  });

  it('unit conversions repeated across formulas', () => {
    const cells: Cells = { ...orders(5), E2: '=B2/1000000', E3: '=SUM(C2:C6)/1000000', E4: '=C4/1000000', F2: '=B2/52', F3: '=C3*52', F4: '=B4/52+C4' };
    expect(ids(sheet('Summary', cells))).not.toContain('hardcoded-constants');
  });

  it('running counts and running totals with anchored ranges', () => {
    const cells: Cells = {
      ...orders(40),
      ...fill('D', 2, 41, (r) => `=COUNTIF($A$2:A${r},A${r})`),
      ...fill('E', 2, 41, (r) => `=SUM($B$2:B${r})`),
    };
    const found = ids(sheet('Ops', cells));
    expect(found).not.toContain('fixed-long-ranges');
    expect(found).not.toContain('inconsistent-column');
  });

  it('a modern ops workbook built on Tables and XLOOKUP', () => {
    const lines = sheet('Lines', {
      A1: 'SKU', B1: 'Qty', C1: 'Unit cost', D1: 'Extended', E1: 'Due', F1: 'Status', G1: 'Share',
      ...fill('A', 2, 41, (r) => `SKU-${r}`),
      ...fill('B', 2, 41, (r) => r % 17),
      ...fill('C', 2, 41, () => '=XLOOKUP([@SKU],Items[SKU],Items[Unit cost],0)'),
      ...fill('D', 2, 41, () => '=[@Qty]*[@[Unit cost]]'),
      ...fill('E', 2, 41, () => '=WORKDAY([@Ordered],XLOOKUP([@SKU],Items[SKU],Items[Lead days]),Holidays)'),
      ...fill('F', 2, 41, () => '=IFS([@Due]<TODAY(),"Late",[@Due]-TODAY()<=3,"Due soon",TRUE,"On time")'),
      ...fill('G', 2, 41, () => '=[@Extended]/SUM([Extended])'),
      D42: '=SUBTOTAL(109,[Extended])',
    });
    const summary = sheet('Summary', {
      A1: 'Warehouse', B1: 'Spend', C1: 'Lines', D1: 'Late',
      ...fill('A', 2, 7, (r) => `WH-${r}`),
      ...fill('B', 2, 7, (r) => `=SUMIFS(Lines[Extended],Lines[Warehouse],$A${r})`),
      ...fill('C', 2, 7, (r) => `=COUNTIFS(Lines[Warehouse],$A${r})`),
      ...fill('D', 2, 7, (r) => `=COUNTIFS(Lines[Warehouse],$A${r},Lines[Status],"Late")`),
      B8: '=SUM(B2:B7)', C8: '=SUM(C2:C7)', D8: '=SUM(D2:D7)',
      F2: '=SORT(UNIQUE(Lines[SKU]))',
      H2: '=LET(late,FILTER(Lines[SKU],Lines[Status]="Late",""),ROWS(late))',
    });
    expect(scan(lines, summary)).toEqual([]);
  });

  it('a horizontally filled forecast row', () => {
    const cells: Cells = { A1: 'Month', B2: 1200, ...Object.fromEntries(Array.from({ length: 11 }, (_, i) => [`${numberToCol(i + 3)}2`, `=${numberToCol(i + 2)}2*(1+$O$1)`])) };
    expect(scan(sheet('Forecast', cells))).toEqual([]);
  });
});

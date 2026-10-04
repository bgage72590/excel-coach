import { describe, expect, it } from 'vitest';
import { colToNumber, numberToCol, offsetCell, parseRange, rangeSize, stripSheet } from '../src/engine/address';
import { compareGrids } from '../src/engine/compare';
import { eomonth, excelProper, excelTrim, serial } from '../src/engine/data';
import { disallowedNumbers, forbiddenStrings, scanFormula } from '../src/engine/formula';
import { gradeRules, gradeStructure, gradeValues, gradeVariant, slidingRangeHint } from '../src/engine/grade';
import { MASTERY_PASSES, emptyProgress, recordAttempt, recordPass, statusOf, progressFor } from '../src/engine/progress';
import { Rng } from '../src/engine/rng';

describe('address helpers', () => {
  it('converts columns both ways', () => {
    expect(colToNumber('A')).toBe(1);
    expect(colToNumber('Z')).toBe(26);
    expect(colToNumber('AA')).toBe(27);
    expect(numberToCol(28)).toBe('AB');
    expect(numberToCol(703)).toBe('AAA');
  });
  it('parses sheet-qualified ranges', () => {
    expect(stripSheet("'Coach-x'!K2:K5")).toBe('K2:K5');
    expect(parseRange('Coach!$B$2:$D$21')).toEqual({ start: { row: 2, col: 2 }, end: { row: 21, col: 4 } });
    expect(rangeSize('G2:J7')).toEqual({ rows: 6, cols: 4 });
    expect(offsetCell('G2:J7', 2, 3)).toBe('J4');
  });
});

describe('formula scanner', () => {
  it('finds functions, numbers and strings', () => {
    const s = scanFormula('=SUMIFS(Spend[Amount],Spend[Vendor],$F2,Spend[Date],">="&G$1,Spend[Date],"<="&EOMONTH(G$1,0))');
    expect(s.functions).toEqual(['SUMIFS', 'EOMONTH']);
    expect(s.numbers).toEqual([0]);
    expect(s.strings).toEqual(['>=', '<=']);
  });
  it('ignores digits inside references, structured refs and function names', () => {
    expect(scanFormula('=SUM(A1:B20)+LOG10(C3)+Sales[Q1 2026]+AA100').numbers).toEqual([]);
    expect(scanFormula("='My Sheet 2'!A1*2").numbers).toEqual([2]);
    expect(scanFormula('=F2#*3').numbers).toEqual([3]);
  });
  it('reads percentages and decimals', () => {
    expect(scanFormula('=B2*8%+0.5').numbers).toEqual([0.08, 0.5]);
  });
  it('flags typed numbers and text', () => {
    expect(disallowedNumbers('=IF(B2<10000,B2*0.05,B2*0.08)', [0])).toEqual([10000, 0.05, 0.08]);
    expect(disallowedNumbers('=XLOOKUP(1,(A:A=H2)*(B:B=I2),C:C)', [0, 1])).toEqual([]);
    expect(forbiddenStrings('=SUMIFS(Inventory[Value],Inventory[Warehouse],"Reno")', ['Reno', 'Dallas'])).toEqual(['Reno']);
    expect(forbiddenStrings('=XLOOKUP(I2,Items[SKU],Items[Unit cost],"Discontinued")', ['Reno'])).toEqual([]);
  });
  it('strips new-function prefixes', () => {
    expect(scanFormula('=_xlfn.XLOOKUP(A2,B:B,C:C)').functions).toEqual(['XLOOKUP']);
  });
});

describe('Excel semantics', () => {
  it('matches TRIM and PROPER', () => {
    expect(excelTrim('  maya   rose ')).toBe('maya rose');
    expect(excelProper('NOVAK')).toBe('Novak');
    expect(excelProper("o'brien-SMITH")).toBe("O'Brien-Smith");
  });
  it('computes serial dates and EOMONTH', () => {
    expect(serial(2026, 1, 1)).toBe(46023);
    expect(eomonth(serial(2026, 2, 10))).toBe(serial(2026, 2, 28));
    expect(eomonth(serial(2024, 2, 1))).toBe(serial(2024, 2, 29));
  });
});

describe('grading', () => {
  const area = { kind: 'cells' as const, range: 'I2:I4', consistency: 'all' as const };
  const read = (values: (string | number)[], formulas: string[], r1c1 = formulas) => ({
    address: 'I2:I4',
    values: values.map((v) => [v]),
    formulas: formulas.map((f) => [f]),
    r1c1: r1c1.map((f) => [f]),
  });

  it('passes a clean filled column', () => {
    const r = read([1, 2, 3], ['=A2', '=A3', '=A4'], ['=RC[-8]', '=RC[-8]', '=RC[-8]']);
    const s = gradeStructure(area, r, 3, 1);
    expect(s.items.every((i) => i.status === 'pass')).toBe(true);
    expect(gradeValues(r, [[1], [2], [3]]).item.status).toBe('pass');
  });

  it('flags typed values and odd formulas', () => {
    const r = read([1, 2, 3], ['=A2', '2', '=A4+0'], ['=RC[-8]', '2', '=RC[-8]+0']);
    const s = gradeStructure(area, r, 3, 1);
    expect(s.items.find((i) => i.id === 'formula')?.focus).toBe('I3');
    expect(s.items.find((i) => i.id === 'consistent')?.status).toBe('fail');
  });

  it('explains wrong values and errors', () => {
    const g = gradeValues(read([1, '#N/A', 5], ['=A2', '=A3', '=A4']), [[1], [2], [3]]);
    expect(g.item.status).toBe('fail');
    expect(g.item.detail).toContain('2 cells');
    expect(g.item.detail).toContain('#N/A');
    expect(g.marks.filter((m) => !m.ok).map((m) => m.address)).toEqual(['I3', 'I4']);
  });

  it('requires a single spilling formula', () => {
    const spill = { kind: 'spill' as const, anchor: 'G2' };
    const ok = gradeStructure(spill, { address: 'G2:G4', values: [['a'], ['b'], ['']], formulas: [['=SORT(UNIQUE(A:A))'], ['b'], ['']], r1c1: [[''], [''], ['']] }, 2, 1);
    expect(ok.items[0].status).toBe('pass');
    const bad = gradeStructure(spill, { address: 'G2:G4', values: [['a'], ['b'], ['']], formulas: [['=A2'], ['=A3'], ['']], r1c1: [[''], [''], ['']] }, 2, 1);
    expect(bad.items[0].status).toBe('fail');
  });

  it('allows spilled cells to the right when spillOk', () => {
    const area2 = { kind: 'cells' as const, range: 'B2:D2', consistency: 'columns' as const, spillOk: true };
    const s = gradeStructure(area2, { address: 'B2:D2', values: [['a', 'b', 'c']], formulas: [['=TEXTSPLIT(A2,"|")', 'b', 'c']], r1c1: [['x', 'b', 'c']] }, 1, 3);
    expect(s.items.find((i) => i.id === 'formula')?.status).toBe('pass');
  });

  it('extra spill rows count as mismatches', () => {
    expect(compareGrids([['a'], ['b'], ['c']], [['a'], ['b']]).length).toBe(1);
  });

  it('explains ranges that slide when filled', () => {
    expect(slidingRangeHint(['=SUMIFS(F2:F49,C2:C49,H2)'], 4)).toContain('F3:F50');
    expect(slidingRangeHint(['=SUMIFS($F$2:$F$49,$C$2:$C$49,H2)'], 4)).toBeUndefined();
    expect(slidingRangeHint(['=SUMIFS(Inventory[Value],Inventory[Warehouse],H2)'], 4)).toBeUndefined();
    expect(slidingRangeHint(['=SUM(F2:F41)'], 1)).toBeUndefined();
    expect(slidingRangeHint(['=COUNTIF(A2:A9,"B2:B3")'], 3)).toContain('A3:A10');
  });

  it('leaves one-row ranges alone in a column filled down', () => {
    expect(slidingRangeHint(['=SUM(B4:D4)'], 5, { rows: 5, cols: 1 })).toBeUndefined();
    expect(slidingRangeHint(['=SUM(B4:D4)/SUM(B2:B9)'], 5, { rows: 5, cols: 1 })).toContain('B3:B10');
    expect(slidingRangeHint(['=SUM(C2:C9)'], 4, { rows: 1, cols: 4 })).toBeUndefined();
  });

  it('recognizes Table column names that slid during a fill', async () => {
    const { tableColumnsSlid } = await import('../src/engine/grade');
    expect(tableColumnsSlid('=SUMIFS(Spend[Category],Spend[Date],RC6)', '=SUMIFS(Spend[Amount],Spend[Vendor],RC6)')).toBe(true);
    expect(tableColumnsSlid('=SUMIFS(Spend[Amount],Spend[Vendor],RC7)', '=SUMIFS(Spend[Amount],Spend[Vendor],RC6)')).toBe(false);
    expect(tableColumnsSlid('=A1', undefined)).toBe(false);
  });

  it('applies rules', () => {
    const items = gradeRules({ require: [{ pattern: /LET\(/, label: 'LET', advice: 'use LET' }], allowNumbers: [0] }, ['=ROUNDUP(C2*(D2+E2),0)']);
    expect(items.map((i) => i.status)).toEqual(['fail', 'pass']);
  });

  it('describes variant failures with the explanation', () => {
    const v = gradeVariant(0, 'the table is re-sorted', 'Your lookup depends on row order.', read([1, 9, 3], ['=A2', '=A3', '=A4']), [[1], [2], [3]]);
    expect(v.item.status).toBe('fail');
    expect(v.item.detail).toContain('row order');
    expect(v.badCells).toEqual(['I3']);
  });
});

describe('progress', () => {
  it('masters after three passes on different data and schedules a review', () => {
    let state = emptyProgress();
    const now = Date.UTC(2026, 9, 4);
    for (let i = 0; i < MASTERY_PASSES; i++) {
      state = { ...state, session: { exerciseId: 'x', seed: 100 + i, sheet: 's', startedAt: now - 60_000, attempts: 0, hintsShown: 0, revealed: false } };
      state = recordAttempt(state);
      const out = recordPass(state, now);
      state = out.state;
      expect(out.newlyMastered).toBe(i === MASTERY_PASSES - 1);
    }
    const p = progressFor(state, 'x');
    expect(p.mastered).toBe(true);
    expect(statusOf(p, now)).toBe('mastered');
    expect(statusOf(p, now + 3 * 86_400_000)).toBe('review');
    expect(state.cleanStreak).toBe(3);
  });

  it('does not count a pass after revealing the answer', () => {
    let state = emptyProgress();
    state = { ...state, session: { exerciseId: 'x', seed: 1, sheet: 's', startedAt: 0, attempts: 1, hintsShown: 0, revealed: true } };
    const out = recordPass(state, 1000);
    expect(out.passesTowardMastery).toBe(0);
    expect(out.personalBest).toBe(false);
  });
});

describe('rng', () => {
  it('is deterministic per seed', () => {
    const a = new Rng(42);
    const b = new Rng(42);
    expect([a.next(), a.int(1, 9), a.pick(['x', 'y', 'z'])]).toEqual([b.next(), b.int(1, 9), b.pick(['x', 'y', 'z'])]);
  });
});

describe('platform wording', () => {
  it('starts Power Query from a Table per platform', async () => {
    const { localize } = await import('../src/engine/platform');
    const mac = localize('Choose {fromTable:Shipments}.{macPqNote}', 'mac');
    expect(mac).toContain('Blank query');
    expect(mac).toContain('Excel.CurrentWorkbook(){[Name="Shipments"]}[Content]');
    expect(mac).toContain('no From Table/Range');
    const win = localize('Choose {fromTable:Shipments}.{macPqNote}', 'windows');
    expect(win).toBe('Choose Data › From Table/Range (with a cell in `Shipments` selected).');
    expect(localize('Press {absKey}', 'mac')).toBe('Press ⌘T');
  });

  it('names helper queries on Mac only, where a Blank query starts as Query', async () => {
    const { localize } = await import('../src/engine/platform');
    expect(localize('choose it, {nameQuery:Items}then Close & Load', 'mac')).toBe(
      'choose it, name it Items in Query Settings › Name (a Blank query starts as Query), then Close & Load',
    );
    expect(localize('choose it, {nameQuery:Items}then Close & Load', 'windows')).toBe('choose it, then Close & Load');
  });
});

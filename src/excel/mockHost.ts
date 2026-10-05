import type { RangeRead } from '../engine/fix';
import { parseRange, rangeAddress } from '../engine/address';
import type { SheetFormulas } from '../engine/scan';
import type { StepProbe } from '../engine/guide';
import type { CheckItem, CheckReport, Exercise, Grid, StepDone } from '../engine/types';
import { fixSheetNameFor, type CoachHost } from './host';

const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms));

/**
 * Stand-in for Excel used when the panel runs in a normal browser with ?mock.
 * The first check on a rep fails one variant; the second passes. Design work only.
 */
export class MockHost implements CoachHost {
  readonly live = false;
  private attempts = new Map<string, number>();

  async setup(): Promise<void> {
    await sleep(500);
  }

  async check(ex: Exercise<any>, data: unknown, sheet: string, seed: number): Promise<CheckReport> {
    await sleep(750);
    const key = `${sheet}:${seed}`;
    const n = (this.attempts.get(key) ?? 0) + 1;
    this.attempts.set(key, n);
    const passing = n > 1;

    const items: CheckItem[] = [];
    for (const insp of ex.inspections?.(data) ?? []) items.push({ id: insp.kind, label: insp.label, status: 'pass' });
    const area = ex.layout(data).answer;
    if (area.kind !== 'pivot') {
      items.push({ id: 'formula', label: area.kind === 'spill' ? `One formula in ${area.anchor} returns the whole result` : 'Every answer cell has a formula', status: 'pass' });
      items.push({ id: 'values', label: 'Correct on the current data', status: 'pass' });
      for (const [i, req] of (ex.rules?.require ?? []).entries()) items.push({ id: `require-${i}`, label: req.label, status: 'pass' });
    }
    ex.variants.forEach((v, i) => {
      const fail = !passing && i === 0;
      items.push({
        id: `variant-${i}`,
        label: `Still correct when ${v.label}`,
        status: fail ? 'fail' : 'pass',
        detail: fail ? `K7 showed 4.1 instead of 3.85.${v.explain ? ` ${v.explain}` : ''}` : undefined,
        focus: fail ? 'K7' : undefined,
      });
    });
    if (!passing && ex.variants.length === 0 && items.length) {
      items[items.length - 1] = { ...items[items.length - 1], status: 'fail', detail: 'Rows has Month. Drag Warehouse into Rows, and remove anything else.' };
    }
    const passed = items.every((i) => i.status !== 'fail');
    return { passed, items, marks: [], focus: items.find((i) => i.status === 'fail')?.focus };
  }

  async sheetExists(): Promise<boolean> {
    return true;
  }

  async select(): Promise<void> {}

  /** Steps read as done after a few looks (answers show a wrong-value note first), so the walkthrough can be clicked through. */
  private looks = new Map<string, number>();
  async probe(_ex: Exercise<any>, _data: unknown, sheet: string, done: StepDone): Promise<StepProbe> {
    await sleep(150);
    if (done.kind === 'check') return { done: false };
    const key = `${sheet}:${JSON.stringify(done)}`;
    const n = (this.looks.get(key) ?? 0) + 1;
    this.looks.set(key, n);
    if (n >= 2 && n <= 5 && done.kind === 'answer') {
      return { done: false, note: 'I2 shows $0.00, but it should show $46,211.38.', focus: 'I2', formula: '=SUMIFS(Inventory[Warehouse],Inventory[Value],H2)' };
    }
    return { done: n >= (done.kind === 'answer' ? 6 : 3) };
  }

  async writeInputs(): Promise<void> {
    await sleep(300);
  }

  async copySheetForFix(sheet: string): Promise<string> {
    await sleep(500);
    return fixSheetNameFor(sheet);
  }

  /** Reads from the mock workbook. The copy reads the same as the original, so a fix check passes. */
  async readRange(_sheet: string, address: string): Promise<RangeRead> {
    await sleep(300);
    const book = await this.readWorkbook();
    const sheet = book.sheets[0];
    const want = parseRange(address);
    const origin = parseRange(sheet.address).start;
    const rows = want.end.row - want.start.row + 1;
    const cols = want.end.col - want.start.col + 1;
    const pick = (grid: Grid): Grid =>
      Array.from({ length: rows }, (_, r) => Array.from({ length: cols }, (_, c) => grid[want.start.row - origin.row + r]?.[want.start.col - origin.col + c] ?? ''));
    const formulas = pick(sheet.formulas);
    // Mock values: formulas show a plausible number, everything else shows itself.
    const values = formulas.map((row) => row.map((v) => (typeof v === 'string' && v.startsWith('=') ? 412.5 : v)));
    return { address: rangeAddress(want.start, rows, cols), values, formulas };
  }

  async readWorkbook(): Promise<{ name: string; sheets: SheetFormulas[]; truncated: boolean }> {
    await sleep(600);
    // A small workbook with typical upgrade spots, for designing the My Work screen.
    const rows = 30;
    const formulas: (string | number)[][] = [['SKU', 'Warehouse', 'Qty', 'Cost', 'Lookup', 'Total', 'Key']];
    const r1c1: (string | number)[][] = [['SKU', 'Warehouse', 'Qty', 'Cost', 'Lookup', 'Total', 'Key']];
    for (let r = 2; r <= rows; r++) {
      const typedOver = r === 17;
      formulas.push([`SKU-${100 + r}`, 'Reno', r * 3, 4.5, `=VLOOKUP(A${r},Items!$A$2:$F$400,4,FALSE)`, typedOver ? 412.5 : `=C${r}*D${r}*1.0825`, `=A${r}&B${r}`]);
      r1c1.push([`SKU-${100 + r}`, 'Reno', r * 3, 4.5, '=VLOOKUP(RC[-4],Items!R2C1:R400C6,4,FALSE)', typedOver ? 412.5 : '=RC[-3]*RC[-2]*1.0825', '=RC[-6]&RC[-5]']);
    }
    return { name: 'Weekly ops report.xlsx', truncated: false, sheets: [{ sheet: 'Report', address: `A1:G${rows}`, formulas, r1c1 }] };
  }
}

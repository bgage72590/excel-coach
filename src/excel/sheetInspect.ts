import { stripSheet } from '../engine/address';
import { gradeSheetCheck, gradeValidationList, hiddenTarget, isWholeLines, parseListSource, type ListSource, type SheetFacts, type ValidationFacts } from '../engine/sheetChecks';
import type { CheckItem, Grid, Inspection, SheetCheck } from '../engine/types';

type Ctx = Excel.RequestContext;

/** The ExcelApi set each check reads with, and what to call it when this Excel can't. */
const NEEDS: Record<SheetCheck['kind'], { api: string; what: string }> = {
  freeze: { api: '1.7', what: 'frozen panes' },
  filter: { api: '1.3', what: 'filter buttons' },
  numberFormat: { api: '1.1', what: 'number formats' },
  bold: { api: '1.9', what: 'bold formatting' },
  filled: { api: '1.9', what: 'fill colors' },
  hidden: { api: '1.2', what: 'hidden rows and columns' },
  minWidth: { api: '1.9', what: 'column widths' },
  conditionalFormat: { api: '1.6', what: 'conditional formatting' },
  values: { api: '1.1', what: 'cell values' },
  formulas: { api: '1.1', what: 'formulas' },
};

/** Data validation needs ExcelApi 1.8; resolving a named list needs 1.4, which it implies. */
const VALIDATION_API = '1.8';

/** Office.js error codes for an API this build of Excel doesn't have. */
const MISSING_API = new Set(['ApiNotFound', 'NotImplemented', 'UnsupportedFeature']);

const SHEET_ROWS = 1_048_576;
const SHEET_COLS = 16_384;

/** False only when Office says this Excel lacks the ExcelApi set. Without Office, let the call decide. */
function supports(version: string): boolean {
  const requirements = typeof Office === 'undefined' ? undefined : Office.context?.requirements;
  return requirements ? requirements.isSetSupported('ExcelApi', version) : true;
}

function missingApi(err: unknown): boolean {
  const e = (err ?? {}) as { code?: string; debugInfo?: { code?: string } };
  return MISSING_API.has(e.code ?? e.debugInfo?.code ?? '');
}

const cantReport = (id: string, label: string, what: string): CheckItem => ({ id, label, status: 'skip', detail: `This version of Excel can’t report ${what} to add-ins.` });

/**
 * Converts freezePanes.getLocationOrNullObject() to frozen row and column counts. null means
 * nothing is frozen. Frozen rows alone come back as entire rows (every column), frozen columns
 * alone as entire columns (every row).
 */
export function frozenPanes(location: { rowCount: number; columnCount: number } | null): { rows: number; cols: number } {
  if (!location) return { rows: 0, cols: 0 };
  return { rows: location.rowCount >= SHEET_ROWS ? 0 : location.rowCount, cols: location.columnCount >= SHEET_COLS ? 0 : location.columnCount };
}

/**
 * Host side of the 'sheet' and 'validationList' inspections: reads the facts from Excel and grades
 * them with engine/sheetChecks.ts. `index` makes the item id unique: `sheet-<index>` / `validation-<index>`.
 */
export async function inspectSheet(ctx: Ctx, ws: Excel.Worksheet, insp: Extract<Inspection, { kind: 'sheet' }>, index: number): Promise<CheckItem> {
  const id = `sheet-${index}`;
  const { api, what } = NEEDS[insp.check.kind];
  if (!supports(api)) return cantReport(id, insp.label, what);
  let facts: SheetFacts | null;
  try {
    facts = await readFacts(ctx, ws, insp.check);
  } catch (err) {
    if (missingApi(err)) return cantReport(id, insp.label, what);
    throw err;
  }
  return facts ? { ...gradeSheetCheck(insp, facts), id } : cantReport(id, insp.label, what);
}

/**
 * The cells a cell-by-cell check reads. Whole columns or rows are clipped to the part of the sheet
 * that holds anything, so "F:G" doesn't load a million cells; null when nothing in them does.
 * Formatting alone doesn't count: a stray fill in H40, or a format run down to the last row, would
 * otherwise widen the read and fail a check on the header row.
 */
async function cellsToRead(ctx: Ctx, ws: Excel.Worksheet, address: string): Promise<Excel.Range | null> {
  if (!isWholeLines(address)) return ws.getRange(address);
  const used = ws.getRange(address).getIntersectionOrNullObject(ws.getUsedRange(true));
  await ctx.sync();
  return used.isNullObject ? null : used;
}

/** Reads what one SheetCheck needs. null when this Excel can't say (the sheet's AutoFilter on old builds). */
async function readFacts(ctx: Ctx, ws: Excel.Worksheet, check: SheetCheck): Promise<SheetFacts | null> {
  switch (check.kind) {
    case 'freeze': {
      const location = ws.freezePanes.getLocationOrNullObject();
      location.load(['rowCount', 'columnCount']);
      await ctx.sync();
      return { kind: 'freeze', ...frozenPanes(location.isNullObject ? null : location) };
    }

    case 'filter': {
      // A Table's filter buttons and the sheet's AutoFilter are separate; either one counts.
      const tables = ws.tables;
      tables.load('items/showFilterButton');
      await ctx.sync();
      if (tables.items.some((t) => t.showFilterButton)) return { kind: 'filter', on: true };
      if (!supports('1.9')) return null;
      try {
        ws.autoFilter.load('enabled');
        await ctx.sync();
        return { kind: 'filter', on: ws.autoFilter.enabled };
      } catch {
        return null;
      }
    }

    case 'numberFormat': {
      const range = await cellsToRead(ctx, ws, check.range);
      if (!range) return { kind: 'numberFormat', address: check.range, formats: [] };
      range.load(['address', 'numberFormat']);
      await ctx.sync();
      return { kind: 'numberFormat', address: stripSheet(range.address), formats: range.numberFormat as Grid };
    }

    case 'bold': {
      const range = await cellsToRead(ctx, ws, check.range);
      if (!range) return { kind: 'bold', address: check.range, bold: [] };
      range.load('address');
      const props = range.getCellProperties({ format: { font: { bold: true } } });
      await ctx.sync();
      return { kind: 'bold', address: stripSheet(range.address), bold: props.value.map((row) => row.map((p) => p.format?.font?.bold === true)) };
    }

    case 'filled': {
      const range = await cellsToRead(ctx, ws, check.range);
      if (!range) return { kind: 'filled', address: check.range, colors: [] };
      range.load('address');
      const props = range.getCellProperties({ format: { fill: { color: true } } });
      await ctx.sync();
      return { kind: 'filled', address: stripSheet(range.address), colors: props.value.map((row) => row.map((p) => p.format?.fill?.color ?? null)) };
    }

    case 'hidden': {
      const target = hiddenTarget(check);
      const range = ws.getRange(target.address);
      range.load(target.columns ? 'columnHidden' : 'rowHidden');
      await ctx.sync();
      return { kind: 'hidden', hidden: target.columns ? range.columnHidden : range.rowHidden };
    }

    case 'minWidth': {
      const range = ws.getRange(check.range);
      range.load('address');
      const props = range.getColumnProperties({ format: { columnWidth: true } });
      await ctx.sync();
      // Typed as one entry per column. Take the first row if a build ever returns one row per range row.
      const columns = Array.isArray(props.value[0]) ? (props.value[0] as Excel.ColumnProperties[]) : props.value;
      return { kind: 'minWidth', address: stripSheet(range.address), widths: columns.map((p) => p.format?.columnWidth ?? 0) };
    }

    case 'conditionalFormat': {
      // The collection holds every rule that touches the range; where each applies says whether one covers it.
      const formats = ws.getRange(check.range).conditionalFormats;
      formats.load('items/id');
      await ctx.sync();
      const count = formats.items.length;
      if (!count || !supports('1.9')) return { kind: 'conditionalFormat', count };
      try {
        const areas = formats.items.map((format) => format.getRanges().areas.load('items/address'));
        await ctx.sync();
        return { kind: 'conditionalFormat', count, appliesTo: areas.map((a) => a.items.map((r) => stripSheet(r.address))) };
      } catch {
        return { kind: 'conditionalFormat', count };
      }
    }

    case 'values': {
      const range = await cellsToRead(ctx, ws, check.range);
      if (!range) return { kind: 'values', address: check.range, values: [] };
      range.load(['address', 'values']);
      await ctx.sync();
      return { kind: 'values', address: stripSheet(range.address), values: range.values as Grid };
    }

    case 'formulas': {
      const range = await cellsToRead(ctx, ws, check.range);
      if (!range) return { kind: 'formulas', address: check.range, formulas: [] };
      range.load(['address', 'formulas']);
      await ctx.sync();
      return { kind: 'formulas', address: stripSheet(range.address), formulas: range.formulas as Grid };
    }
  }
}

export async function inspectValidation(ctx: Ctx, ws: Excel.Worksheet, insp: Extract<Inspection, { kind: 'validationList' }>, index: number): Promise<CheckItem> {
  const id = `validation-${index}`;
  const what = 'data validation';
  if (!supports(VALIDATION_API)) return cantReport(id, insp.label, what);

  const validation = ws.getRange(insp.cell).dataValidation;
  try {
    validation.load(['type', 'rule']);
    await ctx.sync();
  } catch (err) {
    if (missingApi(err)) return cantReport(id, insp.label, what);
    throw err;
  }

  const type = String(validation.type);
  const list = type === 'List' ? validation.rule?.list : undefined;
  const source = typeof list?.source === 'string' ? list.source : '';
  const facts: ValidationFacts = { type, source, inCellDropDown: list?.inCellDropDown };
  if (source.trim().startsWith('=')) Object.assign(facts, await readListSource(ctx, ws, parseListSource(source)));
  return { ...gradeValidationList(insp, facts), id };
}

type ListRead = Pick<ValidationFacts, 'sourceValues' | 'sourceMissing'>;

/**
 * The cells a list rule's reference points at, as they display: a range on this or another sheet,
 * a spill, or a defined name (the practice sheet's own first, then the workbook's, the order Excel
 * resolves them in). `sourceMissing` says what isn't there; neither field is set when the reference
 * can't be evaluated, which the grader explains.
 */
async function readListSource(ctx: Ctx, ws: Excel.Worksheet, src: ListSource): Promise<ListRead> {
  if (src.kind === 'items' || src.kind === 'formula') return {};
  // range.text: the dropdown offers each cell as it displays, so 5% rather than 0.05.
  const shown = (range: Excel.Range): ListRead => (range.isNullObject ? {} : { sourceValues: range.text as Grid });
  try {
    let sheet = ws;
    if (src.sheet !== undefined) {
      const other = ctx.workbook.worksheets.getItemOrNullObject(src.sheet);
      await ctx.sync();
      if (other.isNullObject) return { sourceMissing: 'sheet' };
      sheet = other;
    }

    if (src.kind === 'name') {
      for (const names of src.sheet !== undefined ? [sheet.names] : [ws.names, ctx.workbook.names]) {
        const named = names.getItemOrNullObject(src.name);
        await ctx.sync();
        if (named.isNullObject) continue;
        // A name for a constant or a formula has no range to read.
        const range = named.getRangeOrNullObject();
        range.load('text');
        await ctx.sync();
        return shown(range);
      }
      return { sourceMissing: 'name' };
    }

    if (src.spill) {
      if (!supports('1.12')) return {};
      const range = sheet.getRange(src.address).getSpillingToRangeOrNullObject();
      range.load('text');
      await ctx.sync();
      return range.isNullObject ? { sourceMissing: 'spill' } : shown(range);
    }

    // Whole columns or rows: read only the part that holds anything.
    const range = isWholeLines(src.address) ? sheet.getRange(src.address).getIntersectionOrNullObject(sheet.getUsedRange(true)) : sheet.getRange(src.address);
    range.load('text');
    await ctx.sync();
    return range.isNullObject ? { sourceValues: [] } : shown(range);
  } catch {
    return {};
  }
}

import type { Rng } from './rng';

/** A single cell value as Excel reports it. Empty cells come back as "". */
export type Cell = string | number | boolean | null;
export type Grid = Cell[][];

/** An expected cell may be a matcher instead of a literal, e.g. any "Total" label. */
export interface CellMatcher {
  match: RegExp;
  describe: string;
}
export type ExpectedCell = Cell | CellMatcher;
export type ExpectedGrid = ExpectedCell[][];

export type Platform = 'mac' | 'windows' | 'web';

export type ModuleId =
  | 'tables'
  | 'sumifs'
  | 'lookups'
  | 'dynamic'
  | 'cleaning'
  | 'letlambda'
  | 'pivots'
  | 'powerquery'
  | 'whatif'
  | 'charts'
  | 'dates'
  | 'modeling'
  | 'bughunt'
  | 'macros'
  | 'drills'
  | 'missions';

export interface ModuleInfo {
  id: ModuleId;
  title: string;
  blurb: string;
}

// ---------- sheet layout ----------

export interface ColumnSpec {
  header: string;
  /** Excel number format, e.g. "$#,##0.00" or "yyyy-mm-dd". */
  format?: string;
}

/** Tabular data written with a header row. `asTable: false` leaves it as a plain range. */
export interface DataBlock {
  kind: 'data';
  at: string;
  table: string;
  columns: ColumnSpec[];
  rows: Grid;
  asTable: boolean;
}

/** Free-form cells: labels, inputs, or (with formulas: true) prewritten formulas. */
export interface CellsBlock {
  kind: 'cells';
  at: string;
  values: Grid;
  role: 'label' | 'header' | 'input' | 'formula';
  format?: string;
  /** Per-column number formats; overrides `format` for the columns it names. */
  formats?: (string | undefined)[];
}

export type Block = DataBlock | CellsBlock;

export type AnswerArea =
  /**
   * One formula per cell. `consistency`: 'all' = one formula through the whole range, 'columns' =
   * one per column (filled down), 'rows' = one per row (filled across, e.g. a 12-month forecast
   * line), 'none' = no consistency check.
   * `spillOk` lets cells be filled by a spill from the left (e.g. TEXTSPLIT).
   * `liveValues`: the cells come from a what-if data table, which Office.js reports as plain values
   * (no =TABLE formula), so the variants prove the grid is live instead of the formula text.
   */
  | { kind: 'cells'; range: string; format?: string; consistency: 'all' | 'columns' | 'rows' | 'none'; spillOk?: boolean; liveValues?: boolean }
  /** One formula at the anchor that spills the whole result. */
  | { kind: 'spill'; anchor: string; format?: string; formats?: (string | undefined)[] }
  /** A calculated column the learner adds to a Table. */
  | { kind: 'tableColumn'; table: string; column: string; format?: string }
  /** A PivotTable built anywhere in the workbook; checked structurally. */
  | { kind: 'pivot' }
  /**
   * A table loaded by Power Query anywhere in the workbook. The checker finds the table whose
   * header row holds exactly `columns` (any order, case-insensitive), confirms a query loaded it,
   * and compares its rows with expected(d). `order: 'any'` ignores row order.
   */
  | { kind: 'query'; columns: string[]; order: 'any' | 'asis' }
  /** Individual cells checked one by one: Goal Seek inputs, Solver decisions, constraint cells. */
  | { kind: 'cellChecks'; checks: CellCheck[] }
  /** Structure only: charts, slicers and similar objects checked by inspections. */
  | { kind: 'objects' }
  /**
   * Bug hunt: the coach prewrites formulas in `range` with planted mistakes (a 'formula' block).
   * expected(d) is the CORRECT grid for `range`. The checker reads `range` on the current data and
   * under every variant (whether or not the current data passes, because a typed-in number only
   * shows once the data changes). A bug counts as fixed when all of its cells match in every run.
   * Cells outside every bug must keep matching too, so a "fix" that breaks a correct cell fails.
   * The answer area is not styled yellow: the sheet should look like a finished workbook.
   */
  | { kind: 'bugHunt'; range: string; bugs: PlantedBug[] };

/** One mistake planted in a bug-hunt workbook. */
export interface PlantedBug {
  /** Stable id within the exercise, e.g. "hardcoded-rate". */
  id: string;
  /** Shown once fixed, e.g. "The tax rate was typed into the Total formulas". Past tense, specific. */
  label: string;
  /** Every cell (A1, on the practice sheet, inside `range`) whose value is wrong because of this bug. */
  cells: string[];
}

/**
 * One cell's requirement. All fields are optional; each one present is enforced.
 * `answer: true` styles the cell as a place to type (yellow) at setup.
 */
export interface CellCheck {
  cell: string;
  label: string;
  answer?: boolean;
  /** 'number': must hold a typed number, not a formula. 'formula': must hold a formula. */
  holds?: 'number' | 'formula';
  /** The value must be within `tolerance` (default 0.005) of this number. */
  value?: number;
  tolerance?: number;
  min?: number;
  max?: number;
  integer?: boolean;
  /** The cell must still contain this formula (compared ignoring case and spaces). */
  formula?: string;
  /** Shown when the check fails. */
  advice?: string;
}

export interface Layout {
  blocks: Block[];
  answer: AnswerArea;
  /** Other answer areas to style at setup (later steps of a mission). */
  alsoStyle?: AnswerArea[];
  /** Header/label cell shown above or beside the answer area, if any. */
  answerLabel?: { at: string; values: Grid };
  /**
   * Leave the sheet unformatted: no autofit, no frozen header row, and no header styling on
   * non-Table data blocks. For exercises that check formatting the learner applies (macros, drills).
   */
  plain?: boolean;
}

// ---------- perturbation ----------

/** Input regions the checker may rewrite to test a formula, then restore. */
export type InputWrite =
  | { kind: 'range'; address: string; values: Grid }
  | { kind: 'table'; table: string; columns: string[]; rows: Grid };

export interface Variant<D> {
  /** Completes "Still correct when …". */
  label: string;
  /** Shown when the formula breaks under this change, e.g. "Your lookup depends on row order." */
  explain?: string;
  apply(d: D, rng: Rng): D;
}

// ---------- rules and structural checks ----------

export interface RequireRule {
  pattern: RegExp;
  /** Shown as the check label, e.g. "Uses the Table's column name". */
  label: string;
  /** Shown when it fails. */
  advice: string;
}

export interface Rules {
  require?: RequireRule[];
  /** Typed-in numbers are flagged unless listed here. Omit to allow any number. */
  allowNumbers?: number[];
  /** Text that must come from cells, not be typed into the formula. */
  forbidText?: { values: string[]; advice: string };
}

export type Inspection =
  | { kind: 'tableExists'; table: string; at: string; rows: number; label: string }
  | { kind: 'tableColumn'; table: string; column: string; label: string }
  | { kind: 'lambdaName'; name: string; params: number; label: string }
  | {
      kind: 'pivot';
      rows: string;
      /** Omit or '' when nothing should be in Columns. */
      columns?: string;
      /** A field that must be in the Filters area. */
      filter?: string;
      valuesField: string;
      summarizeBy: 'Sum' | 'Count' | 'Average' | 'Max' | 'Min';
      showAs: 'PercentOfRowTotal' | 'PercentOfColumnTotal' | 'PercentOfGrandTotal' | 'None';
      label: string;
    }
  /**
   * A chart on the practice sheet. `secondaryLine`: one series is a line on the secondary axis
   * (a combo chart). `title`: the chart has a visible title other than the default "Chart Title".
   */
  | { kind: 'chart'; label: string; minSeries?: number; types?: RegExp; secondaryLine?: boolean; title?: boolean }
  /** A slicer whose caption names `field` (Table or PivotTable slicer). */
  | { kind: 'slicer'; field: string; label: string }
  /**
   * A data-validation dropdown (a List rule) on `cell` offering exactly `options` (any order,
   * case-insensitive). The list may be typed into the rule or point at cells holding the options.
   */
  | { kind: 'validationList'; cell: string; options: string[]; label: string }
  /** A fact about the practice sheet's formatting or state; see SheetCheck. */
  | { kind: 'sheet'; check: SheetCheck; label: string; advice?: string };

/**
 * Sheet facts the checker can read through Office.js, for macros, drills and modeling. Ranges are
 * A1 addresses on the practice sheet. Graded by gradeSheetCheck in engine/sheetChecks.ts from the
 * facts the host reads (host code in excel/sheetInspect.ts).
 */
export type SheetCheck =
  /** Exactly `rows` rows and `cols` columns frozen. */
  | { kind: 'freeze'; rows: number; cols: number }
  /** Filter buttons are on (the sheet's AutoFilter, or a Table on the sheet with its filter buttons showing). */
  | { kind: 'filter'; on: boolean }
  /** Every cell in `range` has a number format matching `matches`. `describe`: e.g. "currency with 2 decimals". */
  | { kind: 'numberFormat'; range: string; matches: RegExp; describe: string }
  /** Every cell in `range` is bold (or, with bold: false, none is). */
  | { kind: 'bold'; range: string; bold: boolean }
  /** Every cell in `range` has a fill color (or, with filled: false, none has). White counts as no fill. */
  | { kind: 'filled'; range: string; filled: boolean }
  /** The whole columns (e.g. "D:E") or rows (e.g. "5:9") are hidden, or with hidden: false, visible. */
  | { kind: 'hidden'; columns?: string; rows?: string; hidden: boolean }
  /** Every column in `range` is at least `points` wide (autofit, or set by hand). */
  | { kind: 'minWidth'; range: string; points: number }
  /** At least one conditional format applies to `range`. */
  | { kind: 'conditionalFormat'; range: string }
  /** The cells in `range` show these values (compared like answers). `describe`: e.g. "the totals row". */
  | { kind: 'values'; range: string; expected: ExpectedGrid; describe: string }
  /** Every cell in `range` holds a formula (optionally matching `pattern`), or with formulas: false, none does. */
  | { kind: 'formulas'; range: string; formulas: boolean; pattern?: RegExp };

// ---------- exercises ----------

export interface Concept {
  summary: string;
  syntax: string;
  example?: string;
  tip?: string;
}

// ---------- guided walkthrough ----------

/**
 * Where on the practice sheet something lives, for the coach to select: an A1 range ("F2:F49"),
 * several ranges ("F6,F9,F12"), or a Table reference ("Inventory[Value]", "Inventory[#Data]",
 * "Inventory[#All]"). Table references follow the Table as it grows.
 */
export type SheetSpot = string;

/** One piece of a formula in a walkthrough. A step's parts, joined, spell the whole formula. */
export interface FormulaPart {
  text: string;
  /** Plain words, e.g. "what to add up: the Value column". Omit for punctuation like "," or ")". */
  means?: string;
  /** Where it is on the sheet; the learner can tap Show to select it. */
  at?: SheetSpot;
}

/** A place the coach can point at by selecting it. */
export interface SheetPointer {
  /** Button text, e.g. "Show the Dallas rows". */
  label: string;
  at: SheetSpot;
  /** Shown once selected, e.g. "The status bar's Sum is the number your formula should give." */
  note?: string;
}

/** How the coach spots that a walkthrough step is done. Without one, the learner presses Next. */
export type StepDone =
  /** The active cell is inside `range` on the practice sheet. */
  | { kind: 'select'; range: string }
  /**
   * Answer cells hold formulas showing the expected values. `cells` (A1, inside a 'cells' answer
   * area) narrows it to part of the area, e.g. the first cell before filling down. Spill and
   * tableColumn answers are always judged whole.
   */
  | { kind: 'answer'; cells?: string }
  /** An inspection passes: a Table exists, a PivotTable is laid out, a chart is there. */
  | { kind: 'inspect'; inspection: Inspection }
  /** Some Table, whatever its name, starts at `at` on the practice sheet (before the learner renames it). */
  | { kind: 'tableAt'; at: string }
  /** The check passes. The last step. */
  | { kind: 'check' };

/** One small action in a walkthrough. */
export interface GuideStep {
  /** The action, e.g. "Click `I2` and type this formula, then press {enter}." Placeholders allowed. */
  do: string;
  /** Why it works, or what to expect on screen. */
  why?: string;
  /** The formula to type in this step, part by part. */
  formula?: FormulaPart[];
  /** Places to point at: each becomes a Show button. */
  show?: SheetPointer[];
  done?: StepDone;
}

export interface Exercise<D = unknown> {
  id: string;
  module: ModuleId;
  title: string;
  /** "Replaces …" line: the manual workflow this skill retires. */
  replaces: string;
  minutes: number;
  /** Requires a Microsoft 365-only function. */
  m365?: boolean;
  task(d: D): string;
  concept: Concept;
  /** Progressive hints. `{tableKey}`, `{absKey}`, `{nameManager}` are replaced per platform. */
  hints: string[];
  solution(d: D): string;
  make(rng: Rng): D;
  layout(d: D): Layout;
  /** Expected answer values: the answer grid, or for 'query' the output rows without the header. Unused for pivot, cellChecks and objects. */
  expected(d: D): ExpectedGrid;
  /** Regions the checker rewrites for each variant. */
  inputs(d: D): InputWrite[];
  variants: Variant<D>[];
  rules?: Rules;
  /** Structural checks that need Excel (a Table exists, a LAMBDA is defined, a PivotTable's layout). */
  inspections?(d: D): Inspection[];
  /** Workbook names this exercise creates; removed when a fresh sheet is set up. */
  ownsNames?: string[];
  /** A step-by-step walkthrough for learning the skill. A guided rep doesn't count toward mastery. */
  guide?(d: D): GuideStep[];
}

/** A concrete rep: an exercise plus the data generated from a seed. */
export interface Instance {
  exerciseId: string;
  seed: number;
  sheet: string;
}

// ---------- check results ----------

export type CheckStatus = 'pass' | 'fail' | 'skip';

export interface CheckItem {
  id: string;
  label: string;
  status: CheckStatus;
  detail?: string;
  /** First cell to look at, as an A1 address on the practice sheet. */
  focus?: string;
}

export interface CellMark {
  address: string;
  ok: boolean;
}

export interface CheckReport {
  passed: boolean;
  items: CheckItem[];
  marks: CellMark[];
  focus?: string;
  /** The learner's distinct formulas from the answer area, when there is one. Used for AI hints. */
  formulas?: string[];
}

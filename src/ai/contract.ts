/**
 * The wire contract between the panel and the local hint server (server/hintPlugin.ts).
 *
 * Privacy: this is everything a hint request carries. Hints are only offered on practice sheets,
 * which hold synthetic data the coach generated, never the learner's own workbook. Check details
 * can quote a value from the practice sheet ("K7 showed 4.1 instead of 3.85"); nothing else from
 * the sheet is sent.
 */

/**
 * Mirrors CheckStatus in engine/types.ts. Kept separate so the server loads this file without the
 * engine; buildHintPayload stops compiling if the two drift apart.
 */
export type HintCheckStatus = 'pass' | 'fail' | 'skip';

/** Mirrors Platform in engine/types.ts, the same way. Claude needs it to name the right menus and keys. */
export type HintPlatform = 'mac' | 'windows' | 'web';

export interface HintCheck {
  label: string;
  status: HintCheckStatus;
  detail?: string;
}

export interface HintPayload {
  /** Exercise title, or "Mission · Step" for a mission step. */
  title: string;
  /** The task text, localized for the learner's platform. */
  task: string;
  /** Which Excel the learner is in, so menu paths and shortcuts match theirs. */
  platform: HintPlatform;
  /** The learner's distinct formulas from the answer area (CheckReport.formulas), in reading order. */
  formulas: string[];
  checks: HintCheck[];
  /** The coach's own hints the learner has already opened, in order. */
  hintsShown: string[];
}

export type HintStatusResponse = { enabled: boolean };
export type HintResponse = { hint: string } | { error: string };

export const HINT_ROUTE = '/api/hint';
export const HINT_STATUS_ROUTE = '/api/hint/status';

/** Largest request body the server reads. The limits below keep a clamped payload well under it. */
export const HINT_BODY_LIMIT = 16 * 1024;

/** Caps applied on both sides: the panel trims before sending, the server trims again. */
export const HINT_LIMITS = {
  title: 120,
  task: 1200,
  formulas: 10,
  formula: 500,
  checks: 12,
  label: 160,
  detail: 240,
  hints: 5,
  hint: 300,
} as const;

const STATUSES: readonly HintCheckStatus[] = ['pass', 'fail', 'skip'];
const PLATFORMS: readonly HintPlatform[] = ['mac', 'windows', 'web'];

export const isCheckStatus = (v: unknown): v is HintCheckStatus => STATUSES.includes(v as HintCheckStatus);
export const isHintPlatform = (v: unknown): v is HintPlatform => PLATFORMS.includes(v as HintPlatform);

/** Cuts text to `max` characters, ending in an ellipsis when anything was dropped. */
export function clip(text: string, max: number): string {
  const t = text.trim();
  return t.length <= max ? t : `${t.slice(0, max - 1).trimEnd()}…`;
}

// String literals (including quoted sheet names) are skipped, so a name like 'Q1 Data' isn't read
// as a cell reference.
const LITERAL = /("(?:[^"]|"")*"|'(?:[^']|'')*')/;
const A1_REF = /(^|[^A-Za-z0-9_.])(\$?)([A-Z]{1,3})(\$?)(\d+)(?![A-Za-z0-9_(])/g;

const columnNumber = (letters: string) => [...letters].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);

/**
 * A formula with each relative reference written as its offset from the first one, like R1C1 but
 * without knowing the formula's own cell. Every copy of a filled formula has the same shape; a
 * typed-over cell or a reference that slipped a row doesn't. $-anchored parts stay as written.
 */
export function formulaShape(formula: string): string {
  let row0: number | undefined;
  let col0: number | undefined;
  return formula
    .split(LITERAL)
    .map((part, i) =>
      i % 2 === 1
        ? part
        : part.replace(A1_REF, (_, lead: string, colAbs: string, col: string, rowAbs: string, row: string) => {
            const c = colAbs ? `$${col}` : `C[${columnNumber(col) - (col0 ??= columnNumber(col))}]`;
            const r = rowAbs ? `$${row}` : `R[${Number(row) - (row0 ??= Number(row))}]`;
            return lead + c + r;
          }),
    )
    .join('');
}

/**
 * Picks at most `max` formulas. A filled column gives one distinct formula per row, so the first
 * `max` would be the same formula over and over; instead one formula of each shape is kept first,
 * rarest shapes first, so the odd one out always makes it. Leftover room is filled in order, and the
 * picks keep their reading order.
 */
export function pickFormulas(formulas: string[], max: number): string[] {
  if (formulas.length <= max) return formulas;
  const shapes = new Map<string, number[]>();
  formulas.forEach((f, i) => {
    const key = formulaShape(f);
    const group = shapes.get(key);
    if (group) group.push(i);
    else shapes.set(key, [i]);
  });
  const firsts = [...shapes.values()].sort((a, b) => a.length - b.length || a[0] - b[0]).map((group) => group[0]);
  const kept = new Set(firsts.slice(0, max));
  for (let i = 0; kept.size < max; i++) kept.add(i);
  return formulas.filter((_, i) => kept.has(i));
}

/**
 * Applies HINT_LIMITS. Failed checks are kept ahead of passing ones when the list is too long,
 * since they're what the explanation is about; the original order is otherwise preserved.
 */
export function clampHintPayload(p: HintPayload): HintPayload {
  const ranked = p.checks.map((c, i) => ({ c, i })).sort((a, b) => Number(b.c.status === 'fail') - Number(a.c.status === 'fail') || a.i - b.i);
  const kept = ranked
    .slice(0, HINT_LIMITS.checks)
    .sort((a, b) => a.i - b.i)
    .map(({ c }) => ({
      label: clip(c.label, HINT_LIMITS.label),
      status: c.status,
      ...(c.detail?.trim() ? { detail: clip(c.detail, HINT_LIMITS.detail) } : {}),
    }));
  return {
    title: clip(p.title, HINT_LIMITS.title),
    task: clip(p.task, HINT_LIMITS.task),
    platform: p.platform,
    formulas: pickFormulas(p.formulas.map((f) => f.trim()).filter(Boolean), HINT_LIMITS.formulas).map((f) => clip(f, HINT_LIMITS.formula)),
    checks: kept,
    // Hints get more specific as they go, so the latest ones are the ones worth keeping.
    hintsShown: p.hintsShown
      .map((h) => h.trim())
      .filter(Boolean)
      .slice(-HINT_LIMITS.hints)
      .map((h) => clip(h, HINT_LIMITS.hint)),
  };
}

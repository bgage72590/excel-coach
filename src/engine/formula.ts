/**
 * Lightweight formula scanner. It does not evaluate anything (Excel does that);
 * it finds what a learner typed: functions used, literal numbers and literal text.
 */

export interface FormulaScan {
  /** Formula without the leading "=", with _xlfn./_xlws. prefixes removed. */
  body: string;
  /** Upper-cased function names in order of appearance. */
  functions: string[];
  /** Numeric literals typed into the formula (percentages already divided by 100). */
  numbers: number[];
  /** Text literals typed into the formula (without quotes). */
  strings: string[];
}

export function isFormula(value: unknown): value is string {
  return typeof value === 'string' && value.startsWith('=');
}

export function normalizeFormula(formula: string): string {
  return formula
    .replace(/^=/, '')
    .replace(/_xlfn\.|_xlws\.|_xlpm\./gi, '')
    .trim();
}

export function scanFormula(formula: string): FormulaScan {
  const body = normalizeFormula(formula);
  const strings: string[] = [];

  // Pull out "text literals" (Excel escapes a quote as "").
  let code = '';
  for (let i = 0; i < body.length; i++) {
    const ch = body[i];
    if (ch === '"') {
      let j = i + 1;
      let text = '';
      while (j < body.length) {
        if (body[j] === '"') {
          if (body[j + 1] === '"') {
            text += '"';
            j += 2;
            continue;
          }
          break;
        }
        text += body[j];
        j++;
      }
      strings.push(text);
      code += ' "" ';
      i = j;
    } else {
      code += ch;
    }
  }

  // Structured references: remove bracketed parts, innermost first.
  let prev = '';
  while (prev !== code) {
    prev = code;
    code = code.replace(/\[[^[\]]*\]/g, ' ');
  }

  // Sheet prefixes, then A1 references (including whole rows/columns like A:A or 2:2).
  code = code
    .replace(/'[^']*'!/g, ' ')
    .replace(/\b[A-Za-z_][\w.]*!/g, ' ')
    .replace(/\$?\b[A-Za-z]{1,3}\$?\d+\b#?/g, ' ')
    .replace(/\$?\b[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}\b/g, ' ')
    .replace(/\$?\b\d+:\$?\d+\b/g, ' ');

  const functions: string[] = [];
  for (const m of code.matchAll(/([A-Za-z_][A-Za-z0-9_.]*)\s*\(/g)) functions.push(m[1].toUpperCase());

  const numbers: number[] = [];
  const numberPattern = /(?<![A-Za-z0-9_.])(\d+(?:\.\d+)?(?:[eE][+-]?\d+)?)(%?)(?![A-Za-z0-9_(.])/g;
  for (const m of code.matchAll(numberPattern)) {
    const n = Number(m[1]);
    numbers.push(m[2] ? n / 100 : n);
  }

  return { body, functions, numbers, strings };
}

/** Numbers in the formula that aren't in the allowed list. */
export function disallowedNumbers(formula: string, allowed: readonly number[]): number[] {
  return scanFormula(formula).numbers.filter((n) => !allowed.some((a) => Math.abs(a - n) < 1e-12));
}

/** Text literals that match any of `forbidden` (case-insensitive, trimmed). */
export function forbiddenStrings(formula: string, forbidden: readonly string[]): string[] {
  const wanted = new Set(forbidden.map((s) => s.trim().toLowerCase()));
  return scanFormula(formula).strings.filter((s) => wanted.has(s.trim().toLowerCase()));
}

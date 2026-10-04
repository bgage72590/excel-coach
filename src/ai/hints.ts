import type { CheckReport, Platform } from '../engine/types';
import { HINT_ROUTE, HINT_STATUS_ROUTE, clampHintPayload, type HintPayload } from './contract';

/**
 * The panel's side of AI hints. The panel never holds an API key: it asks the local dev server,
 * which adds the key and calls Claude (server/hintPlugin.ts).
 */

/** Shown when the request fails before the server can answer. */
export const CLIENT_MESSAGES = {
  unreachable: 'Couldn’t reach the hint server. Check that npm run dev is still running.',
  timeout: 'The explanation took too long. Try again.',
  failed: 'Something went wrong getting an explanation. Try again.',
} as const;

// A little longer than the server's own 20 s limit, so the server's message wins when it can.
const CLIENT_TIMEOUT_MS = 25_000;

export class HintError extends Error {
  override name = 'HintError';
}

let status: Promise<boolean> | undefined;
let known: boolean | undefined;

/**
 * Whether hints are available: a dev server with an API key. Asked once per panel load; any
 * failure counts as off. A static build (GitHub Pages) has no hint server, so it doesn't ask.
 */
export function hintStatus(): Promise<boolean> {
  status ??= (async () => {
    if (!import.meta.env.DEV) return false;
    try {
      const res = await fetch(HINT_STATUS_ROUTE, { headers: { Accept: 'application/json' } });
      if (!res.ok) return false;
      const body: unknown = await res.json();
      return typeof body === 'object' && body !== null && (body as { enabled?: unknown }).enabled === true;
    } catch {
      return false;
    }
  })().then((on) => (known = on));
  return status;
}

/** The status, if it's already known; lets a view render the button without a flash. */
export const knownHintStatus = (): boolean | undefined => known;

/** Whether a report has something to explain. The server refuses a payload with no failed check. */
export const canExplain = (report: CheckReport): boolean => report.items.some((i) => i.status === 'fail');

/** Builds the request from what the learner sees: the task, the hints opened, and the last check. */
export function buildHintPayload({
  title,
  task,
  platform,
  report,
  hintsShown,
}: {
  title: string;
  task: string;
  platform: Platform;
  report: CheckReport;
  hintsShown: string[];
}): HintPayload {
  return clampHintPayload({
    title,
    task,
    platform,
    formulas: report.formulas ?? [],
    checks: report.items.map(({ label, status, detail }) => ({ label, status, ...(detail ? { detail } : {}) })),
    hintsShown,
  });
}

/** Asks for an explanation. Resolves with the hint text; rejects with a HintError to show as is. */
export async function requestHint(payload: HintPayload, signal?: AbortSignal): Promise<string> {
  const ctrl = new AbortController();
  const timer = setTimeout(() => ctrl.abort(new HintError(CLIENT_MESSAGES.timeout)), CLIENT_TIMEOUT_MS);
  const stop = () => ctrl.abort(signal?.reason);
  if (signal?.aborted) stop();
  else signal?.addEventListener('abort', stop, { once: true });

  let ok = false;
  let body: unknown;
  try {
    const res = await fetch(HINT_ROUTE, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', Accept: 'application/json' },
      body: JSON.stringify(payload),
      signal: ctrl.signal,
    });
    ok = res.ok;
    // Anything but JSON means something other than the hint server answered.
    body = await res.json();
  } catch (e) {
    if (signal?.aborted) throw e;
    throw ctrl.signal.reason instanceof HintError ? ctrl.signal.reason : new HintError(CLIENT_MESSAGES.unreachable);
  } finally {
    clearTimeout(timer);
    signal?.removeEventListener('abort', stop);
  }

  const fields = (typeof body === 'object' && body !== null ? body : {}) as { hint?: unknown; error?: unknown };
  if (ok && typeof fields.hint === 'string' && fields.hint.trim()) return fields.hint;
  throw new HintError(typeof fields.error === 'string' && fields.error ? fields.error : CLIENT_MESSAGES.failed);
}

import { Readable } from 'node:stream';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import {
  HINT_MAX_TOKENS,
  HINT_MODEL,
  MESSAGES,
  SYSTEM_PROMPT,
  buildApiRequest,
  buildUserMessage,
  explainMistake,
  handleHintRequest,
  isLoopbackAddress,
  isLoopbackHost,
  isLoopbackOrigin,
  parseHintPayload,
  readBody,
  tidyHint,
  type HintDeps,
  type HintHttpRequest,
} from '../server/hint';
import { HINT_BODY_LIMIT, HINT_LIMITS, clampHintPayload, formulaShape, pickFormulas, type HintPayload } from '../src/ai/contract';
import type { CheckReport } from '../src/engine/types';

const KEY = 'sk-ant-api03-test-key-0123456789abcdef';
const BANNED = /\b(please|simply|just|easy|easily|successfully|leverage|seamless)\b|!/i;

const payload = (over: Partial<HintPayload> = {}): HintPayload => ({
  title: 'Look up prices with XLOOKUP',
  task: 'In `F2:F31`, return each SKU’s price from the `Products` Table.',
  platform: 'mac',
  formulas: ['=XLOOKUP(A2,Products[SKU],Products[Price])'],
  checks: [
    { label: 'Every answer cell has a formula', status: 'pass' },
    { label: 'Still correct when the Products Table is sorted', status: 'fail', detail: '`F7` showed 4.1 instead of 3.85.' },
    { label: 'Uses the Table’s column names', status: 'skip' },
  ],
  hintsShown: ['XLOOKUP takes the value to find, where to look, and what to return.'],
  ...over,
});

/** A Messages API reply as fetch would return it. */
const reply = (body: unknown, status = 200) => new Response(JSON.stringify(body), { status, headers: { 'content-type': 'application/json' } });
const textReply = (text: string, stop = 'end_turn') => reply({ type: 'message', role: 'assistant', content: [{ type: 'text', text }], stop_reason: stop });

const post = (body: unknown, over: Partial<HintHttpRequest> = {}): HintHttpRequest => ({
  method: 'POST',
  url: '/api/hint',
  remoteAddress: '127.0.0.1',
  headers: { host: 'localhost:3000', origin: 'https://localhost:3000', 'content-type': 'application/json' },
  body: typeof body === 'string' ? body : JSON.stringify(body),
  ...over,
});

const status = (over: Partial<HintHttpRequest> = {}): HintHttpRequest => ({ method: 'GET', url: '/api/hint/status', remoteAddress: '127.0.0.1', headers: { host: 'localhost:3000' }, ...over });

function stubFetch(respond: () => Response | Promise<Response>) {
  return vi.fn<typeof fetch>(async () => respond());
}

// ---------- prompt ----------

describe('hint prompt', () => {
  it('asks for at most three sentences about the learner’s own formula, without the answer', () => {
    expect(SYSTEM_PROMPT).toMatch(/at most 3 short sentences/);
    expect(SYSTEM_PROMPT).toMatch(/their formula/);
    expect(SYSTEM_PROMPT).toMatch(/Don't write the full answer formula/);
    expect(SYSTEM_PROMPT).toMatch(/backticks/);
    expect(SYSTEM_PROMPT).toMatch(/no headings/);
    expect(SYSTEM_PROMPT).toMatch(/please, simply, just, easy/);
    expect(SYSTEM_PROMPT).toMatch(/exclamation marks/);
    expect(SYSTEM_PROMPT).toMatch(/keyboard shortcuts for the learner's platform/);
  });

  it('lays the payload out in tagged sections', () => {
    const msg = buildUserMessage(payload());
    expect(msg).toContain('<exercise>Look up prices with XLOOKUP</exercise>');
    expect(msg).toContain('<platform>Excel for Mac</platform>');
    expect(msg).toContain('<task>In `F2:F31`');
    expect(msg).toContain('1. =XLOOKUP(A2,Products[SKU],Products[Price])');
    expect(msg).toContain('FAILED: Still correct when the Products Table is sorted. `F7` showed 4.1 instead of 3.85.');
    expect(msg).toContain('PASSED: Every answer cell has a formula');
    expect(msg).toContain('NOT CHECKED: Uses the Table’s column names');
    expect(msg).toContain('<hints_already_shown>\n1. XLOOKUP takes the value');
  });

  it('names the learner’s Excel', () => {
    expect(buildUserMessage(payload({ platform: 'windows' }))).toContain('<platform>Excel for Windows</platform>');
    expect(buildUserMessage(payload({ platform: 'web' }))).toContain('<platform>Excel for the web</platform>');
  });

  it('says so when there are no formulas or hints', () => {
    const msg = buildUserMessage(payload({ formulas: [], hintsShown: [] }));
    expect(msg).toMatch(/<formulas>\n\(none: /);
    expect(msg).toContain('<hints_already_shown>\n(none)\n</hints_already_shown>');
  });

  it('only shows details for failed checks', () => {
    const msg = buildUserMessage(payload({ checks: [{ label: 'A', status: 'pass', detail: 'ignored' }, { label: 'B', status: 'fail', detail: 'shown' }] }));
    expect(msg).not.toContain('ignored');
    expect(msg).toContain('FAILED: B. shown');
  });

  it('builds a small, thinking-off request for Sonnet 5.5', () => {
    const req = buildApiRequest(payload());
    expect(req.model).toBe('claude-sonnet-5-5');
    expect(HINT_MODEL).toBe('claude-sonnet-5-5');
    expect(req.max_tokens).toBe(HINT_MAX_TOKENS);
    expect(req.max_tokens).toBeLessThanOrEqual(300);
    expect(req.thinking).toEqual({ type: 'between_tools' });
    expect(req.output_config).toEqual({ effort: 'low' });
    expect(req.system).toBe(SYSTEM_PROMPT);
    expect(req.messages).toEqual([{ role: 'user', content: buildUserMessage(payload()) }]);
    // Sampling parameters are a 400 on this model.
    expect(req).not.toHaveProperty('temperature');
  });
});

// ---------- validation and limits ----------

describe('hint payload', () => {
  it('accepts a well-formed payload', () => {
    const r = parseHintPayload(payload());
    expect(r.ok).toBe(true);
    if (r.ok) expect(r.payload).toEqual(payload());
  });

  it.each([
    ['not an object', 'hello'],
    ['an array', [payload()]],
    ['no title', { ...payload(), title: '' }],
    ['a numeric task', { ...payload(), task: 4 }],
    ['no checks', { ...payload(), checks: undefined }],
    ['no platform', { ...payload(), platform: undefined }],
    ['an unknown platform', { ...payload(), platform: 'linux' }],
  ])('rejects %s', (_, body) => {
    expect(parseHintPayload(body)).toEqual({ ok: false, error: MESSAGES.invalid });
  });

  it('needs a failed check to explain', () => {
    const body = payload({ checks: [{ label: 'Every answer cell has a formula', status: 'pass' }] });
    expect(parseHintPayload(body)).toEqual({ ok: false, error: MESSAGES.nothingFailed });
  });

  it('drops malformed checks and unknown fields', () => {
    const body = {
      ...payload(),
      values: [[1, 2, 3]],
      apiKey: 'nope',
      checks: [...payload().checks, { label: 'Bad status', status: 'maybe' }, { status: 'fail' }, 'text', null],
      formulas: ['=A1', 42, null],
    };
    const r = parseHintPayload(body);
    expect(r.ok).toBe(true);
    if (!r.ok) return;
    expect(Object.keys(r.payload).sort()).toEqual(['checks', 'formulas', 'hintsShown', 'platform', 'task', 'title']);
    expect(r.payload.checks).toHaveLength(3);
    expect(r.payload.formulas).toEqual(['=A1']);
    expect(buildUserMessage(r.payload)).not.toContain('nope');
  });

  it('keeps at most 10 formulas of 500 characters', () => {
    const long = `=${'A1+'.repeat(400)}1`;
    const r = parseHintPayload(payload({ formulas: Array.from({ length: 14 }, (_, i) => (i === 0 ? long : `=B${i}`)) }));
    expect(r.ok).toBe(true);
    if (!r.ok) return;
    expect(r.payload.formulas).toHaveLength(HINT_LIMITS.formulas);
    expect(r.payload.formulas[0]).toHaveLength(HINT_LIMITS.formula);
    expect(r.payload.formulas[0].endsWith('…')).toBe(true);
    expect(r.payload.formulas[1]).toBe('=B1');
  });

  it('caps text fields and check details', () => {
    const r = parseHintPayload(
      payload({
        title: 'T'.repeat(500),
        task: 'x'.repeat(5000),
        checks: [{ label: 'L'.repeat(900), status: 'fail', detail: 'd'.repeat(2000) }],
        hintsShown: Array.from({ length: 9 }, (_, i) => `Hint ${i + 1} ${'h'.repeat(900)}`),
      }),
    );
    expect(r.ok).toBe(true);
    if (!r.ok) return;
    expect(r.payload.title).toHaveLength(HINT_LIMITS.title);
    expect(r.payload.task).toHaveLength(HINT_LIMITS.task);
    expect(r.payload.checks[0].label).toHaveLength(HINT_LIMITS.label);
    expect(r.payload.checks[0].detail).toHaveLength(HINT_LIMITS.detail);
    expect(r.payload.hintsShown).toHaveLength(HINT_LIMITS.hints);
    // The latest hints are the most specific, so they're the ones kept.
    expect(r.payload.hintsShown[0]).toMatch(/^Hint 5 /);
    expect(r.payload.hintsShown.every((h) => h.length === HINT_LIMITS.hint)).toBe(true);
  });

  it('keeps failed checks when there are too many', () => {
    const checks = Array.from({ length: 30 }, (_, i) => ({ label: `Check ${i}`, status: i === 25 ? ('fail' as const) : ('pass' as const) }));
    const clamped = clampHintPayload(payload({ checks }));
    expect(clamped.checks).toHaveLength(HINT_LIMITS.checks);
    expect(clamped.checks.some((c) => c.label === 'Check 25' && c.status === 'fail')).toBe(true);
    // Original order is kept.
    expect(clamped.checks.map((c) => c.label)).toEqual([...clamped.checks.map((c) => c.label)].sort((a, b) => Number(a.slice(6)) - Number(b.slice(6))));
  });

  it('fits the largest clamped payload under the body limit', () => {
    const big = clampHintPayload({
      title: 'T'.repeat(1000),
      task: 'x'.repeat(10_000),
      platform: 'windows',
      formulas: Array.from({ length: 50 }, () => '='.padEnd(2000, 'A')),
      checks: Array.from({ length: 50 }, () => ({ label: 'L'.repeat(1000), status: 'fail' as const, detail: 'd'.repeat(1000) })),
      hintsShown: Array.from({ length: 50 }, () => 'h'.repeat(1000)),
    });
    expect(Buffer.byteLength(JSON.stringify(big))).toBeLessThan(HINT_BODY_LIMIT);
  });
});

describe('formula picking', () => {
  // A 30-row answer in F2:F31, filled down, with the rate typed over in F28.
  const filled = Array.from({ length: 30 }, (_, i) => (i === 26 ? '=C28*D28*1.0825' : `=C${i + 2}*D${i + 2}*$H$1`));

  it('gives every copy of a filled formula the same shape', () => {
    expect(formulaShape('=C2*D2*$H$1')).toBe(formulaShape('=C30*D30*$H$1'));
    expect(formulaShape('=SUM($B$2:B2)')).toBe(formulaShape('=SUM($B$2:B9)'));
    // Filled across, the column letters move instead.
    expect(formulaShape('=B$1*$A2')).toBe(formulaShape('=C$1*$A2'));
  });

  it('tells apart a typed-in number and a reference that slipped a row', () => {
    expect(formulaShape('=C28*D28*1.0825')).not.toBe(formulaShape('=C28*D28*$H$1'));
    expect(formulaShape('=C27*D28*$H$1')).not.toBe(formulaShape('=C28*D28*$H$1'));
  });

  it('leaves text, quoted sheet names, functions and names alone', () => {
    expect(formulaShape('=IF(A2="B2","Q1",B2)')).toBe('=IF(C[0]R[0]="B2","Q1",C[1]R[0])');
    expect(formulaShape("='Q1 Data'!B2+LOG10(C2)")).toBe("='Q1 Data'!C[0]R[0]+LOG10(C[1]R[0])");
    expect(formulaShape('=Products[Price]*TaxRate')).toBe('=Products[Price]*TaxRate');
  });

  it('keeps the odd formula out of a long filled column', () => {
    const picked = pickFormulas(filled, HINT_LIMITS.formulas);
    expect(picked).toHaveLength(HINT_LIMITS.formulas);
    expect(picked).toContain('=C28*D28*1.0825');
    expect(picked[0]).toBe('=C2*D2*$H$1');
    // Reading order is kept.
    expect(picked).toEqual(filled.filter((f) => picked.includes(f)));
  });

  it('keeps one of each shape before any repeats, rarest first', () => {
    const twoColumns = Array.from({ length: 20 }, (_, i) => [`=B${i + 2}*2`, `=C${i + 2}+1`]).flat();
    const picked = pickFormulas([...twoColumns, '=SUM(B2:B21)', '=B22*3'], 3);
    expect(picked).toEqual(['=B2*2', '=SUM(B2:B21)', '=B22*3']);
  });

  it('changes nothing when the list fits', () => {
    expect(pickFormulas(['=B9', '=A1', '=B9+1'], 10)).toEqual(['=B9', '=A1', '=B9+1']);
  });

  it('applies to the payload on both sides', () => {
    const clamped = clampHintPayload(payload({ formulas: filled }));
    expect(clamped.formulas).toContain('=C28*D28*1.0825');
    const r = parseHintPayload(payload({ formulas: filled }));
    expect(r.ok && r.payload.formulas).toEqual(clamped.formulas);
  });
});

// ---------- the reply ----------

describe('tidyHint', () => {
  it('flattens markdown into one plain paragraph', () => {
    expect(tidyHint('## Why\n**Your** range stops at `B20`.\n- Extend it.')).toBe('Why Your range stops at `B20`. Extend it.');
  });

  it('drops exclamation marks and curls apostrophes outside code only', () => {
    expect(tidyHint("You're close! Check `='Q1 Data'!B2` and don't type 0.08!")).toBe('You’re close. Check `=\'Q1 Data\'!B2` and don’t type 0.08.');
  });

  it('turns runs of exclamation marks into one period, before closing punctuation too', () => {
    expect(tidyHint('It works!!')).toBe('It works.');
    expect(tidyHint('Great work!) Next, anchor `H1`.')).toBe('Great work.) Next, anchor `H1`.');
    expect(tidyHint('Use `=Sheet2!B2` here.')).toBe('Use `=Sheet2!B2` here.');
  });

  it('turns a fenced block into one code span, and drops empty spans', () => {
    expect(tidyHint('Compare it with this:\n```excel\n=SUM(B2:B10)\n```\nIt stops short.')).toBe('Compare it with this: `=SUM(B2:B10)` It stops short.');
    expect(tidyHint('Your formula ``` =SUM(B2:B10) ``` stops short.')).toBe('Your formula `=SUM(B2:B10)` stops short.');
    expect(tidyHint('```SUM``` stops short.')).toBe('`SUM` stops short.');
    expect(tidyHint('Your range `` stops at `B20`.')).toBe('Your range stops at `B20`.');
  });

  it('closes no code span that was never closed', () => {
    expect(tidyHint('Your range stops at `B20.')).toBe('Your range stops at B20.');
  });

  it('strips numbered lists, emphasis and links', () => {
    expect(tidyHint('1. Your range stops short.\n2) Extend it.')).toBe('Your range stops short. Extend it.');
    expect(tidyHint('*Your* range uses _relative_ references, so `C2 * D2` shifts.')).toBe('Your range uses relative references, so `C2 * D2` shifts.');
    expect(tidyHint('See [XLOOKUP](https://support.microsoft.com/xlookup) for the arguments.')).toBe('See XLOOKUP for the arguments.');
    // Not emphasis: multiplication and names with underscores.
    expect(tidyHint('The rate is B2 * C2 in Tax_Rate.')).toBe('The rate is B2 * C2 in Tax_Rate.');
  });

  it('drops filler words outside code and keeps sentence case', () => {
    expect(tidyHint('Just extend `B2:B10`, it simply stops short. You just need `$`. Please check `"just"`.')).toBe(
      'Extend `B2:B10`, it stops short. You need `$`. Check `"just"`.',
    );
    expect(tidyHint('Adjust the range.')).toBe('Adjust the range.');
  });

  it('trims a reply cut off by max_tokens back to its last full sentence', () => {
    expect(tidyHint('Your range stops short. Extend it to the last row of the', true)).toBe('Your range stops short.');
    expect(tidyHint('Your range stops short of the', true)).toBe('Your range stops short of the…');
  });
});

// ---------- loopback ----------

describe('loopback checks', () => {
  it.each(['127.0.0.1', '127.8.9.10', '::1', '::ffff:127.0.0.1'])('accepts %s', (a) => {
    expect(isLoopbackAddress(a)).toBe(true);
  });

  it.each([undefined, '', '192.168.1.20', '10.0.0.2', '::ffff:192.168.1.20', 'fe80::1', '128.0.0.1'])('rejects %s', (a) => {
    expect(isLoopbackAddress(a)).toBe(false);
  });

  it('checks the Host header', () => {
    expect(isLoopbackHost('localhost:3000')).toBe(true);
    expect(isLoopbackHost('127.0.0.1:3000')).toBe(true);
    expect(isLoopbackHost('[::1]:3000')).toBe(true);
    expect(isLoopbackHost('coach.localhost')).toBe(true);
    expect(isLoopbackHost('evil.example:3000')).toBe(false);
    expect(isLoopbackHost('localhost.evil.example')).toBe(false);
  });

  it('checks the Origin header', () => {
    expect(isLoopbackOrigin(undefined)).toBe(true);
    expect(isLoopbackOrigin('https://localhost:3000')).toBe(true);
    expect(isLoopbackOrigin('http://127.0.0.1:5173')).toBe(true);
    expect(isLoopbackOrigin('https://evil.example')).toBe(false);
    expect(isLoopbackOrigin('null')).toBe(false);
  });
});

describe('readBody', () => {
  it('reads a stream under the limit', async () => {
    expect(await readBody(Readable.from([Buffer.from('{"a":'), Buffer.from('1}')]), 100)).toBe('{"a":1}');
  });

  it('gives up past the limit but drains the stream', async () => {
    const chunks = Array.from({ length: 5 }, () => Buffer.alloc(10, 97));
    const stream = Readable.from(chunks);
    expect(await readBody(stream, 25)).toBeNull();
    expect(stream.readableEnded).toBe(true);
  });

  it('measures strings in bytes', async () => {
    expect(await readBody('’’’’', 8)).toBeNull();
    expect(await readBody('’’', 8)).toBe('’’');
  });
});

// ---------- the handler ----------

describe('handleHintRequest', () => {
  let log: string[];
  let deps: (fetchImpl?: typeof fetch) => HintDeps;

  beforeEach(() => {
    log = [];
    deps = (fetchImpl) => ({ apiKey: KEY, fetch: fetchImpl ?? stubFetch(() => textReply('Your range stops short.')), log: (m) => log.push(m) });
  });

  it('ignores other paths', async () => {
    expect(await handleHintRequest({ ...status(), url: '/src/main.tsx' }, deps())).toBeUndefined();
    expect(await handleHintRequest({ ...status(), url: '/api/hints' }, deps())).toBeUndefined();
  });

  it('reports enabled: false without a key', async () => {
    expect(await handleHintRequest(status(), { log: () => {} })).toEqual({ status: 200, body: { enabled: false } });
    expect(await handleHintRequest(status(), { apiKey: '   ' })).toEqual({ status: 200, body: { enabled: false } });
  });

  it('reports enabled: true with a key, and ignores a query string', async () => {
    expect(await handleHintRequest(status({ url: '/api/hint/status?t=1' }), deps())).toEqual({ status: 200, body: { enabled: true } });
  });

  it('refuses to explain without a key, before calling Claude', async () => {
    const f = stubFetch(() => textReply('x'));
    const res = await handleHintRequest(post(payload()), { fetch: f });
    expect(res).toEqual({ status: 503, body: { error: MESSAGES.off } });
    expect(f).not.toHaveBeenCalled();
  });

  it.each([
    ['a LAN address', { remoteAddress: '192.168.1.20' }],
    ['a missing address', { remoteAddress: undefined }],
    ['a foreign Host (DNS rebinding)', { headers: { host: 'evil.example:3000', 'content-type': 'application/json' } }],
    ['a foreign HTTP/2 authority', { headers: { ':authority': 'evil.example:3000', 'content-type': 'application/json' } }],
    ['a foreign Origin', { headers: { host: 'localhost:3000', origin: 'https://evil.example', 'content-type': 'application/json' } }],
  ])('rejects %s', async (_, over) => {
    const f = stubFetch(() => textReply('x'));
    const res = await handleHintRequest(post(payload(), over), deps(f));
    expect(res).toEqual({ status: 403, body: { error: MESSAGES.notLocal } });
    expect(f).not.toHaveBeenCalled();
    const st = await handleHintRequest(status(over), deps(f));
    expect(st?.status).toBe(403);
  });

  it('only allows GET on status and POST on the hint route', async () => {
    expect(await handleHintRequest(status({ method: 'POST' }), deps())).toMatchObject({ status: 405, headers: { Allow: 'GET' } });
    expect(await handleHintRequest(post(payload(), { method: 'GET' }), deps())).toMatchObject({ status: 405, headers: { Allow: 'POST' } });
  });

  it('requires JSON', async () => {
    const res = await handleHintRequest(post(payload(), { headers: { host: 'localhost:3000', 'content-type': 'text/plain' } }), deps());
    expect(res).toEqual({ status: 415, body: { error: MESSAGES.notJson } });
  });

  it('rejects oversized bodies by Content-Length, as a string, and as a stream', async () => {
    const f = stubFetch(() => textReply('x'));
    const big = JSON.stringify(payload({ task: 'x'.repeat(HINT_BODY_LIMIT) }));
    const headers = { host: 'localhost:3000', 'content-type': 'application/json' };
    const byLength = await handleHintRequest(post('{}', { headers: { ...headers, 'content-length': String(HINT_BODY_LIMIT + 1) } }), deps(f));
    const byString = await handleHintRequest(post(big), deps(f));
    const byStream = await handleHintRequest(post('', { body: Readable.from([Buffer.from(big)]) }), deps(f));
    for (const res of [byLength, byString, byStream]) expect(res).toEqual({ status: 413, body: { error: MESSAGES.tooLarge } });
    expect(f).not.toHaveBeenCalled();
  });

  it('rejects bad JSON and incomplete payloads', async () => {
    expect(await handleHintRequest(post('{"title":'), deps())).toEqual({ status: 400, body: { error: MESSAGES.invalid } });
    expect(await handleHintRequest(post({ title: 'x' }), deps())).toEqual({ status: 400, body: { error: MESSAGES.invalid } });
  });

  it('calls the Messages API with the key in the header only', async () => {
    const f = stubFetch(() => textReply('Your lookup range is fixed to `A2:A20`, so sorted rows fall outside it. Point it at the whole Table column instead.'));
    const res = await handleHintRequest(post(payload()), deps(f));
    expect(res).toEqual({ status: 200, body: { hint: 'Your lookup range is fixed to `A2:A20`, so sorted rows fall outside it. Point it at the whole Table column instead.' } });
    expect(f).toHaveBeenCalledTimes(1);
    const [url, init] = f.mock.calls[0];
    expect(url).toBe('https://api.anthropic.com/v1/messages');
    expect(init?.method).toBe('POST');
    const headers = init?.headers as Record<string, string>;
    expect(headers['x-api-key']).toBe(KEY);
    expect(headers['anthropic-version']).toBe('2023-06-01');
    expect(headers['content-type']).toBe('application/json');
    // fallbacks: 'default' is a 400 without its beta header.
    expect(headers['anthropic-beta']).toBe('server-side-fallback-2026-07-01');
    expect(buildApiRequest(payload()).fallbacks).toBe('default');
    expect(init?.signal).toBeInstanceOf(AbortSignal);
    const body = JSON.parse(String(init?.body));
    expect(body).toEqual(buildApiRequest(payload()));
    expect(String(init?.body)).not.toContain(KEY);
  });

  it('accepts a streamed request body', async () => {
    const res = await handleHintRequest(post('', { body: Readable.from([Buffer.from(JSON.stringify(payload()))]) }), deps());
    expect(res).toEqual({ status: 200, body: { hint: 'Your range stops short.' } });
  });

  it('reads only text blocks from the reply', async () => {
    const f = stubFetch(() =>
      reply({
        content: [
          { type: 'thinking', thinking: '', signature: 'abc' },
          { type: 'text', text: 'Your range stops short.' },
        ],
        stop_reason: 'end_turn',
      }),
    );
    expect(await handleHintRequest(post(payload()), deps(f))).toEqual({ status: 200, body: { hint: 'Your range stops short.' } });
  });

  it('reads the text after a server-side fallback', async () => {
    const f = stubFetch(() =>
      reply({
        model: 'claude-sonnet-5',
        content: [
          { type: 'fallback', from: { model: 'claude-sonnet-5-5' }, to: { model: 'claude-sonnet-5' } },
          { type: 'text', text: 'Your range stops short.' },
        ],
        stop_reason: 'end_turn',
      }),
    );
    expect(await handleHintRequest(post(payload()), deps(f))).toEqual({ status: 200, body: { hint: 'Your range stops short.' } });
  });

  it.each([
    [400, 'invalid_request_error', MESSAGES.rejected],
    [401, 'authentication_error', MESSAGES.badKey],
    [402, 'billing_error', MESSAGES.noCredit],
    [403, 'permission_error', MESSAGES.noAccess],
    [404, 'not_found_error', MESSAGES.noAccess],
    [413, 'request_too_large', MESSAGES.rejected],
    [429, 'rate_limit_error', MESSAGES.rateLimited],
    [500, 'api_error', MESSAGES.busy],
    [529, 'overloaded_error', MESSAGES.busy],
  ])('maps API status %i to a friendly message', async (code, type, message) => {
    const f = stubFetch(() => reply({ type: 'error', error: { type, message: `upstream says ${KEY}` }, request_id: 'req_123' }, code));
    const res = await handleHintRequest(post(payload()), deps(f));
    expect(res).toEqual({ status: 502, body: { error: message } });
    expect(log).toEqual([`Claude API ${code} ${type} (request req_123): upstream says [key]`]);
  });

  it('logs which field the API rejected, on one short line', async () => {
    const why = `fallbacks:\n  Extra inputs are not permitted ${'x'.repeat(400)}`;
    const f = stubFetch(() => reply({ type: 'error', error: { type: 'invalid_request_error', message: why }, request_id: 'req_9' }, 400));
    await handleHintRequest(post(payload()), deps(f));
    expect(log[0]).toMatch(/^Claude API 400 invalid_request_error \(request req_9\): fallbacks: Extra inputs are not permitted x+$/);
    expect(log[0]).not.toContain('\n');
    expect(log[0].length).toBeLessThan(300);
  });

  it('handles an API error with an unreadable body', async () => {
    const f = stubFetch(() => new Response('<html>Bad gateway</html>', { status: 502 }));
    expect(await handleHintRequest(post(payload()), deps(f))).toEqual({ status: 502, body: { error: MESSAGES.busy } });
  });

  it('handles a refusal', async () => {
    const f = stubFetch(() => reply({ content: [], stop_reason: 'refusal', stop_details: { type: 'refusal', category: 'general_harms' } }));
    expect(await handleHintRequest(post(payload()), deps(f))).toEqual({ status: 502, body: { error: MESSAGES.declined } });
  });

  it('handles an empty reply', async () => {
    const f = stubFetch(() => reply({ content: [], stop_reason: 'end_turn' }));
    expect(await handleHintRequest(post(payload()), deps(f))).toEqual({ status: 502, body: { error: MESSAGES.empty } });
  });

  it('trims a reply that hit max_tokens', async () => {
    const f = stubFetch(() => textReply('Your range stops short. Extend it down to the', 'max_tokens'));
    expect(await handleHintRequest(post(payload()), deps(f))).toEqual({ status: 200, body: { hint: 'Your range stops short.' } });
  });

  it('handles a network failure', async () => {
    const f = vi.fn<typeof fetch>(async () => {
      throw new TypeError('fetch failed');
    });
    expect(await handleHintRequest(post(payload()), deps(f))).toEqual({ status: 502, body: { error: MESSAGES.offline } });
  });

  it('times out', async () => {
    // Waits until the request's own signal gives up, like a hung connection.
    const hang = vi.fn<typeof fetch>((_url, init) => new Promise((_, reject) => init?.signal?.addEventListener('abort', () => reject(init.signal!.reason))));
    const res = await explainMistake(payload(), { apiKey: KEY, fetch: hang, timeoutMs: 20, log: (m) => log.push(m) });
    expect(res).toEqual({ status: 504, body: { error: MESSAGES.timeout } });
    expect(log).toEqual(['Claude API timed out']);
  });

  it('never puts the key in a response or a log line', async () => {
    const replies: (() => Response | Promise<Response>)[] = [
      () => textReply(`Here is ${KEY.slice(0, 6)} and more.`),
      () => reply({ error: { type: 'authentication_error', message: `invalid x-api-key ${KEY}` } }, 401),
      () => reply({ error: { type: 'api_error', message: KEY } }, 500),
      () => reply({ content: [], stop_reason: 'refusal' }),
      () => {
        throw new Error(`connect failed for ${KEY}`);
      },
    ];
    const requests = [post(payload()), status(), post('{'), post(payload(), { remoteAddress: '10.0.0.5' }), post(payload(), { method: 'PUT' })];
    for (const respond of replies) {
      for (const req of requests) {
        const res = await handleHintRequest(req, deps(stubFetch(respond)));
        expect(JSON.stringify(res)).not.toContain(KEY);
      }
    }
    expect(log.length).toBeGreaterThan(0);
    expect(log.join('\n')).not.toContain(KEY);
  });
});

// ---------- copy ----------

describe('hint copy', () => {
  it('follows the copy rules in every message the panel can show', async () => {
    const { CLIENT_MESSAGES } = await import('../src/ai/hints');
    for (const text of [...Object.values(MESSAGES), ...Object.values(CLIENT_MESSAGES)]) {
      expect(text, text).not.toMatch(BANNED);
      expect(text, text).not.toContain("'");
      expect(text, text).toMatch(/^[A-Z]/);
      expect(text, text).toMatch(/\.$/);
    }
  });
});

// ---------- the panel's client ----------

describe('hint client', () => {
  const report: CheckReport = {
    passed: false,
    items: [
      { id: 'formula', label: 'Every answer cell has a formula', status: 'pass' },
      { id: 'variant-0', label: 'Still correct when rows are added', status: 'fail', detail: '`F32` is blank.', focus: 'F32' },
    ],
    marks: [{ address: 'F2', ok: true }],
    focus: 'F32',
    formulas: ['=SUM(B2:B31)'],
  };

  beforeEach(() => {
    vi.resetModules();
  });

  afterEach(() => {
    vi.unstubAllGlobals();
    vi.unstubAllEnvs();
  });

  it('builds a payload from a check report without ids, marks or focus cells', async () => {
    const { buildHintPayload } = await import('../src/ai/hints');
    const p = buildHintPayload({ title: 'Totals', task: 'Total `B2:B31`.', platform: 'mac', report, hintsShown: ['Use SUM.'] });
    expect(p).toEqual({
      title: 'Totals',
      task: 'Total `B2:B31`.',
      platform: 'mac',
      formulas: ['=SUM(B2:B31)'],
      checks: [
        { label: 'Every answer cell has a formula', status: 'pass' },
        { label: 'Still correct when rows are added', status: 'fail', detail: '`F32` is blank.' },
      ],
      hintsShown: ['Use SUM.'],
    });
    expect(buildHintPayload({ title: 'T', task: 't', platform: 'windows', report: { ...report, formulas: undefined }, hintsShown: [] }).formulas).toEqual([]);
  });

  it('offers an explanation only when a check failed', async () => {
    const { canExplain } = await import('../src/ai/hints');
    expect(canExplain(report)).toBe(true);
    expect(canExplain({ ...report, passed: true, items: report.items.filter((i) => i.status === 'pass') })).toBe(false);
    // An empty report doesn't pass, but there's nothing in it to explain.
    expect(canExplain({ passed: false, items: [], marks: [] })).toBe(false);
  });

  it('reports hints on when the server says so, and asks only once', async () => {
    const f = vi.fn<typeof fetch>(async () => reply({ enabled: true }));
    vi.stubGlobal('fetch', f);
    const { hintStatus, knownHintStatus } = await import('../src/ai/hints');
    expect(knownHintStatus()).toBeUndefined();
    expect(await hintStatus()).toBe(true);
    expect(await hintStatus()).toBe(true);
    expect(knownHintStatus()).toBe(true);
    expect(f).toHaveBeenCalledTimes(1);
    expect(f.mock.calls[0][0]).toBe('/api/hint/status');
  });

  it('reports hints off in a static build without asking a server', async () => {
    vi.stubEnv('DEV', false);
    const f = vi.fn<typeof fetch>(async () => reply({ enabled: true }));
    vi.stubGlobal('fetch', f);
    const { hintStatus, knownHintStatus } = await import('../src/ai/hints');
    expect(await hintStatus()).toBe(false);
    expect(knownHintStatus()).toBe(false);
    expect(f).not.toHaveBeenCalled();
  });

  it.each([
    ['the server says off', () => reply({ enabled: false })],
    ['there is no server (static hosting returns a 404 page)', () => new Response('<html>Not found</html>', { status: 404 })],
    ['the answer is not JSON', () => new Response('<html></html>', { status: 200 })],
    ['the network fails', () => Promise.reject(new TypeError('Failed to fetch'))],
  ])('reports hints off when %s', async (_, respond) => {
    vi.stubGlobal('fetch', vi.fn(respond));
    const { hintStatus } = await import('../src/ai/hints');
    expect(await hintStatus()).toBe(false);
  });

  it('posts the payload and returns the hint', async () => {
    const f = vi.fn<typeof fetch>(async () => reply({ hint: 'Your range stops at `B31`.' }));
    vi.stubGlobal('fetch', f);
    const { requestHint } = await import('../src/ai/hints');
    expect(await requestHint(payload())).toBe('Your range stops at `B31`.');
    const [url, init] = f.mock.calls[0];
    expect(url).toBe('/api/hint');
    expect(init?.method).toBe('POST');
    expect((init?.headers as Record<string, string>)['Content-Type']).toBe('application/json');
    expect(JSON.parse(String(init?.body))).toEqual(payload());
  });

  it('shows the server’s message when the request fails', async () => {
    vi.stubGlobal('fetch', vi.fn(async () => reply({ error: MESSAGES.rateLimited }, 502)));
    const { requestHint, HintError } = await import('../src/ai/hints');
    const err = await requestHint(payload()).catch((e: unknown) => e);
    expect(err).toBeInstanceOf(HintError);
    expect((err as Error).message).toBe(MESSAGES.rateLimited);
  });

  it('explains when the hint server can’t be reached', async () => {
    vi.stubGlobal('fetch', vi.fn(async () => Promise.reject(new TypeError('Failed to fetch'))));
    const { requestHint, CLIENT_MESSAGES } = await import('../src/ai/hints');
    await expect(requestHint(payload())).rejects.toThrow(CLIENT_MESSAGES.unreachable);
  });

  it('falls back to a generic message for an unexpected reply', async () => {
    vi.stubGlobal('fetch', vi.fn(async () => reply({ something: 'else' })));
    const { requestHint, CLIENT_MESSAGES } = await import('../src/ai/hints');
    await expect(requestHint(payload())).rejects.toThrow(CLIENT_MESSAGES.failed);
  });

  it('stops when the caller aborts', async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn((_url: string, init: RequestInit) => new Promise((_, reject) => init.signal?.addEventListener('abort', () => reject(new DOMException('Aborted', 'AbortError'))))),
    );
    const { requestHint, HintError } = await import('../src/ai/hints');
    const ctrl = new AbortController();
    const pending = requestHint(payload(), ctrl.signal).catch((e: unknown) => e);
    ctrl.abort();
    const err = await pending;
    expect(err).not.toBeInstanceOf(HintError);
  });
});

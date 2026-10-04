import { describe, expect, it } from 'vitest';
import localManifest from '../manifest.xml?raw';
import {
  HOSTED_ADDIN_ID,
  HOSTED_DISPLAY_NAME,
  LOCAL_ORIGIN,
  basePathOf,
  builtSiteProblems,
  hostedManifest,
  localUrls,
  normalizePagesUrl,
} from '../scripts/pagesManifest.mjs';

const PAGES = 'https://owner.github.io/excel-coach/';
const hosted = hostedManifest(localManifest, PAGES);

const escapeRe = (text: string) => text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
const idOf = (xml: string) => xml.match(/<Id>([^<]*)<\/Id>/)?.[1];
const attr = (xml: string, element: string) => xml.match(new RegExp(`<${element}\\b[^>]*?DefaultValue="([^"]*)"`))?.[1];
const resource = (xml: string, id: string) => xml.match(new RegExp(`id="${escapeRe(id)}" DefaultValue="([^"]*)"`))?.[1];

describe('PAGES_URL', () => {
  it('ends in a slash and lowercases the host', () => {
    expect(normalizePagesUrl('https://owner.github.io/excel-coach/')).toBe(PAGES);
    expect(normalizePagesUrl('  https://Owner.GitHub.io/excel-coach  ')).toBe(PAGES);
    expect(normalizePagesUrl('https://owner.github.io')).toBe('https://owner.github.io/');
    expect(normalizePagesUrl('https://owner.github.io/excel.coach/')).toBe('https://owner.github.io/excel.coach/');
  });
  it('rejects addresses Excel can’t load the panel from', () => {
    expect(() => normalizePagesUrl(undefined)).toThrow(/Set PAGES_URL/);
    expect(() => normalizePagesUrl('')).toThrow(/Set PAGES_URL/);
    expect(() => normalizePagesUrl('owner.github.io/excel-coach')).toThrow(/isn’t a web address/);
    expect(() => normalizePagesUrl('http://owner.github.io/excel-coach/')).toThrow(/https/);
    expect(() => normalizePagesUrl('https://localhost:3000/')).toThrow(/this computer/);
    expect(() => normalizePagesUrl('https://127.0.0.1/excel-coach/')).toThrow(/this computer/);
    expect(() => normalizePagesUrl('https://coach.localhost/')).toThrow(/this computer/);
    expect(() => normalizePagesUrl('https://owner.github.io/excel-coach/?v=2')).toThrow(/query/);
    expect(() => normalizePagesUrl('https://owner.github.io/excel-coach/#top')).toThrow(/fragment/);
    // URL keeps an empty query or fragment in href, which would end up in every manifest URL.
    expect(() => normalizePagesUrl('https://owner.github.io/excel-coach/?')).toThrow(/query/);
    expect(() => normalizePagesUrl('https://owner.github.io/excel-coach/#')).toThrow(/fragment/);
    expect(() => normalizePagesUrl('https://owner.github.io/excel-coach/taskpane.html')).toThrow(/not a file/);
    expect(() => normalizePagesUrl('https://owner.github.io/excel-coach/manifest.xml')).toThrow(/Drop the file name/);
    expect(() => normalizePagesUrl('https://owner.github.io/excel.coach')).toThrow(/end a folder’s address with \//);
  });
});

describe('base path', () => {
  it('is the repository path for a project site', () => {
    expect(basePathOf(PAGES)).toBe('/excel-coach/');
    expect(basePathOf('https://owner.github.io/excel-coach')).toBe('/excel-coach/');
    expect(basePathOf('https://owner.github.io/tools/excel-coach/')).toBe('/tools/excel-coach/');
  });
  it('is the root for a user site or a custom domain', () => {
    expect(basePathOf('https://owner.github.io/')).toBe('/');
    expect(basePathOf('https://coach.example.com')).toBe('/');
  });
});

describe('hosted manifest', () => {
  it('starts from a local manifest that loads from the dev server', () => {
    expect(localManifest).toContain(`<SourceLocation DefaultValue="${LOCAL_ORIGIN}/taskpane.html"/>`);
    expect(localUrls(localManifest).length).toBeGreaterThan(5);
  });

  it('leaves no address on this computer', () => {
    expect(hosted).not.toMatch(/localhost|127\.0\.0\.1/i);
    expect(localUrls(hosted)).toEqual([]);
  });

  it('moves every local URL to the same path under PAGES_URL', () => {
    const local = [...localManifest.matchAll(new RegExp(`${escapeRe(LOCAL_ORIGIN)}/([^"<]*)`, 'g'))].map((m) => m[1]);
    const moved = [...hosted.matchAll(new RegExp(`${escapeRe(PAGES)}([^"<]*)`, 'g'))].map((m) => m[1]);
    expect(moved).toEqual(local);
    expect(local).toContain('taskpane.html');
    expect(local).toContain('commands.html');
    expect(local).toContain('assets/icon-80.png');
  });

  it('points each kind of URL at Pages', () => {
    expect(attr(hosted, 'SourceLocation')).toBe(`${PAGES}taskpane.html`);
    expect(attr(hosted, 'IconUrl')).toBe(`${PAGES}assets/icon-32.png`);
    expect(attr(hosted, 'HighResolutionIconUrl')).toBe(`${PAGES}assets/icon-64.png`);
    expect(attr(hosted, 'SupportUrl')).toBe(`${PAGES}taskpane.html`);
    expect(resource(hosted, 'Taskpane.Url')).toBe(`${PAGES}taskpane.html`);
    expect(resource(hosted, 'Commands.Url')).toBe(`${PAGES}commands.html`);
    expect(resource(hosted, 'Icon.16x16')).toBe(`${PAGES}assets/icon-16.png`);
  });

  it('lists the Pages origin, with no path, as the app domain', () => {
    expect(hosted).toContain('<AppDomain>https://owner.github.io</AppDomain>');
  });

  it('uses its own add-in Id and display name', () => {
    const guid = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/;
    expect(idOf(localManifest)).toMatch(guid);
    expect(HOSTED_ADDIN_ID).toMatch(guid);
    expect(idOf(hosted)).toBe(HOSTED_ADDIN_ID);
    expect(idOf(hosted)).not.toBe(idOf(localManifest));
    expect(attr(hosted, 'DisplayName')).toBe(HOSTED_DISPLAY_NAME);
    expect(HOSTED_DISPLAY_NAME).toBe('Excel Coach (web)');
  });

  it('labels its ribbon button differently from the local add-in', () => {
    // The label is also the tooltip's title. With both add-ins installed, it tells the buttons apart.
    expect(resource(localManifest, 'CoachButton.Label')).toBe('Excel Coach');
    expect(resource(hosted, 'CoachButton.Label')).toBe(HOSTED_DISPLAY_NAME);
  });

  it('keeps the same Id on every build and for every address', () => {
    // Installed copies are tied to this Id. If it changes, Excel treats the next build as a
    // different add-in, and every installed copy has to be removed and installed again.
    expect(HOSTED_ADDIN_ID).toBe('c16d946b-e5c7-4067-bc7a-5340a1bd6aa5');
    expect(hostedManifest(localManifest, PAGES)).toBe(hosted);
    expect(idOf(hostedManifest(localManifest, 'https://coach.example.com/'))).toBe(HOSTED_ADDIN_ID);
  });

  it('changes nothing else', () => {
    const before = localManifest.split('\n');
    const after = hosted.split('\n');
    expect(after).toHaveLength(before.length);
    const changed = before.filter((line, i) => line !== after[i]);
    expect(changed.length).toBeGreaterThan(0);
    expect(changed.every((line) => line.includes(LOCAL_ORIGIN) || /<Id>|<DisplayName |id="CoachButton\.Label"/.test(line))).toBe(true);
  });

  it('escapes the address for XML', () => {
    expect(hostedManifest(localManifest, 'https://coach.example.com/a&b/')).toContain('"https://coach.example.com/a&amp;b/taskpane.html"');
  });
});

describe('fails loudly', () => {
  const withLine = (line: string) => localManifest.replace('</DefaultSettings>', `  ${line}\n  </DefaultSettings>`);
  const lineOf = (xml: string, text: string) => xml.slice(0, xml.indexOf(text)).split('\n').length;

  it('on a local URL it doesn’t know how to move', () => {
    for (const url of ['https://localhost:3001/taskpane.html', 'http://localhost:3000/taskpane.html', 'https://127.0.0.1:3000/x', 'https://[::1]:3000/x']) {
      const source = withLine(`<SourceLocation DefaultValue="${url}"/>`);
      expect(() => hostedManifest(source, PAGES)).toThrow(`line ${lineOf(source, url)}: ${url}`);
    }
  });

  it('on a lookalike of the dev server address', () => {
    const source = withLine('<SourceLocation DefaultValue="https://localhost:30001/taskpane.html"/>');
    expect(() => hostedManifest(source, PAGES)).toThrow('https://localhost:30001/taskpane.html');
  });

  it('when the Id, display name or button label is missing or repeated', () => {
    expect(() => hostedManifest(localManifest.replace(/<Id>[^<]*<\/Id>/, ''), PAGES)).toThrow('exactly one <Id>, but has 0');
    expect(() => hostedManifest(localManifest.replace('</Hosts>', '</Hosts><Id>x</Id>'), PAGES)).toThrow('exactly one <Id>, but has 2');
    expect(() => hostedManifest(localManifest.replace(/<DisplayName [^>]*\/>/, ''), PAGES)).toThrow('exactly one <DisplayName>');
    expect(() => hostedManifest(localManifest.replace(/<bt:String id="CoachButton\.Label"[^>]*\/>/, ''), PAGES)).toThrow(
      'exactly one CoachButton.Label string, but has 0',
    );
  });

  it('on a bad PAGES_URL', () => {
    expect(() => hostedManifest(localManifest, '')).toThrow(/Set PAGES_URL/);
    expect(() => hostedManifest(localManifest, 'https://localhost:3000/')).toThrow(/this computer/);
  });
});

describe('built site', () => {
  const BASE = '/excel-coach/';
  const page = (src: string) => `<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
    <script type="module" crossorigin src="${src}"></script>`;
  const site = (files: Record<string, string> = {}) => ({
    'taskpane.html': page(`${BASE}assets/taskpane-abc.js`),
    'commands.html': '<script>Office.onReady(function () {});</script>',
    'assets/taskpane-abc.js': 'fetch("/api/hint/status").then((r) => r.ok);',
    ...files,
  });

  it('passes a site built under the base path', () => {
    expect(builtSiteProblems(site(), BASE)).toEqual([]);
    expect(builtSiteProblems(site({ 'taskpane.html': page('/assets/taskpane-abc.js') }), '/')).toEqual([]);
  });

  it('fails when the page doesn’t load its scripts from the base path', () => {
    const problems = builtSiteProblems(site({ 'taskpane.html': page('/assets/taskpane-abc.js') }), BASE);
    expect(problems).toContain(
      'dist/taskpane.html doesn’t load its scripts from /excel-coach/assets/. Check that vite.config.ts still reads COACH_BASE.',
    );
    expect(problems).toContain('dist/taskpane.html loads /assets/taskpane-abc.js, which is outside /excel-coach/.');
    expect(builtSiteProblems({ 'commands.html': '' }, BASE)).toEqual(['dist/taskpane.html is missing.']);
  });

  it('fails on any root-relative file outside the base path', () => {
    const problems = builtSiteProblems(site({ 'commands.html': '<link rel="icon" href="/favicon.ico">' }), BASE);
    expect(problems).toEqual(['dist/commands.html loads /favicon.ico, which is outside /excel-coach/.']);
  });

  it('fails when code would call this computer, at any port or address', () => {
    const code = ['const a = "http://localhost:3000/api/hint";', 'const b = `//127.0.0.1:8080`;', "const c = 'https://[::1]:3000';"].join('\n');
    expect(builtSiteProblems(site({ 'assets/taskpane-abc.js': code }), BASE)).toEqual([
      'dist/assets/taskpane-abc.js line 1 refers to this computer: http://localhost:3000/api/hint',
      'dist/assets/taskpane-abc.js line 2 refers to this computer: 127.0.0.1:8080',
      'dist/assets/taskpane-abc.js line 3 refers to this computer: https://[::1]:3000',
    ]);
    expect(builtSiteProblems(site({ 'assets/x.css': '@import url(https://localhost:5173/a.css);' }), BASE)).toEqual([
      'dist/assets/x.css line 1 refers to this computer: https://localhost:5173/a.css',
    ]);
  });

  it('ignores code that only names a local host without loading from it', () => {
    // Libraries check location.hostname this way, and minified code is full of numbers.
    const code = 'isLocalhost=location.hostname==="localhost"||h==="127.0.0.1";v="10.0.0.0";n=mylocalhost';
    expect(builtSiteProblems(site({ 'assets/taskpane-abc.js': code }), BASE)).toEqual([]);
    expect(localUrls(code).map((u) => u.url)).toEqual(['localhost', '127.0.0.1']);
  });
});

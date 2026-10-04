// Turns manifest.xml, which loads the panel from the local dev server, into the manifest for
// the copy hosted on GitHub Pages, and checks the built site. Pure string work, shared by
// scripts/build-pages.mjs and tests/pages.test.ts. Types live in pagesManifest.d.mts.

/** Where `npm run dev` serves the panel. manifest.xml points every URL here. */
export const LOCAL_ORIGIN = 'https://localhost:3000';

/**
 * The hosted add-in's Id. It differs from manifest.xml's so the hosted and local add-ins can be
 * installed side by side, and it never changes between builds so Excel sees each update as the
 * same add-in. Changing it strands every installed copy.
 */
export const HOSTED_ADDIN_ID = 'c16d946b-e5c7-4067-bc7a-5340a1bd6aa5';

/** The hosted add-in's name, and its ribbon button's label, so it stands apart from the local one. */
export const HOSTED_DISPLAY_NAME = 'Excel Coach (web)';

const EXAMPLE_URL = 'https://<user>.github.io/excel-coach/';

const escapeRe = (text) => text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

// LOCAL_ORIGIN followed by a path, or standing alone as an attribute or element value (AppDomain).
// Anything else, like another port, is left alone for localUrls() to catch.
const LOCAL_ORIGIN_RE = new RegExp(`${escapeRe(LOCAL_ORIGIN)}(/|(?=["'<\\s]))`, 'g');

// Any address on this computer, with or without a scheme or port. The lookarounds skip names and
// numbers that only contain one, like isLocalhost or 10.0.0.0.
const LOCAL_URL_RE = /(?<![\w-])(?:https?:\/\/)?(?:localhost|127(?:\.\d{1,3}){3}|\[::1\]|0\.0\.0\.0)(?![\w-])(?::\d+)?[^\s"'`<>()\\]*/gi;

// A local address code could load: one with a scheme or a port, unlike a bare hostname check.
const LOADABLE_RE = /^https?:\/\/|^[^/]+:\d/i;

const isLocalHost = (host) => /^(?:localhost|.+\.localhost|127(?:\.\d{1,3}){3}|\[::1\]|0\.0\.0\.0)$/i.test(host);

const xmlEscape = (text) =>
  text.replace(/[&<>"']/g, (ch) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&apos;' })[ch]);

/**
 * Checks PAGES_URL and returns it in canonical form, always ending in a slash:
 * `https://User.github.io/excel-coach` becomes `https://user.github.io/excel-coach/`.
 */
export function normalizePagesUrl(raw) {
  const value = raw?.trim();
  if (!value) throw new Error(`Set PAGES_URL to the address GitHub Pages serves the panel from, for example ${EXAMPLE_URL}`);
  let url;
  try {
    url = new URL(value);
  } catch {
    throw new Error(`PAGES_URL isn’t a web address: ${value}`);
  }
  if (url.protocol !== 'https:') throw new Error(`PAGES_URL must start with https://, because Excel loads add-ins only over HTTPS: ${value}`);
  if (isLocalHost(url.hostname)) throw new Error(`PAGES_URL points at this computer. Use the GitHub Pages address, for example ${EXAMPLE_URL}`);
  // The raw text, because URL drops an empty query or fragment from search and hash but keeps it in href.
  if (url.username || url.password || /[?#]/.test(value)) {
    throw new Error(`PAGES_URL can’t include a user name, a query or a # fragment: ${value}`);
  }
  if (/\/[^/]*\.[^/]*$/.test(url.pathname)) {
    throw new Error(`PAGES_URL is the site’s address, not a file. Drop the file name, or end a folder’s address with /: ${value}`);
  }
  if (!url.pathname.endsWith('/')) url.pathname += '/';
  return url.href;
}

/** The path Vite builds under: `/excel-coach/` for a project site, `/` for a user site or custom domain. */
export function basePathOf(pagesUrl) {
  return new URL(normalizePagesUrl(pagesUrl)).pathname;
}

/** Every address on this computer left in `xml`, with its 1-based line number. */
export function localUrls(xml) {
  return [...xml.matchAll(LOCAL_URL_RE)].map((m) => ({ line: xml.slice(0, m.index).split('\n').length, url: m[0] }));
}

/**
 * What would break the built site on Pages. `files` maps each built html, js and css file's path
 * in dist/ to its text. Returns one message per problem, so an empty list means the site is fine.
 */
export function builtSiteProblems(files, base) {
  const problems = [];
  // Without the base path, Pages serves the page but not its scripts, and the panel stays blank.
  const page = files['taskpane.html'];
  if (page === undefined) problems.push('dist/taskpane.html is missing.');
  else if (!page.includes(`src="${base}assets/`)) {
    problems.push(`dist/taskpane.html doesn’t load its scripts from ${base}assets/. Check that vite.config.ts still reads COACH_BASE.`);
  }
  for (const [file, text] of Object.entries(files)) {
    if (file.endsWith('.html')) {
      for (const [, ref] of text.matchAll(/\b(?:src|href)="(\/(?!\/)[^"]*)"/g)) {
        if (!ref.startsWith(base)) problems.push(`dist/${file} loads ${ref}, which is outside ${base}.`);
      }
    }
    // Nor may the panel call the dev server: on any other computer, localhost is that computer.
    for (const u of localUrls(text).filter((u) => LOADABLE_RE.test(u.url))) {
      problems.push(`dist/${file} line ${u.line} refers to this computer: ${u.url}`);
    }
  }
  return problems;
}

function replaceOne(xml, pattern, replace, what) {
  const count = [...xml.matchAll(new RegExp(pattern.source, 'g'))].length;
  if (count !== 1) throw new Error(`manifest.xml should have exactly one ${what}, but has ${count}.`);
  return xml.replace(pattern, replace);
}

/**
 * The hosted manifest: every LOCAL_ORIGIN URL moves to `pagesUrl` (a bare origin, as in
 * AppDomain, becomes the Pages origin), with HOSTED_ADDIN_ID, and HOSTED_DISPLAY_NAME as the
 * name and ribbon button label. Throws if any address on this computer is left, so a broken
 * manifest never ships.
 */
export function hostedManifest(source, pagesUrl) {
  const url = normalizePagesUrl(pagesUrl);
  const origin = new URL(url).origin;
  let xml = replaceOne(source, /<Id>[^<]*<\/Id>/, `<Id>${HOSTED_ADDIN_ID}</Id>`, '<Id>');
  xml = replaceOne(xml, /(<DisplayName\b[^>]*?\bDefaultValue=")[^"]*(")/, `$1${xmlEscape(HOSTED_DISPLAY_NAME)}$2`, '<DisplayName>');
  // The button's label is also its tooltip's title. Without this, both add-ins show the same button.
  xml = replaceOne(
    xml,
    /(<bt:String\s+id="CoachButton\.Label"\s+DefaultValue=")[^"]*(")/,
    `$1${xmlEscape(HOSTED_DISPLAY_NAME)}$2`,
    'CoachButton.Label string',
  );
  xml = xml.replace(LOCAL_ORIGIN_RE, (_, slash) => xmlEscape(slash === '/' ? url : origin));

  const leftovers = localUrls(xml);
  if (leftovers.length) {
    const list = leftovers.map((u) => `  line ${u.line}: ${u.url}`).join('\n');
    throw new Error(
      `The hosted manifest would still load from this computer:\n${list}\n` +
        `Only ${LOCAL_ORIGIN} URLs are moved to PAGES_URL. Point manifest.xml at that address, or update scripts/pagesManifest.mjs.`,
    );
  }
  return xml;
}

// Builds the panel for GitHub Pages, so Excel on Windows and Excel on the web can load it.
//   PAGES_URL=https://<user>.github.io/excel-coach/ npm run build:pages
// Writes the site to dist/ under PAGES_URL's path, plus dist/manifest.xml pointing at PAGES_URL.
// The GitHub workflow (.github/workflows/pages.yml) runs this; see HOSTING.md.
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { build } from 'vite';
import { HOSTED_ADDIN_ID, HOSTED_DISPLAY_NAME, basePathOf, builtSiteProblems, hostedManifest, normalizePagesUrl } from './pagesManifest.mjs';

const root = path.dirname(path.dirname(fileURLToPath(import.meta.url)));
const outDir = path.join(root, 'dist');

function fail(message) {
  console.error(`\n  ${message.replaceAll('\n', '\n  ')}\n`);
  process.exit(1);
}

let pagesUrl;
let manifest;
try {
  pagesUrl = normalizePagesUrl(process.env.PAGES_URL);
  // Transform before building, so a bad manifest fails in a second rather than after the build.
  manifest = hostedManifest(fs.readFileSync(path.join(root, 'manifest.xml'), 'utf8'), pagesUrl);
} catch (error) {
  fail(error.message);
}
const base = basePathOf(pagesUrl);

// vite.config.ts takes the base path from COACH_BASE. A build never serves, so COACH_HTTP skips
// loading the dev certificate, and the warning about it on machines that don't have one.
process.env.COACH_BASE = base;
process.env.COACH_HTTP = '1';
process.chdir(root);
await build();

// The pages must load their scripts from the base path and never call the dev server.
const built = Object.fromEntries(
  fs
    .readdirSync(outDir, { recursive: true })
    .filter((file) => /\.(?:html|js|css)$/.test(file))
    .map((file) => [file.split(path.sep).join('/'), fs.readFileSync(path.join(outDir, file), 'utf8')]),
);
const problems = builtSiteProblems(built, base);
if (problems.length) fail(`The built site wouldn’t work on GitHub Pages:\n${problems.map((p) => `  ${p}`).join('\n')}`);

fs.writeFileSync(path.join(outDir, 'manifest.xml'), manifest);

console.log(`\nBuilt Excel Coach for ${pagesUrl}`);
console.log(`  Site:      dist/ (base path ${base})`);
console.log(`  Manifest:  dist/manifest.xml ("${HOSTED_DISPLAY_NAME}", Id ${HOSTED_ADDIN_ID})`);
console.log('Check the manifest with `npm run validate:pages`. HOSTING.md covers publishing and installing.\n');

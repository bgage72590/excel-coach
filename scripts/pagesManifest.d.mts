// Types for pagesManifest.mjs, so tests can import it under the TypeScript project.

/** Where `npm run dev` serves the panel. manifest.xml points every URL here. */
export const LOCAL_ORIGIN: 'https://localhost:3000';

/** The hosted add-in's fixed Id, different from manifest.xml's. */
export const HOSTED_ADDIN_ID: string;

/** The hosted add-in's name and ribbon button label. */
export const HOSTED_DISPLAY_NAME: string;

export interface LocalUrl {
  /** 1-based line in the manifest. */
  line: number;
  url: string;
}

/** Checks PAGES_URL and returns it in canonical form, always ending in a slash. Throws on anything Excel can’t load. */
export function normalizePagesUrl(raw: string | undefined): string;

/** The path Vite builds under: `/excel-coach/` for a project site, `/` for a user site or custom domain. */
export function basePathOf(pagesUrl: string): string;

/** Every address on this computer left in `xml`, with its line number. */
export function localUrls(xml: string): LocalUrl[];

/**
 * What would break the built site on Pages, given each built html, js and css file's text keyed
 * by its path in dist/. One message per problem; an empty list means the site is fine.
 */
export function builtSiteProblems(files: Record<string, string>, base: string): string[];

/** The hosted manifest for `pagesUrl`. Throws if any address on this computer is left. */
export function hostedManifest(source: string, pagesUrl: string): string;

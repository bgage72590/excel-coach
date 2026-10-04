/// <reference types="vitest/config" />
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import { hintPlugin } from './server/hintPlugin';

// Excel only loads add-ins over HTTPS. `npm run certs` creates a trusted
// localhost certificate in ~/.office-addin-dev-certs; `npm run dev:http`
// skips it for previewing the panel in a normal browser.
function devHttps() {
  if (process.env.COACH_HTTP === '1' || process.env.VITEST) return undefined;
  const dir = path.join(os.homedir(), '.office-addin-dev-certs');
  const file = (name: string) => path.join(dir, name);
  if (!fs.existsSync(file('localhost.crt'))) {
    console.warn('\n  No dev certificate found. Run `npm run certs` once, then restart.\n');
    return undefined;
  }
  return {
    key: fs.readFileSync(file('localhost.key')),
    cert: fs.readFileSync(file('localhost.crt')),
    ca: fs.readFileSync(file('ca.crt')),
  };
}

export default defineConfig({
  // GitHub Pages serves the panel from /<repo>/; local dev serves it from the root.
  base: process.env.COACH_BASE ?? '/',
  plugins: [react(), hintPlugin()],
  server: { port: 3000, strictPort: true, https: devHttps() },
  build: {
    target: 'es2022',
    // Fluent UI dominates the bundle; it's served locally to Excel, so one chunk is fine.
    chunkSizeWarningLimit: 1200,
    rollupOptions: {
      input: { taskpane: 'taskpane.html', commands: 'commands.html' },
    },
  },
  test: { environment: 'node', include: ['tests/**/*.test.ts'] },
});

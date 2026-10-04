import { loadEnv, type Plugin } from 'vite';
import { handleHintRequest } from './hint';

/**
 * Local AI hints: the dev server answers GET /api/hint/status and POST /api/hint by asking Claude
 * to explain the learner's mistake on a practice sheet. The logic lives in hint.ts; this wires it
 * into Vite.
 *
 * Privacy: the API key is read from .env.local with loadEnv and stays in this Node process. It has
 * no VITE_ prefix, so Vite never exposes it to the panel, and no response or log line contains it.
 * Without a key the status route reports { enabled: false } and the panel hides the feature. A
 * static build (GitHub Pages) has no server at all, so hints are off there too.
 */
export function hintPlugin(): Plugin {
  return {
    name: 'coach-hints',
    apply: 'serve',
    configureServer(server) {
      const { mode, envDir, root, logger } = server.config;
      const apiKey = loadEnv(mode, envDir ?? root, 'ANTHROPIC_').ANTHROPIC_API_KEY?.trim() || undefined;
      const log = (message: string) => logger.warn(`[hints] ${message}`, { timestamp: true });
      if (apiKey) logger.info('  AI hints are on.');

      server.middlewares.use((req, res, next) => {
        handleHintRequest({ method: req.method, url: req.url, remoteAddress: req.socket.remoteAddress, headers: req.headers, body: req }, { apiKey, log })
          .then((out) => {
            if (!out) return next();
            res.statusCode = out.status;
            res.setHeader('Content-Type', 'application/json; charset=utf-8');
            res.setHeader('Cache-Control', 'no-store');
            res.setHeader('X-Content-Type-Options', 'nosniff');
            for (const [name, value] of Object.entries(out.headers ?? {})) res.setHeader(name, value);
            res.end(JSON.stringify(out.body));
          })
          .catch(next);
      });
    },
  };
}

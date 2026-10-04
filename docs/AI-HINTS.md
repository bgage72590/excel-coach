# AI hints

When a check fails on a practice sheet, **Explain my mistake** asks Claude (`claude-sonnet-5-5`) what's wrong with your formula. You get at most three short sentences: what's wrong, why Excel gives that result, and one nudge. Claude is told not to write the answer formula for you. Each explanation belongs to one check, so checking again clears it.

It's off by default. Without a key the button never appears, and nothing is sent anywhere.

## Turn it on

1. Copy `.env.example` to `.env.local` in the project folder. `.env.local` is git-ignored.
2. Paste your Anthropic API key after `ANTHROPIC_API_KEY=`.
3. Restart `npm run dev`. The terminal prints "AI hints are on."
4. Reload the panel in Excel (or close and reopen it).

To turn it off, delete the key (or `.env.local`) and restart `npm run dev`.

Vite's `loadEnv` also reads `ANTHROPIC_API_KEY` from the shell that runs `npm run dev`. If you export that variable in your shell profile, hints are on even without `.env.local`.

## Where it works

- **Exercises and missions only.** Hints are offered on practice sheets, which hold made-up data the coach generated. They're never offered on My Work, guided fixes or anything else that reads your own workbooks.
- **Only from the local dev server.** The `/api/hint` route is part of `npm run dev` on your computer. The GitHub Pages copy has no server, so the button doesn't appear there.

## Where the key lives

- Only in the dev server's Node process. The panel never receives it: the variable has no `VITE_` prefix, so Vite doesn't expose it to browser code, and no response contains it.
- It's sent only to `api.anthropic.com`, in the `x-api-key` header. It's never logged.
- The route answers requests from this computer only: the connection must come from a loopback address, and the `Host` and `Origin` must be `localhost` (or a `*.localhost` name), a 127.x.x.x address, or `[::1]`. A web page open in another tab can't use your key.

## Exactly what's sent

Each request carries these fields and nothing else. The server drops any other field and trims the rest before calling Claude.

| Field | What it is | Limit |
|-------|------------|-------|
| Title | The exercise title, or the mission and step title | 120 characters |
| Task | The task text you see in the panel | 1,200 characters |
| Platform | Excel for Mac, Windows or the web, so menu paths and shortcuts match yours | One of the three |
| Formulas | Your distinct formulas from the answer area, in reading order | 10 formulas, 500 characters each |
| Check results | Each check's label, status (passed, failed, not checked) and failure detail | 12 checks; labels 160 and details 240 characters |
| Hints shown | The coach's hints you've already opened | The latest 5, 300 characters each |

A filled-down column has a different formula in every row (`=C2*D2`, `=C3*D3`, …), so a long list is cut down by pattern: one formula of each shape is kept first, rarest first, and the rest of the room goes to the others in order. A cell you typed over, or one whose reference slipped a row, has a shape of its own, so it goes ahead of the repeats.

A failure detail can quote a value from the practice sheet, for example "`K7` showed 4.1 instead of 3.85". Those values are the coach's made-up data. No other cell values, sheet names or workbook contents are sent.

Requests over 16 KB are refused. Claude has 20 seconds to answer, and each reply is capped at 300 tokens.

## Cost and model settings

- Model `claude-sonnet-5-5` with thinking off (`between_tools`) and `low` effort, since the reply is three sentences.
- The request opts into server-side fallback (`fallbacks: "default"`, beta header `server-side-fallback-2026-07-01`). If Sonnet 5.5's safety classifier declines a request in a category that has a recommended fallback, the API reruns it on that model instead of returning a refusal. Other declines show "Claude couldn’t explain this one."
- A request is a few hundred input tokens and at most 300 output tokens. Nothing is sent until you select the button.

## When something goes wrong

The panel shows a short message with a **Try again** button. The terminal running `npm run dev` logs the API status, error type, request id and the API's error message, shortened and with the key masked. It never logs the key or what you sent. The request id helps if you need to look a request up in the Claude Console.

| Message | What to do |
|---------|------------|
| Claude didn’t accept the API key | The key is wrong or revoked. Fix `.env.local`, then restart `npm run dev`. |
| Your Anthropic account is out of credit | Add credit in the Claude Console. |
| This API key can’t use Claude Sonnet 5.5 | Check which workspace the key belongs to. |
| Couldn’t reach the hint server | `npm run dev` stopped. Start it again and reload the panel. |
| Claude couldn’t process this request | The API rejected the request (a 400). The line in the terminal usually names the field it rejected. |

## Code

- `server/hintPlugin.ts`: the Vite plugin. Reads the key and mounts the routes.
- `server/hint.ts`: the routes, validation, the prompt and the API call, testable without Vite.
- `src/ai/contract.ts`: the request shape and limits, shared by the panel and the server.
- `src/ai/hints.ts`: `hintStatus()` and `requestHint()` for the panel.
- `src/components/HintPanel.tsx`: the button and the "From Claude" card, used by ExerciseView and MissionView.
- `tests/hints.test.ts`: prompt, validation, limits, formula picking, reply tidying, loopback checks, error messages and logs, and that the key never appears in a response or a log line.

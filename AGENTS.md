# AGENTS.md

Office Add-in for Excel (KuroAgent) — AI chat task pane + custom functions. Built on the Microsoft Office Add-in template.

## Project Structure

```
src/
  functions/functions.js   — Excel custom functions (=KUROAGENT(...))
  taskpane/taskpane.js     — Task pane entry (Office.onReady → Excel.run)
  taskpane/agentChat.js    — Orchestrator + execution loop
  taskpane/agent/          — AI: ai.js (SSE streaming), parser.js (JSON extraction + repair),
                           — operations.js (60+ operation registry), prompts.js
  taskpane/excel/          — context.js, snapshot.js
  taskpane/ui/             — chat.js, config.js, paste.js
  taskpane/ppt/            — PowerPoint support (in development)
  commands/commands.js     — Ribbon button handler
```

**Build output:** `dist/` — served by webpack dev server on HTTPS `localhost:3000`.

**Three entry points:**
- `src/functions/functions.js` — Custom Excel functions (must use `@customfunction` JSDoc tag)
- `src/taskpane/taskpane.js` — Task pane (Excel API via `Office.onReady()` → `Excel.run()`)
- `src/commands/commands.js` — Ribbon button handler

## Commands

```bash
npm install
npm run dev-server        # webpack dev server (HTTPS, localhost:3000)
npm run build             # production build to dist/
npm run build:dev         # dev build with source maps
npm run watch             # rebuild on changes
npm run start             # build + launch Excel Desktop (Windows) — runs prestart hook
npm run start:web         # build + launch in browser
npm run stop              # stop debugging session
npm run lint              # office-addin-lint check
npm run lint:fix          # auto-fix lint issues
npm run validate          # validate manifest.xml
npm run test              # node --test tests/**/*.test.js
npm run test:watch        # watch mode for tests
```

## Dev CLI

After `npm link`, the `kuroagent` CLI replaces manual dev-server management:

```bash
kuroagent                          # Start dev server + sideload taskpane
kuroagent file.xlsx                # Start + open specified workbook
kuroagent -m manifest.xml          # Use a custom manifest
kuroagent --no-open                # Server only, don't open Excel
kuroagent --stop                   # Kill dev server on port 3000
kuroagent --update                 # Pull latest from GitHub, reinstall deps, relink CLI
kuroagent --install                # (Re)install Node dependencies (from any directory)
```

The CLI auto-discovers the project root (checks `EXCEL_HOME` env, `~/.kuroagent`, walks up from cwd, or parent of `bin/kuroagent`).

## Configuration

LLM settings are configured in the KuroAgent panel UI (Add-ins tab):

| Setting | Description | Default |
|---------|-------------|---------|
| Endpoint | LLM API URL | `https://openrouter.ai/api/v1/chat/completions` |
| Model | Model name | `gpt-4` |
| API Key | Bearer token (empty for local LLMs) | — |

Local LLM endpoints: Ollama `:11434`, LM Studio `:1234`, llama.cpp `:8082`, vLLM `:8000`. Any OpenAI-compatible `/v1/chat/completions` works.

Tailscale remote LLM: `tailscale serve --bg --https=443 http://localhost:8082`, then set endpoint to `https://<machine>.ts.net/v1/chat/completions`.

## Gotchas

- **Source is JavaScript, not TypeScript.** `tsconfig.json` is only for `@types/office-js` resolution. `package.json` has `"type": "module"` (ESM) — source files use `import`/`export`.
- **`manifest.xml` does NOT exist.** Active manifests are `manifest.dev.xml` (Excel) and `manifest.ppt.xml` (PowerPoint, in development). The CLI falls back to `manifest.dev.xml` when `manifest.xml` is absent.
- **Custom functions namespace is `KUROAGENT`** (not `CONTOSO`). Set in `manifest.dev.xml` `<bt:String id="Functions.Namespace">`.
- **`npm run start` has a `prestart` hook** that runs `npm run build` first.
- **Webpack config is CJS** (`webpack.config.cjs`) despite ESM source. Entry paths reference `./src/...` files.
- **Production build** swaps `https://localhost:3000/` → `urlProd` in `webpack.config.cjs:11-13` (resolves from `VERCEL_URL` env var, falls back to `https://excel-ten-theta.vercel.app/`).
- **`opencode.jsonc`** contains API keys — never commit changes to it. It's in `.npmignore`.
- **Prettier** uses `office-addin-prettier-config` (invoked via `npm run prettier`).
- **ESLint** extends `plugin:office-addins/recommended` (`.eslintrc.json`).
- **Babel** transpiles to ES5 for IE 11 compatibility (`babel.config.json` + `core-js`/`regenerator-runtime` polyfills).

## Tests

Tests exist for the parser/operations layer:

```bash
npm run test        # Run all tests: tests/operations.test.js
npm run test:watch  # Watch mode
```

Test scope: `parser.js` (JSON extraction from LLM responses), `operations.js` (operation validation against registry), and mock Excel execution tracking. Tests do NOT cover the task pane UI, custom functions, or actual Excel API calls.

## Verification

```bash
npm run lint              # MUST report 0 errors
npm run build             # MUST exit 0
npm run test              # MUST exit 0
```

## Custom Functions

All Excel-facing functions need the `@customfunction` JSDoc tag:

```javascript
/**
 * @customfunction
 */
function myFunc(input) { ... }
```

`functions.json` metadata is auto-generated by `CustomFunctionsMetadataPlugin` during build. Streaming functions (e.g., `clock`, `increment`) use `CustomFunctions.StreamingInvocation`.

## Deployment

**No `vercel.json` in this repo.** The `webpack.config.cjs` `urlProd` default is `https://excel-ten-theta.vercel.app/`. Production deployment is configured via Vercel CLI or the Vercel dashboard.

**After deployment:**
1. Update `manifest.dev.xml` URLs to point to your production URL
2. Validate: `npm run validate`
3. Sideload the updated manifest in Excel

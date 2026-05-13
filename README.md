# open-excel — KuroAgent Local Development

Office Add-in for Excel with AI chat — local development package.

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Excel](https://img.shields.io/badge/Excel-2016%2F365-blue)](https://www.office.com)

> This is the open-source local development mirror of KuroAgent.
> It contains everything needed to run the add-in locally with your own LLM.

## What is this?

KuroAgent lets you chat with any LLM directly inside Excel and automate spreadsheet tasks with natural language.

This repo has the source code, install scripts, and CLI tooling. No cloud deployment — just clone, install, and run against any OpenAI-compatible LLM endpoint (local or cloud).

## Quick Install

### Windows

```powershell
iwr https://raw.githubusercontent.com/yannassoumou/open-excel/master/install.ps1 -OutFile install.ps1
powershell -ExecutionPolicy Bypass -File .\install.ps1
```

### Linux / macOS

```bash
curl -fsSL https://raw.githubusercontent.com/yannassoumou/open-excel/master/install.sh | bash
```

Both scripts install dependencies, generate HTTPS dev certs, and put the `excel` CLI on your PATH.

### Manual setup (without the one-liner)

```bash
git clone https://github.com/yannassoumou/open-excel.git
cd open-excel
npm install
npx office-addin-dev-certs install  # HTTPS dev certs for webpack
npm link                            # puts `excel` on PATH
```

## Developer CLI

Once installed, the `excel` CLI replaces the manual `npm run dev-server` workflow:

```bash
# Start dev server (launches it automatically if not running)
excel

# Use a specific manifest
excel -m path/to/manifest.xml

# Start server only, no browser/Excel
excel --no-open

# Stop the dev server
excel --stop

excel --help                        # full help
```

The CLI starts the webpack dev server on `https://localhost:3000` (or the port in `EXCEL_PORT`), waits until it's ready, then opens Excel with the add-in sideloaded.

On macOS and Linux, Excel desktop sideloading is not supported via registry — the CLI opens Excel Online in your browser instead.

## Configuration

Open the KuroAgent panel in Excel (Add-ins tab), then set:

| Setting | Description | Default |
|---------|-------------|---------|
| **Endpoint** | LLM API URL | `https://openrouter.ai/api/v1/chat/completions` |
| **Model** | Model name | `gpt-4` |
| **API Key** | Bearer token (leave empty for local LLMs) | — |

### Supported Local LLM Endpoints

| Tool | Default URL |
|------|-------------|
| Ollama | `http://localhost:11434/v1/chat/completions` |
| LM Studio | `http://localhost:1234/v1/chat/completions` |
| llama.cpp | `http://localhost:8082/v1/chat/completions` |
| vLLM | `http://localhost:8000/v1/chat/completions` |

Any OpenAI-compatible `/v1/chat/completions` endpoint works.

## Tailscale Setup (for remote LLM servers)

If your LLM runs on another machine on your Tailscale network:

1. On the LLM machine, expose the server via `tailscale serve`:
   ```
   tailscale serve --bg --https=443 http://localhost:8082
   ```
2. In the KuroAgent panel, set the endpoint to `https://<machine>.ts.net/v1/chat/completions`

Tailscale provides TLS so Excel Online (HTTPS) won't block it with mixed content errors.

## Excel Online (no install)

If you just want to try it without installing anything:

1. Open [excel.office.com](https://excel.office.com)
2. Open a blank workbook
3. **Insert -> Get Add-ins -> Custom Add-ins**
4. Load manifest from your running dev server: `https://localhost:3000/manifest.xml`

## Features

- **AI Chat** — chat with any LLM inside Excel
- **60+ operations** — write values/formulas, format cells, build charts, pivot tables, tables, conditional formatting, shapes, data validation, comments, named ranges
- **JSON auto-repair** — handles malformed LLM responses automatically
- **Batch execution** — 2-5 operations per round with step progress
- **Custom Functions** — `=KUROAGENT("your prompt")` in cells
- **Multi-sheet context** — reads workbook structure (sheets, tables, charts) for smarter LLM answers

## Architecture

```
User → Chat UI → LLM (streaming) → JSON Parser → Operation Validator → Excel API
                ↑                                              ↓
                └──── Context (sheet data) ←───────────────────┘
```

Key files:
- `src/taskpane/agentChat.js` — orchestrator + execution loop
- `src/taskpane/agent/ai.js` — SSE streaming to OpenAI-compatible APIs
- `src/taskpane/agent/parser.js` — JSON extraction + repair
- `src/taskpane/agent/operations.js` — 60+ operation registry + executor
- `src/taskpane/agent/prompts.js` — system prompt architecture
- `src/taskpane/excel/context.js` — workbook state reading
- `src/taskpane/ui/chat.js` — chat rendering, collapsible step groups

## Build commands

```bash
npm run build            # production build → dist/ (not in this repo)
npm run dev-server       # webpack dev server on https://localhost:3000
npm run start            # build + launch Excel Desktop (Windows)
npm run start:web        # build + launch in browser
npm run lint             # ESLint check
npm run test             # run tests
npm run validate         # validate manifest.xml
```

## License

MIT — see [LICENSE](LICENSE)

## Author

**Yann Assoumou** — [GitHub](https://github.com/yannassoumou)

Full source and production deployment live at [yannassoumou/excel](https://github.com/yannassoumou/excel).

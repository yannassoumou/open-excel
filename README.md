# open-excel — KuroAgent Local Dev

Office Add-in for Excel with AI chat — local development package.

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Excel](https://img.shields.io/badge/Excel-2016%2F365-blue)](https://www.office.com)

> This is the open-source local development mirror of KuroAgent.
> Clone, install, and run against any OpenAI-compatible LLM (local or cloud).

## Quick Install

### Windows

```powershell
iwr https://raw.githubusercontent.com/yannassoumou/open-excel/master/setup.ps1 -OutFile setup.ps1
powershell -ExecutionPolicy Bypass -File .\setup.ps1
```

### Linux / macOS

```bash
curl -fsSL https://raw.githubusercontent.com/yannassoumou/open-excel/master/setup.sh | bash
```

Both offer an interactive menu with **install**, **update**, **uninstall**, and **purge** options, or pass the action directly:

```powershell
powershell -ExecutionPolicy Bypass -File .\setup.ps1 install
powershell -ExecutionPolicy Bypass -File .\setup.ps1 update
powershell -ExecutionPolicy Bypass -File .\setup.ps1 uninstall
powershell -ExecutionPolicy Bypass -File .\setup.ps1 purge
```

### Manual Setup

```bash
git clone https://github.com/yannassoumou/open-excel.git
cd open-excel
npm install
npx office-addin-dev-certs install
npm link
```

## Developer CLI

After install, the `kuroagent` CLI replaces `npm run dev-server`:

```bash
kuroagent                          # Start dev server + sideload taskpane
kuroagent -m path/to/manifest.xml  # Use a custom manifest
kuroagent --no-open                # Server only
kuroagent --stop                   # Stop dev server on port 3000
kuroagent --help                   # Full help
```

The CLI starts the webpack server on `https://localhost:3000`, waits until it is ready, then opens Excel Desktop (Windows) or your default browser (Mac/Linux) with the add-in sideloaded. Keep the terminal open for live reloading.

## Configuration

Open the KuroAgent panel in Excel (Add-ins tab), then set:

| Setting | Description | Default |
|---------|-------------|---------|
| **Endpoint** | LLM API URL | `https://openrouter.ai/api/v1/chat/completions` |
| **Model** | Model name | `gpt-4` |
| **API Key** | Bearer token (empty for local LLMs) | — |

### Local LLM Endpoints

| Tool | URL |
|------|-----|
| Ollama | `http://localhost:11434/v1/chat/completions` |
| LM Studio | `http://localhost:1234/v1/chat/completions` |
| llama.cpp | `http://localhost:8082/v1/chat/completions` |
| vLLM | `http://localhost:8000/v1/chat/completions` |

Any OpenAI-compatible `/v1/chat/completions` endpoint works.

### Tailscale Remote LLM

If your LLM runs on another machine on your Tailscale network, expose it with:

```
tailscale serve --bg --https=443 http://localhost:8082
```

Then in KuroAgent set the endpoint to `https://<machine>.ts.net/v1/chat/completions`. Tailscale provides TLS so Excel Online won't block with mixed content errors.

## Excel Online (No Install)

1. Open [excel.office.com](https://excel.office.com)
2. Open a blank workbook
3. **Insert -> Get Add-ins -> Custom Add-ins**
4. Load manifest from your running dev server: `https://localhost:3000/manifest.xml`

## Features

- **AI Chat** — chat with any LLM inside Excel
- **60+ operations** — write values/formulas, format cells, build charts, pivot tables, tables, conditional formatting, shapes, data validation, comments, named ranges
- **JSON auto-repair** — handles malformed LLM responses automatically
- **Batch execution** — 2-5 operations per round with progress tracking
- **Custom Functions** — `=KUROAGENT("your prompt")` in cells
- **Multi-sheet context** — reads workbook structure for smarter LLM answers

## Architecture

- `src/taskpane/agentChat.js` — orchestrator + execution loop
- `src/taskpane/agent/ai.js` — SSE streaming to OpenAI-compatible APIs
- `src/taskpane/agent/parser.js` — JSON extraction + repair
- `src/taskpane/agent/operations.js` — 60+ operation registry + executor
- `src/taskpane/agent/prompts.js` — system prompt architecture
- `src/taskpane/excel/context.js` — workbook state reading
- `src/taskpane/ui/chat.js` — chat rendering, collapsible step groups

## Build Commands

```bash
npm run build            # production build to dist/
npm run dev-server       # webpack dev server on https://localhost:3000
npm run start            # build + launch Excel Desktop (Windows)
npm run start:web        # build + launch in browser
npm run lint             # ESLint check
npm run validate         # validate manifest.xml
```

## Uninstall

Use the same setup script:

```powershell
powershell -ExecutionPolicy Bypass -File .\setup.ps1 uninstall
```

```bash
curl -fsSL https://raw.githubusercontent.com/yannassoumou/open-excel/master/setup.sh | bash -s -- uninstall
```

Use `purge` instead of `uninstall` to delete the `.kuroagent` directory too.

## License

MIT -- see [LICENSE](LICENSE)

## Author

**Yann Assoumou** -- [GitHub](https://github.com/yannassoumou)

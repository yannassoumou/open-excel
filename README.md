# KuroAgent — AI in Excel

> Chat with any LLM directly inside Excel. Automate your spreadsheets with natural language.

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Excel](https://img.shields.io/badge/Excel-2016%2F365-blue)](https://www.office.com)
[![Status](https://img.shields.io/badge/Status-Active-brightgreen)]()
[![GitHub issues](https://img.shields.io/github/issues/yannassoumou/excel)](https://github.com/yannassoumou/excel/issues)

## Features

- **AI Chat in Excel** — Open the KuroAgent panel and chat with your LLM to automate Excel tasks
- **60+ Excel Operations** — Create worksheets, write data, format cells, build charts, pivot tables, tables, conditional formatting, shapes, and more
- **Multi-Provider Support** — OpenRouter, Anthropic Claude, Google Gemini, OpenAI GPT, or any OpenAI-compatible local LLM (Ollama, LM Studio, llama.cpp, vLLM)
- **Smart JSON Parser** — Automatic JSON repair for malformed LLM responses (trailing commas, missing brackets, comments)
- **Batch Execution** — 2-5 operations per batch with step-by-step progress tracking
- **Custom Functions** — Use `=KUROAGENT("your prompt")` directly in Excel cells
- **Session Telemetry** — Anonymous usage tracking for continuous improvement (opt-out available)
- **Cross-Platform** — Works on Excel Desktop (Windows/Mac), Excel Online, and Excel for Linux

## Installation

### For Developers (One-Command Setup)

This installs all dependencies and puts the `excel` CLI on your PATH:

**Linux / Mac:**
```bash
curl -fsSL https://raw.githubusercontent.com/yannassoumou/excel/main/install.sh | bash
```

**Windows:**
```powershell
iwr https://raw.githubusercontent.com/yannassoumou/excel/main/install.ps1 -OutFile install.ps1
powershell -ExecutionPolicy Bypass -File .\install.ps1
```

### For End Users (Just Want the Add-in in Excel)

**Windows (registry sideloading, auto-load on startup):**
```powershell
iwr https://raw.githubusercontent.com/yannassoumou/excel/main/install.ps1 -OutFile install.ps1
powershell -ExecutionPolicy Bypass -File .\install.ps1
```
Then restart Excel. Click the KuroAgent button in the ribbon to open the panel.

**Linux / Mac / Excel Online (recommended):**
1. Open [excel.office.com](https://excel.office.com)
2. Open a blank workbook
3. **Insert -> Get Add-ins -> Custom Add-ins**
4. Load manifest from: `https://excel-ten-theta.vercel.app/manifest.xml`

## Developer CLI

After running `install.sh`, the `excel` CLI is on your PATH. It replaces the old `npm run dev-server` workflow.

### Quick Start

```bash
# Start dev server + sideload
excel

# Use a custom manifest
excel -m path/to/manifest.xml

# Start server only (no Excel/browser)
excel --no-open

# Stop the dev server
excel --stop

# Full help
excel --help
```

### Commands

| Command | What it does |
|---------|-------------|
| `excel` | Checks if dev server is running, starts it if not, waits for ready, opens Excel/browser |
| `excel workbook.xlsx` | Same as above, plus records the workbook path |
| `excel -m PATH` | Use a custom manifest instead of `manifest.dev.xml` |
| `excel --no-open` | Start server only, skip sideload |
| `excel --stop` | Kill the dev server on port 3000 |

### Manual Dev (Without CLI)

If you prefer the raw npm workflow:

```bash
git clone https://github.com/yannassoumou/excel.git
cd excel
npm install
npm run dev-server        # webpack dev server (HTTPS localhost:3000)
npm run start             # Build + launch Excel Desktop
```

## Configuration

Open the KuroAgent panel (HOME tab → **Open KuroAgent** button), then configure:

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
| llama.cpp | `http://localhost:8080/v1/chat/completions` |
| vLLM | `http://localhost:8000/v1/chat/completions` |
| text-generation-webui | `http://localhost:5000/v1/chat/completions` |

## Excel Operations

KuroAgent supports 60+ Excel operations across these categories:

- **Worksheets**: create, delete, rename, copy, hide, show, protect, freeze panes
- **Data**: write values, write formulas, read ranges, sort, delete rows/columns
- **Formatting**: font, fill, borders, alignment, wrap, merge, autofit
- **Tables**: create, delete, add rows/columns, sort, filter, style
- **Charts**: add, delete, set title/series/position/legend/data labels
- **PivotTables**: create, add hierarchies, set aggregation, layout, filters
- **Conditional Formatting**: color scales, data bars, icon sets, text comparison
- **Shapes**: add, delete, set text
- **Data Validation**: add, delete rules
- **Comments & Notes**: add, delete
- **Named Ranges**: add, delete

## Custom Functions

Use AI directly in your cells:

```
=KUROAGENT("Summarize the trends in this data")
=KUROAGENT("Translate to English", "fr→en")
=KUROAGENT("Excel formula to calculate monthly growth")
```

## Architecture

```
User → Chat UI → LLM (streaming) → JSON Parser → Operation Validator → Excel API
                ↑                                              ↓
                └──── Context (sheet data) ←───────────────────┘
```

- **Agent Orchestrator** (`src/taskpane/agentChat.js`) — State management, execution loop
- **AI Module** (`src/taskpane/agent/ai.js`) — SSE streaming chat completions
- **Parser** (`src/taskpane/agent/parser.js`) — JSON extraction with automatic repair
- **Operations** (`src/taskpane/agent/operations.js`) — 60+ operation registry + executor
- **Prompts** (`src/taskpane/agent/prompts.js`) — 3-layer system prompt architecture
- **Context** (`src/taskpane/excel/context.js`) — Workbook state reading (sheets, tables, charts)
- **UI** (`src/taskpane/ui/chat.js`) — Chat rendering with collapsible step groups
- **Telemetry** (`src/taskpane/telemetry.js`) — Anonymous session tracking

## Build & Development

```bash
npm install              # Install dependencies
excel                    # Start everything (replaces npm run dev-server)
npm run build            # Production build → dist/
npm run start            # Build + launch Excel Desktop for debugging
npm run start:web        # Debug in Excel on the web
npm run lint             # ESLint check
npm run test             # Run test suite
npm run validate         # Validate manifest.xml
```

## Project Status

| Component | Status |
|-----------|--------|
| Excel Add-in | **Active** — Fully functional |
| PowerPoint Add-in | **In Development** — Architecture designed, not yet released |
| Custom Functions | **Active** — `=KUROAGENT()` works in cells |
| Telemetry Dashboard | **Active** — Separate repo |

> **Note:** The PowerPoint add-in is still in development. Only the Excel add-in is available in this release.

## License

MIT License — see [LICENSE](LICENSE) for details.

## Author

**Yann Assoumou** — [GitHub](https://github.com/yannassoumou)

## Acknowledgments

Built on the [Microsoft Office Add-in template](https://github.com/OfficeDev/Excel-Custom-Functions-JS). Uses [Office.js](https://learn.microsoft.com/en-us/office/dev/add-ins/reference/office-js/), [Fluent UI](https://developer.microsoft.com/en-us/fluentui#/), and [Webpack 5](https://webpack.js.org/).

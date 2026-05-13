# Contributing to KuroAgent

Thank you for your interest in contributing to KuroAgent! This document covers everything you need to get started.

## Getting Started

### Prerequisites

- **Node.js** (LTS version recommended)
- **npm** (comes with Node.js)
- **Microsoft Excel** (Microsoft 365, Excel 2016 or later)
- **Windows** (for Excel Desktop debugging) or **Excel Online** (cross-platform)

### Setup

```bash
# 1. Fork and clone the repo
git clone https://github.com/YOUR_USERNAME/excel.git
cd excel

# 2. Install dependencies
npm install

# 3. Install Office dev certificates
npm install -g office-addin-dev-certs
office-addin-dev-certs install

# 4. Start the dev server
npm run dev-server

# 5. Sideloading in Excel
npm run start           # Excel Desktop
npm run start:web       # Excel Online
```

## Project Structure

```
src/
├── functions/
│   └── functions.js           # Custom Excel functions (KUROAGENT namespace)
├── taskpane/
│   ├── agentChat.js           # Main orchestrator (state, execution loop)
│   ├── agent/
│   │   ├── ai.js              # LLM streaming (SSE)
│   │   ├── parser.js          # JSON extraction + repair
│   │   ├── prompts.js         # System prompts (3-layer architecture)
│   │   └── operations.js      # 60+ operation registry + executor
│   ├── excel/
│   │   ├── context.js         # Workbook context reading
│   │   └── snapshot.js        # Snapshot tracking
│   ├── ui/
│   │   ├── chat.js            # Chat UI + step grouping
│   │   ├── config.js          # Configuration panel
│   │   └── paste.js           # Paste-to-AI preview
│   ├── telemetry.js           # Anonymous session tracking
│   ├── taskpane.js            # Host detection + init
│   ├── taskpane.html
│   └── agentChat.css
└── commands/
    └── commands.js            # Ribbon button handler
```

## Adding New Operations

To add a new Excel operation:

1. **Define the operation** in `src/taskpane/agent/operations.js`:
   ```javascript
   myNewOperation: {
     required: ["sheet", "range", "value"],
     validate: (op) => typeof op.sheet === "string" && typeof op.range === "string",
     execute: async (op, context) => {
       const range = context.workbook.getRange(op.range);
       range.values = [[op.value]];
       await context.sync();
       return "Done";
     },
   },
   ```

2. **Add to the prompt** in `src/taskpane/agent/prompts.js` (operation reference section)

3. **Add a test** in `tests/operations.test.js`

4. **Test manually** in Excel with `npm run start`

## Code Style

- **JavaScript** (ES5+ with Babel transpilation)
- **ESLint** with `office-addins` preset (`npm run lint`)
- **Prettier** with `office-addin-prettier-config` (`npm run prettier`)
- **No TypeScript source** — TypeScript types are only for Office.js (`@types/office-js`)

## Testing

```bash
npm test          # Run all tests
npm run test:watch # Watch mode
```

Tests cover:
- JSON parser (extraction, repair, edge cases)
- Operation validator (all 60+ operations)
- Mock Excel execution
- Edge cases (empty input, malformed JSON, etc.)

## Commit Guidelines

- Use conventional commits: `feat:`, `fix:`, `docs:`, `chore:`, `refactor:`
- Keep commits focused — one change per commit
- Reference issues in commit messages: `fix: repair JSON trailing commas #42`

## Pull Requests

1. Create a feature branch: `git checkout -b feat/my-feature`
2. Make your changes
3. Run `npm run lint` and `npm test` — both must pass
4. Update documentation if needed
5. Submit a PR with a clear description

## Reporting Issues

Use [GitHub Issues](https://github.com/yannassoumou/excel/issues) with:

- **Bug Report**: steps to reproduce, expected vs actual behavior, Excel version
- **Feature Request**: description, use case, why it matters
- **Question**: tag as `question` — we'll help you out

## Architecture Notes

### 3-Layer System Prompt
The agent uses a 3-layer prompt strategy to combat context window overflow:
1. **System Prompt** — Dense rules at start and end (primacy + recency)
2. **Context Block** — Workbook state injected before user query
3. **Turn Reminder** — One-liner prepended to every user message (~30 tokens)

### JSON Repair
LLMs frequently produce malformed JSON. The parser includes automatic repair for:
- Trailing commas
- Missing closing brackets
- Single-line comments inside JSON
- Unclosed strings
- Trailing text after JSON block

### Operation Batching
Operations are limited to 2-5 per batch to avoid Excel API timeouts and ensure reliable execution. The agent plans multiple batches for complex tasks.

---

Thanks for helping make KuroAgent better!

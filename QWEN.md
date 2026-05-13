# Excel Custom Functions Add-in (exceljs)

## Project Overview

This is an **Office Add-in** for Microsoft Excel that provides **custom functions** and a **task pane** UI. It is built on the Microsoft Office Add-in project template ("Excel-Custom-Functions-JS") and enables users to define and execute JavaScript-based Excel functions directly within spreadsheets, alongside a task pane for interacting with the Excel object model.

### Key Technologies

- **Runtime**: Microsoft Office JavaScript API (`Office.js`, `Excel.js`, Custom Functions Runtime)
- **Language**: JavaScript (with TypeScript types for Office APIs)
- **Build**: Webpack 5 + Babel (transpilation to ES5)
- **Styling**: Fluent UI (Office UI Fabric)
- **Linting**: ESLint with `office-addins` preset
- **Target**: Excel Desktop (Windows/Mac) and Excel on the web

### Architecture

The add-in consists of three main entry points:

| Entry Point | Location | Purpose |
|---|---|---|
| **Custom Functions** | `src/functions/functions.js` | Defines Excel custom functions (e.g., `add`, `clock`, `increment`, `logMessage`) |
| **Task Pane** | `src/taskpane/taskpane.js` + `taskpane.html` | A side panel UI for interacting with Excel via the Excel JavaScript API |
| **Commands** | `src/commands/commands.js` | Ribbon button action handler |

The manifest (`manifest.xml`) declares:
- A `CustomFunctions` extension point that registers the `CONTOSO` namespace for custom functions
- A task pane button on the HOME ribbon tab
- `ReadWriteDocument` permissions

## Building and Running

### Prerequisites

- Node.js (LTS)
- npm
- Microsoft Excel (part of Microsoft 365)

### Commands

| Command | Description |
|---|---|
| `npm install` | Install dependencies |
| `npm run build` | Production build (Webpack, generates `dist/`) |
| `npm run build:dev` | Development build with source maps |
| `npm run watch` | Watch mode — rebuilds on file changes |
| `npm run dev-server` | Start the webpack dev server (HTTPS on `localhost:3000`) |
| `npm run start` | Build and launch the add-in for debugging in Excel Desktop |
| `npm run stop` | Stop a running debugging session |
| `npm run lint` | Check code with ESLint |
| `npm run lint:fix` | Auto-fix linting issues |
| `npm run validate` | Validate `manifest.xml` |

### VS Code Tasks

The project includes predefined VS Code tasks (see `.vscode/tasks.json`):

- **Build (Development)** — default build task
- **Build (Production)** — production build
- **Debug: Excel Desktop** — starts debugging in Excel
- **Dev Server** — runs the webpack dev server
- **Lint: Check for problems** / **Lint: Fix all auto-fixable problems**
- **Watch** — file watching build

## Development Conventions

### Code Style

- Prettier configuration via `office-addin-prettier-config`
- ESLint rules from `plugin:office-addins/recommended`
- ES5 target with polyfills (`core-js`, `regenerator-runtime`)
- TypeScript compiler options are present but the source is JavaScript (types are provided by `@types/office-js` and `@types/custom-functions-runtime`)

### Custom Functions

Functions intended for Excel must be decorated with the `@customfunction` JSDoc tag. Example:

```javascript
/**
 * Add two numbers
 * @customfunction
 * @param {number} first First number
 * @param {number} second Second number
 * @returns {number} The sum of the two numbers.
 */
function add(first, second) {
  return first + second;
}
```

Streaming (invocation-based) functions are supported for real-time updates (see `clock` and `increment` in `src/functions/functions.js`).

### Task Pane

The task pane uses the Excel JavaScript API inside an `Office.onReady()` callback. Code runs within an `Excel.run()` context, calling `context.sync()` to push/pull changes.

### Manifest

- The add-in is sideloaded for development via `https://localhost:3000/`
- The custom functions namespace is `CONTOSO`
- Functions script is served from `/public/functions.js`, metadata from `/public/functions.json`
- For production, update `urlProd` in `webpack.config.js` and the corresponding URLs in `manifest.xml`

### Output Structure

Webpack outputs to `dist/` with the following key files:

```
dist/
├── functions.html      # Custom functions host page
├── taskpane.html       # Task pane UI
├── commands.html       # Commands page
├── functions.js        # Custom functions bundle
├── taskpane.js         # Task pane bundle
├── commands.js         # Commands bundle
├── polyfill.js         # Polyfill bundle
├── assets/             # Static assets (icons)
└── functions.json      # Auto-generated custom functions metadata
```

## Key Files Reference

| File | Purpose |
|---|---|
| `manifest.xml` | Office Add-in manifest (metadata, permissions, extension points) |
| `webpack.config.js` | Build configuration — entry points, plugins, dev server |
| `babel.config.json` | Babel transpilation preset (`@babel/preset-env`) |
| `tsconfig.json` | TypeScript compiler settings (used for type checking) |
| `.eslintrc.json` | ESLint configuration extending `office-addins` recommended rules |
| `src/functions/functions.js` | Custom function definitions |
| `src/taskpane/taskpane.js` | Task pane logic (Excel API interactions) |
| `src/commands/commands.js` | Ribbon command action handler |

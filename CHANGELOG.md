# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added
- One-command install scripts for Windows (`install.ps1`) and Linux/Mac (`install.sh`)
- MIT License
- Open-source project structure with CONTRIBUTING.md
- Installation documentation in README

### Changed
- Custom functions namespace: `CONTOSO` → `KUROAGENT`
- Project branding: all references updated to KuroAgent
- Manifest files reorganized (dev vs prod)

### Removed
- Temp files from repository root
- Template-only content from README

## [1.0.8] — 2025-05-06

### Added
- 60+ Excel operations (worksheets, data, formatting, tables, charts, pivot tables, conditional formatting, shapes, data validation, comments, named ranges, slicers)
- JSON parser with automatic repair (trailing commas, missing brackets, comments)
- 3-layer system prompt architecture
- Step grouping UI with collapsible progress
- Telemetry system (sessions, steps, snapshots, feedback, error tracking)
- Custom functions: `=KUROAGENT()`, `=KUROAGENT_STREAM()`
- Paste-to-AI rich context preview
- Support for local LLMs (Ollama, LM Studio, llama.cpp, vLLM)
- 68 automated tests (parser, validator, mock execution)

### Fixed
- Excel JS API pitfalls (merge vs format.merge, load before set)
- Field name alias normalization for LLM responses
- Slow provider timeout handling (120s first token, 60s subsequent)
- Context reading optimization (batched loads, single sync)

## [1.0.0] — Initial Release

- Base template from Microsoft Office Add-in (Excel-Custom-Functions-JS)
- Custom functions framework
- Task pane UI
- Webpack build system

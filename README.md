# Head Draftsmen App

A desktop **mail-merge** app: fill multiple Word (`.docx`) templates from multiple Excel
(`.xlsx`/`.xls`) files, then export and print the results.

Built with **Electron + React + Vite + TypeScript**. All document logic lives in a
framework-agnostic `core/` module so the same React UI can later back a web app.

## Features

- Upload **multiple `.docx` templates** with `{{ColumnName}}` placeholders.
- Upload **multiple Excel files**; columns are auto-detected.
- **Column-collision detection** across Excel files — resolve which file wins before generating.
- **Placeholder mapping** view shows which placeholders have data.
- Generate **one document per row, per template**.
- **Export** to `.docx`, PDF, or both (your choice per run).
- **Print** generated PDFs from the app.

## How templates work

Put placeholders in your Word document using double braces:

```
Dear {{Name}},

Your account {{AccountId}} is ready.
```

Column headers in your Excel files must match the placeholder names (`Name`, `AccountId`).

## Multi-file merging

Multiple Excel files are merged **by row position** (row _n_ of each file forms record _n_),
and their columns are unioned. If two files share a column name, Head Draftsmen App flags a **collision**
and asks you to choose the winning source before generating.

## Scripts

| Command | Description |
|---|---|
| `npm run dev` | Run the app in development (hot reload). |
| `npm run build` | Build main, preload, and renderer bundles. |
| `npm start` | Preview the production build. |
| `npm test` | Run core unit tests (Vitest). |
| `npm run typecheck` | Type-check node + web projects. |
| `npm run package` | Build a distributable installer (electron-builder). |
| `npm run package:dir` | Build an unpacked app (fast, for testing). |

> **PDF export** uses headless **LibreOffice** for high-fidelity conversion. Install LibreOffice
> to enable PDF output; otherwise generate `.docx` only.

## Project structure

```
electron/   Electron main process + preload IPC bridge
src/        React renderer (portable, web-reusable)
core/       Framework-agnostic logic: excel, templates, merge, generate, export, print
tests/      Unit tests for core modules
```

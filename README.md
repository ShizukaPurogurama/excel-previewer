# Excel Previewer

A lightweight browser-based Excel viewer for inspecting worksheets, selecting header rows, and parsing table data — all without a backend.

## Features

- Load `.xlsx`, `.xls`, `.xlsb`, `.ods`, `.csv`
- Sheet switching
- Header-row selection (auto-detected, manually overridable)
- Parsed table view with currency-aware money formatting (USD, KHR)
- Column visibility control (up to 10 columns) with search
- Manual and auto-reload modes (via File System Access API when available)
- Merged-cell awareness in preview and parsed views
- Click-to-copy for cells and entire columns (Excel-compatible HTML)
- Theme toggle (light/dark) that follows OS preference by default
- Floating quick-actions (changelog, columns, reload)

## Local development

This project is now a Vite + React app.

```bash
npm install
npm run dev
```

The dev server prints a local URL (default `http://localhost:5173/excel-previewer/`).

## Production build

```bash
npm run build
npm run preview
```

`npm run build` emits a static bundle into `dist/` that can be served from any static host.

## Deployment (GitHub Pages)

The `base` in [vite.config.js](vite.config.js) is set to `/excel-previewer/` so that asset URLs resolve under the GitHub Pages project path.

Pushes to `main` are built and deployed by
[.github/workflows/pages-build-deployment.yml](.github/workflows/pages-build-deployment.yml):

1. Install dependencies.
2. Run `npm run build`.
3. Upload `dist/` as the Pages artifact.
4. Publish with `actions/deploy-pages`.

If you fork this repo under a different name, update `base` in `vite.config.js` to match.

## Project layout

```
src/
  main.jsx            # React entry
  App.jsx             # Top-level layout
  styles.css          # Design system
  constants.js        # Presets, storage keys, changelog, version
  components/         # Presentational components
  hooks/              # useTheme, useToasts, useWorkbook, ...
  utils/              # xlsx, clipboard, money, parsedTable, ...
```

## Tech stack

- [Vite](https://vitejs.dev/) + [React 18](https://react.dev/)
- [SheetJS](https://sheetjs.com/) (`xlsx`) for workbook parsing
- Native `<dialog>`, File System Access API, Clipboard API

## Copyright

Copyright (c) SEANG SENGLY

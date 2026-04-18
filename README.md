# Excel Previewer

[![Vite](https://img.shields.io/badge/Vite-5.0+-645FFF?logo=vite&logoColor=white)](https://vitejs.dev/)
[![React](https://img.shields.io/badge/React-18+-61DAFB?logo=react&logoColor=white)](https://react.dev/)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![GitHub Pages](https://img.shields.io/badge/Deployed-GitHub%20Pages-blue.svg)](https://github.com/pages)

A lightweight, client-side Excel viewer for inspecting worksheets, selecting header rows, and parsing table data — no backend required.

## Quick overview

- **File types:** `.xlsx`, `.xls`, `.xlsb`, `.ods`, `.csv`
- **Highlights:** sheet switching, header-row selection (auto/manual), merged-cell aware preview, parsed table view with currency-aware formatting (USD, KHR), column visibility and search, click-to-copy, and theme toggle (light/dark).

## Preview

![Excel Previewer Demo](./assets/demo.gif)

**Features in action:** drag a spreadsheet file, inspect sheets side-by-side, auto-detect or manually pick header rows, and export parsed tables with click-to-copy cells and columns.

## Quick start

1. Install dependencies:

```bash
npm install
```

2. Run development server (Vite):

```bash
npm run dev
```

The dev server prints a local URL (default `http://localhost:5173/excel-previewer/`).

3. Build for production:

```bash
npm run build
npm run preview
```

`npm run build` emits a static bundle into `dist/` that can be hosted on any static host.

## Usage notes

- When available the File System Access API enables auto-reload while editing files.
- Click any cell to copy as plain text, or click a column header to copy column HTML compatible with Excel.
- Parsed table view detects and formats currency values; configure currency presets in `src/constants.js`.

## Deployment

This repo is configured for GitHub Pages. The `base` option in [vite.config.js](vite.config.js) is set to `/excel-previewer/` so assets resolve under the project path.

Pushes to `main` are deployed via the workflow at [.github/workflows/pages-build-deployment.yml](.github/workflows/pages-build-deployment.yml): install dependencies, run `npm run build`, upload `dist/`, then publish.

If you fork the project, update `base` in [vite.config.js](vite.config.js) to match your Pages path.

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

- Vite + React 18
- SheetJS (`xlsx`) for workbook parsing
- Native browser APIs: File System Access, Clipboard, `<dialog>`

## Contributing

Contributions are welcome. Please open issues for bugs or feature requests and submit pull requests for fixes.

## License & author

Copyright (c) SEANG SENGLY

Licensed under the terms in this repository.

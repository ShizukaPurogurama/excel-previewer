# Static Excel Viewer

A lightweight browser-based Excel viewer for inspecting worksheets, selecting header rows, and parsing table data — all without a backend.

## Features

- Load `.xlsx`, `.xls`, `.xlsb`, `.ods`, `.csv`
- Sheet switching
- Header-row selection
- Parsed table view
- Column visibility control (up to 7 columns)
- Manual and auto reload modes
- Merged-cell awareness and preview rendering
- Click-to-copy for table cells
- Theme toggle (light/dark)
- Floating settings actions

## Usage

1. Open `index.html` in a browser (or run a local server).
2. Choose a workbook file.
3. Select sheet + header row.
4. Open **Manage columns** (floating settings) to choose displayed columns.
5. Use reload options as needed.

## Local Run

You can run a simple local server from project root:

```bash
python -m http.server 3000
```

Then open:

- http://localhost:3000

## Tech Stack

- HTML
- CSS
- Vanilla JavaScript
- [SheetJS](https://sheetjs.com/) via CDN

## Copyright

Copyright (c) SEANG SENGLY ***
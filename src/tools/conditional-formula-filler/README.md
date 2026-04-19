# Conditional Formula Filler

Fills Excel cells based on multi-column IF/AND conditions — no formula knowledge required.

Equivalent to a nested formula like:
```
=IF(AND(B5="PSP",D5="KHR"),D2,IF(AND(B5="PSP",D5="USD"),E2,""))
```
…but built visually through a 6-step wizard.

## How to use

1. **Upload** — pick an `.xlsx`, `.xls`, or `.csv` file. The header row is auto-detected; override it by typing a row number or clicking a row in the preview table.

2. **Lookup Values** — define the cases you want to fill. Each case needs a unique label (e.g. `PSP-KHR`) and a fill value (e.g. `009899838`).

3. **Conditions** — choose which columns to check (e.g. `Sender Type`, `Currency`), then pick the exact cell value each case requires in each column.

4. **Targets** — choose which columns to fill. Use *Same rules* to write the same value to all targets, or *Per target* to set a different value per column.

5. **Preview** — inspect the first 20 rows. Changed cells show `before → after`. Summary counts show how many rows will be filled.

6. **Export** — download a new `.xlsx` file. Choose *Entire workbook* to preserve all sheets (only the selected sheet is updated), or *Selected sheet only* for a single-sheet file. The original file is never modified.

## Persistence

Wizard state (including the workbook) is saved automatically to `localStorage` under the key `conditional-formula-filler:v1`. Workbooks larger than 4 MB are stored in IndexedDB instead. Clicking **Clear saved data** wipes all stored state.

## Technical notes

- All processing runs in the browser — no data is uploaded anywhere.
- Uses the `xlsx-js-style` package for reading and writing workbooks.
- Original cell types and styles are preserved where possible.
- Handles files with tens of thousands of rows without freezing the UI.

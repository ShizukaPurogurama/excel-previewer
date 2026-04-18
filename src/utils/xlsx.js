import * as XLSX from 'xlsx';

export async function readWorkbookFromFile(file) {
  const data = await file.arrayBuffer();
  const workbook = XLSX.read(data, {
    type: 'array',
    cellDates: true,
  });

  if (!workbook.SheetNames.length) {
    throw new Error('This workbook does not contain any readable sheets.');
  }

  return workbook;
}

export function alignRowsToExcelRowNumbers(sheet, rows) {
  const reference = sheet['!ref'];
  if (!reference) return rows;

  const range = XLSX.utils.decode_range(reference);
  if (!range.s.r) return rows;

  return Array.from({ length: range.s.r }, () => []).concat(rows);
}

export function expandMergedCellsForDisplay(sheet, rows) {
  const covered = new Set();
  const anchors = new Map();
  const merges = sheet && Array.isArray(sheet['!merges']) ? sheet['!merges'] : [];

  merges.forEach((range) => {
    if (!range || !range.s || !range.e) return;

    const startRow = range.s.r;
    const endRow = range.e.r;
    const startCol = range.s.c;
    const endCol = range.e.c;

    const anchorRow = rows[startRow] || [];
    const anchorValue = anchorRow[startCol];

    for (let rowIndex = startRow; rowIndex <= endRow; rowIndex += 1) {
      if (!rows[rowIndex]) rows[rowIndex] = [];

      for (let colIndex = startCol; colIndex <= endCol; colIndex += 1) {
        if (rowIndex === startRow && colIndex === startCol) continue;

        const cellValue = rows[rowIndex][colIndex];
        if (cellValue === '' || cellValue === null || cellValue === undefined) {
          rows[rowIndex][colIndex] = anchorValue;
        }

        covered.add(rowIndex + ':' + colIndex);
      }
    }

    anchors.set(startRow + ':' + startCol, {
      rowSpan: endRow - startRow + 1,
      colSpan: endCol - startCol + 1,
    });
  });

  return { covered, anchors };
}

export function getMergeStats(sheet) {
  const merges = sheet && Array.isArray(sheet['!merges']) ? sheet['!merges'] : [];
  let rowMerged = 0;
  let columnMerged = 0;
  let bothMerged = 0;

  merges.forEach((range) => {
    if (!range || !range.s || !range.e) return;
    const spansRows = range.e.r > range.s.r;
    const spansColumns = range.e.c > range.s.c;
    if (spansColumns) rowMerged += 1;
    if (spansRows) columnMerged += 1;
    if (spansRows && spansColumns) bothMerged += 1;
  });

  return {
    total: merges.length,
    rowMerged,
    columnMerged,
    bothMerged,
  };
}

export function formatMergeSummary(stats) {
  if (!stats || !stats.total) return 'No merged cells detected.';
  return (
    'Merged regions: ' +
    stats.total +
    ' (row merges: ' +
    stats.rowMerged +
    ', column merges: ' +
    stats.columnMerged +
    ').'
  );
}

export function extractSheetData(workbook, sheetName) {
  const sheet = workbook.Sheets[sheetName];
  const rawRows = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    raw: true,
    defval: '',
    blankrows: true,
  });
  const rows = alignRowsToExcelRowNumbers(sheet, rawRows);
  const mergeMaps = expandMergedCellsForDisplay(sheet, rows);
  const stats = getMergeStats(sheet);

  return {
    rows,
    mergedCovered: mergeMaps.covered,
    mergeAnchors: mergeMaps.anchors,
    mergeStats: stats,
  };
}

export function supportsFileSystemAccess() {
  return typeof window !== 'undefined' && typeof window.showOpenFilePicker === 'function';
}

import {
  clampHeaderIndex,
  formatCellValue,
  isEmptyCell,
  makeUniqueHeaderName,
  normalizeCell,
} from './columns.js';
import {
  buildAmountCurrencyColumnLookup,
  formatAmountCellValue,
  resolveMoneyMappingForHeaders,
} from './money.js';

export function buildParsedTable({
  rows,
  headerIndex,
  mergedCovered,
  mergeAnchors,
  moneyMappings,
}) {
  if (!rows || !rows.length) {
    return {
      headers: [],
      records: [],
      mergedCarryRows: [],
      mergedAnchors: [],
      moneyMappingState: null,
    };
  }

  const safeHeaderIndex = clampHeaderIndex(headerIndex, rows.length);
  const headerRow = rows[safeHeaderIndex] || [];
  const dataRows = rows.slice(safeHeaderIndex + 1);
  const maxColumnCount = Math.max(
    headerRow.length,
    ...dataRows.map((row) => row.length),
    0
  );

  if (!maxColumnCount) {
    return {
      headers: [],
      records: [],
      mergedCarryRows: [],
      mergedAnchors: [],
      moneyMappingState: null,
    };
  }

  const activeColumns = [];
  for (let columnIndex = 0; columnIndex < maxColumnCount; columnIndex += 1) {
    const headerText = normalizeCell(headerRow[columnIndex]);
    const hasDataBelow = dataRows.some((row) => !isEmptyCell(row[columnIndex]));
    if (headerText || hasDataBelow) activeColumns.push(columnIndex);
  }

  if (!activeColumns.length) {
    return {
      headers: [],
      records: [],
      mergedCarryRows: [],
      mergedAnchors: [],
      moneyMappingState: null,
    };
  }

  const seenHeaders = new Map();
  const headers = activeColumns.map((columnIndex) =>
    makeUniqueHeaderName(headerRow[columnIndex], columnIndex, seenHeaders)
  );

  const moneyMappingState = resolveMoneyMappingForHeaders(headers, dataRows, moneyMappings);
  const amountCurrencyColumns = buildAmountCurrencyColumnLookup(
    headers,
    moneyMappingState.mapping
  );

  const records = [];
  const mergedCarryRows = [];
  const mergedAnchorRows = [];

  const coveredSet = mergedCovered || new Set();
  const anchorsMap = mergeAnchors || new Map();

  dataRows.forEach((row, dataRowIndex) => {
    const hasData = activeColumns.some((columnIndex) => !isEmptyCell(row[columnIndex]));
    if (!hasData) return;

    const sourceRowIndex = safeHeaderIndex + 1 + dataRowIndex;

    records.push(
      activeColumns.map((columnIndex, activeHeaderIndex) => {
        const rawValue = row[columnIndex];
        const amountText = formatAmountCellValue(rawValue, {
          activeColumns,
          activeHeaderIndex,
          amountCurrencyColumns,
          headers,
          row,
          sourceColumnIndex: columnIndex,
          moneyMapping: moneyMappingState.mapping,
        });
        if (amountText !== null && amountText !== undefined) return amountText;
        return formatCellValue(rawValue);
      })
    );

    mergedCarryRows.push(
      activeColumns.map((columnIndex) => coveredSet.has(sourceRowIndex + ':' + columnIndex))
    );

    mergedAnchorRows.push(
      activeColumns.map(
        (columnIndex) => anchorsMap.get(sourceRowIndex + ':' + columnIndex) || null
      )
    );
  });

  return {
    headers,
    records,
    mergedCarryRows,
    mergedAnchors: mergedAnchorRows,
    moneyMappingState,
  };
}

export function resolveDefaultColumnSelection(headers, sheetName, presets, maxVisible) {
  const activeHeaders = Array.isArray(headers) ? headers.slice() : [];
  const safeTotal = activeHeaders.length;
  if (!safeTotal) return { indexes: [], source: 'empty' };

  const presetNames = resolveColumnPreset(sheetName, presets);
  if (presetNames && presetNames.length) {
    const presetIndexes = [];
    activeHeaders.forEach((header, index) => {
      if (presetNames.indexOf(String(header).toLowerCase()) !== -1) {
        presetIndexes.push(index);
      }
    });

    if (presetIndexes.length) {
      presetIndexes.sort(
        (a, b) =>
          presetNames.indexOf(String(activeHeaders[a]).toLowerCase()) -
          presetNames.indexOf(String(activeHeaders[b]).toLowerCase())
      );
      return { indexes: presetIndexes.slice(0, maxVisible), source: 'preset' };
    }
  }

  return {
    indexes: Array.from({ length: Math.min(maxVisible, safeTotal) }, (_, index) => index),
    source: 'first-columns',
  };
}

function resolveColumnPreset(sheetName, presets) {
  if (!presets || typeof presets !== 'object') return null;

  const sheetKey = Object.keys(presets).find(
    (key) => key !== '*' && key.toLowerCase() === (sheetName || '').toLowerCase()
  );
  const names = sheetKey ? presets[sheetKey] : presets['*'];
  if (!Array.isArray(names) || !names.length) return null;

  return names.map((n) => String(n).toLowerCase());
}

export function loadColumnSelectionsFromStorage(storageKey, headers) {
  try {
    const sig = headers && headers.length ? headers.join('|') : '';
    if (!sig) return null;

    const raw = window.localStorage.getItem(storageKey);
    if (!raw) return null;

    const data = JSON.parse(raw);
    return data[sig] || null;
  } catch (e) {
    return null;
  }
}

export function saveColumnSelectionsToStorage(storageKey, headers, selectedIndexes) {
  try {
    const sig = headers && headers.length ? headers.join('|') : '';
    if (!sig) return;

    const raw = window.localStorage.getItem(storageKey);
    const data = raw ? JSON.parse(raw) : {};
    data[sig] = Array.from(selectedIndexes);

    const keys = Object.keys(data);
    if (keys.length > 50) delete data[keys[0]];

    window.localStorage.setItem(storageKey, JSON.stringify(data));
  } catch (e) {
    // ignore
  }
}

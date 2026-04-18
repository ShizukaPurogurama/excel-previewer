export function toSpreadsheetColumn(index) {
  let label = '';
  let value = index + 1;
  while (value > 0) {
    const remainder = (value - 1) % 26;
    label = String.fromCharCode(65 + remainder) + label;
    value = Math.floor((value - 1) / 26);
  }
  return label;
}

export function normalizeCell(value) {
  if (value === null || value === undefined) return '';
  if (value instanceof Date) return value.toISOString().slice(0, 10);
  if (typeof value === 'string') return value.trim();
  return String(value).trim();
}

export function isEmptyCell(value) {
  return normalizeCell(value) === '';
}

export function clampHeaderIndex(index, rowCount) {
  if (rowCount <= 0) return 0;
  const safeIndex = Number.isFinite(index) ? Math.floor(index) : 0;
  return Math.min(Math.max(safeIndex, 0), rowCount - 1);
}

export function isNumericLike(value) {
  const text = typeof value === 'string' ? value.trim() : normalizeCell(value);
  return /^[-+]?[$]?\d[\d,]*(\.\d+)?%?$/.test(text);
}

export function isTextLike(value) {
  return /[A-Za-z]/.test(value);
}

export function countNonEmpty(row) {
  return (row || []).reduce(
    (count, value) => count + (isEmptyCell(value) ? 0 : 1),
    0
  );
}

export function arraysEqual(left, right) {
  if (!Array.isArray(left) || !Array.isArray(right) || left.length !== right.length) {
    return false;
  }
  for (let i = 0; i < left.length; i += 1) {
    if (left[i] !== right[i]) return false;
  }
  return true;
}

export function getHeaderSignature(headers) {
  if (!headers || !headers.length) return '';
  return headers.join('|');
}

export function makeUniqueHeaderName(value, columnIndex, seenHeaders) {
  const baseName = normalizeCell(value) || 'Column ' + (columnIndex + 1);
  const key = baseName.toLowerCase();
  const count = seenHeaders.get(key) || 0;
  seenHeaders.set(key, count + 1);
  if (!count) return baseName;
  return baseName + ' (' + (count + 1) + ')';
}

export function formatCellValue(value) {
  if (value === null || value === undefined || value === '') return '';
  if (value instanceof Date) {
    return value.toLocaleDateString(undefined, {
      year: 'numeric',
      month: 'short',
      day: 'numeric',
    });
  }
  return String(value);
}

export function formatInterval(ms) {
  if (ms < 1000) return ms + 'ms';
  if (ms < 60000) return Math.floor(ms / 1000) + 's';
  return Math.floor(ms / 60000) + 'min';
}

export function normalizeColumnPickerSearch(value) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, ' ');
}

export function normalizeHeaderLabel(header) {
  return normalizeCell(header)
    .toLowerCase()
    .replace(/[_-]+/g, ' ')
    .replace(/[^a-z0-9\s]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

// Kept: scoring algorithm for auto-detecting the header row — pure logic, no UI.
import { countNonEmpty, isNumericLike, isTextLike, normalizeCell } from './columns.js';

export function detectHeaderRow(rows) {
  if (!rows.length) return 0;

  const limit = Math.min(rows.length, 25);
  let bestIndex = 0;
  let bestScore = Number.NEGATIVE_INFINITY;
  let firstNonEmptyIndex = 0;
  let foundNonEmptyRow = false;

  for (let rowIndex = 0; rowIndex < limit; rowIndex += 1) {
    const row = rows[rowIndex] || [];
    const values = row.map(normalizeCell).filter(Boolean);
    if (!values.length) continue;

    if (!foundNonEmptyRow) {
      firstNonEmptyIndex = rowIndex;
      foundNonEmptyRow = true;
    }

    const uniqueCount = new Set(values.map((value) => value.toLowerCase())).size;
    const textLikeCount = values.filter(isTextLike).length;
    const numericLikeCount = values.filter(isNumericLike).length;
    const nextRows = rows.slice(rowIndex + 1, rowIndex + 4);
    const belowDensity =
      nextRows.reduce((sum, nextRow) => sum + countNonEmpty(nextRow), 0) /
      Math.max(nextRows.length, 1);
    const supportingRows = nextRows.filter((nextRow) => countNonEmpty(nextRow) > 0).length;

    let score = values.length * 2;
    score += uniqueCount * 1.25;
    score += textLikeCount * 2;
    score += Math.min(belowDensity, values.length);
    score += supportingRows;
    score -= Math.max(0, numericLikeCount - textLikeCount) * 1.5;

    if (values.length === 1) score -= 5;
    if (rowIndex > 0 && countNonEmpty(rows[rowIndex - 1]) === 0) score += 1;

    if (score > bestScore) {
      bestScore = score;
      bestIndex = rowIndex;
    }
  }

  return foundNonEmptyRow ? bestIndex : firstNonEmptyIndex;
}

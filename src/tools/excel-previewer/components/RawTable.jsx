import { useMemo, useRef } from 'react';
import { PREVIEW_ROW_LIMIT } from '../../../constants.js';
import { formatCellValue, toSpreadsheetColumn } from '../utils/columns.js';
import { useCellCopy } from '../hooks/useCellCopy.js';

function resolveWindow(totalRows, headerIndex) {
  if (!totalRows) return { start: 0, end: 0 };
  const start =
    headerIndex >= PREVIEW_ROW_LIMIT
      ? Math.max(0, headerIndex - Math.floor(PREVIEW_ROW_LIMIT / 3))
      : 0;
  return { start, end: Math.min(totalRows, start + PREVIEW_ROW_LIMIT) };
}

export default function RawTable({
  isVisible,
  rows,
  headerIndex,
  mergedCovered,
  mergeAnchors,
  mergeSummary,
  onSelectHeaderRow,
  showStatus,
}) {
  const { handleTableClick } = useCellCopy({ showStatus });
  const copyContextRef = useRef(null);

  const totalRows = rows.length;
  const { start, end } = resolveWindow(totalRows, headerIndex);

  const computed = useMemo(() => {
    if (!totalRows) return null;

    const visibleRows = rows.slice(start, end);
    const headerRowCells = rows[headerIndex] || [];
    const columnCount = Math.max(
      1,
      headerRowCells.length,
      ...visibleRows.map((r) => r.length)
    );

    const colLabels = Array.from({ length: columnCount }, (_, i) =>
      toSpreadsheetColumn(i)
    );

    const rowItems = visibleRows.map((row, vi) => {
      const actualIndex = start + vi;
      const isHdr = actualIndex === headerIndex;
      const cells = [];

      for (let c = 0; c < columnCount; c++) {
        const key = actualIndex + ':' + c;
        const anchor = mergeAnchors ? mergeAnchors.get(key) : null;
        const covered = mergedCovered ? mergedCovered.has(key) : false;
        if (covered && !anchor) continue;

        cells.push({
          columnIndex: c,
          text: formatCellValue(row[c]),
          isEmpty: formatCellValue(row[c]) === '',
          anchor,
        });
      }

      return { actualIndex, isHdr, cells };
    });

    return { columnCount, colLabels, rowItems };
  }, [rows, totalRows, start, end, headerIndex, mergedCovered, mergeAnchors]);

  copyContextRef.current = useMemo(() => {
    if (!computed) return null;
    const headers = ['Row', ...computed.colLabels];
    const records = computed.rowItems.map(({ actualIndex }) => {
      const src = rows[actualIndex] || [];
      return [
        'Row ' + (actualIndex + 1),
        ...Array.from({ length: computed.columnCount }, (_, i) =>
          formatCellValue(src[i])
        ),
      ];
    });
    return { headers, records };
  }, [computed, rows]);

  if (!isVisible) return null;

  if (!totalRows) {
    return (
      <div className="ep-empty">
        <p className="ep-empty-title">Sheet is empty</p>
        <p className="ep-empty-sub">
          This sheet contains no rows. Try selecting a different sheet from the toolbar.
        </p>
      </div>
    );
  }

  const metaText =
    'Rows ' +
    (start + 1) +
    '–' +
    end +
    ' of ' +
    totalRows +
    ' · Click a row button to use it as the header · ' +
    mergeSummary;

  const handleClick = (e) => {
    const btn = e.target.closest('button[data-row-index]');
    if (btn) {
      const idx = Number.parseInt(btn.dataset.rowIndex, 10);
      if (Number.isInteger(idx)) {
        e.stopPropagation();
        onSelectHeaderRow(idx);
      }
      return;
    }
    handleTableClick(e, copyContextRef.current);
  };

  return (
    <div>
      <p className="ep-data-meta">{metaText}</p>
      <div
        className="ep-table-shell"
        onClick={handleClick}
        role="grid"
        aria-label="Raw worksheet rows"
      >
        <table id="ep-raw-table">
          <thead>
            <tr>
              <th scope="col">Row</th>
              {computed.colLabels.map((label) => (
                <th key={label} scope="col">
                  {label}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {computed.rowItems.map(({ actualIndex, isHdr, cells }) => (
              <tr
                key={actualIndex}
                className={isHdr ? 'is-selected' : undefined}
              >
                <th scope="row">
                  <button
                    type="button"
                    className={'row-selector' + (isHdr ? ' is-active' : '')}
                    data-row-index={String(actualIndex)}
                    aria-pressed={isHdr ? 'true' : 'false'}
                    title={
                      isHdr
                        ? 'Current header row'
                        : 'Use row ' + (actualIndex + 1) + ' as header'
                    }
                  >
                    Row {actualIndex + 1}
                  </button>
                </th>
                {cells.map((cell) => {
                  const cls = [
                    cell.isEmpty && 'cell-empty',
                    cell.anchor && 'cell-merged-anchor',
                  ]
                    .filter(Boolean)
                    .join(' ');

                  const tdProps = { className: cls || undefined };
                  if (cell.anchor?.rowSpan > 1)
                    tdProps.rowSpan = cell.anchor.rowSpan;
                  if (cell.anchor?.colSpan > 1)
                    tdProps.colSpan = cell.anchor.colSpan;
                  if (cell.anchor)
                    tdProps.title =
                      'Merged (' +
                      cell.anchor.rowSpan +
                      '×' +
                      cell.anchor.colSpan +
                      ')';

                  return (
                    <td key={cell.columnIndex} {...tdProps}>
                      {cell.text || (cell.isEmpty ? '—' : cell.text)}
                    </td>
                  );
                })}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

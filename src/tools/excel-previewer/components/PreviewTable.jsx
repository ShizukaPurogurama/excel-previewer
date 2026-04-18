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
  const end = Math.min(totalRows, start + PREVIEW_ROW_LIMIT);
  return { start, end };
}

export default function PreviewTable({
  isVisible,
  rows,
  headerIndex,
  mergedCovered,
  mergeAnchors,
  mergeSummary,
  onSelectHeaderRow,
  showStatus,
}) {
  const totalRows = rows.length;
  const { start, end } = resolveWindow(totalRows, headerIndex);

  const { handleTableClick } = useCellCopy({ showStatus });
  const copyContextRef = useRef(null);

  const computed = useMemo(() => {
    if (!totalRows) return null;

    const visibleRows = rows.slice(start, end);
    const headerRowCells = rows[headerIndex] || [];
    const columnCount = Math.max(
      1,
      headerRowCells.length,
      ...visibleRows.map((row) => row.length)
    );

    const columnLabels = [];
    for (let i = 0; i < columnCount; i += 1) columnLabels.push(toSpreadsheetColumn(i));

    const rowItems = visibleRows.map((row, visibleRowIndex) => {
      const actualRowIndex = start + visibleRowIndex;
      const isSelectedHeader = actualRowIndex === headerIndex;
      const cells = [];

      for (let columnIndex = 0; columnIndex < columnCount; columnIndex += 1) {
        const mergeKey = actualRowIndex + ':' + columnIndex;
        const anchor = mergeAnchors ? mergeAnchors.get(mergeKey) : null;
        const covered = mergedCovered ? mergedCovered.has(mergeKey) : false;

        if (covered && !anchor) continue;

        const value = row[columnIndex];
        const text = formatCellValue(value);

        cells.push({
          columnIndex,
          text,
          isEmpty: !text,
          anchor,
        });
      }

      return { actualRowIndex, isSelectedHeader, cells };
    });

    return { visibleRows, columnCount, columnLabels, rowItems };
  }, [rows, totalRows, start, end, headerIndex, mergedCovered, mergeAnchors]);

  copyContextRef.current = useMemo(() => {
    if (!computed) return null;
    const previewHeaders = ['Row', ...computed.columnLabels];
    const records = computed.rowItems.map(({ actualRowIndex }) => {
      const record = ['Row ' + (actualRowIndex + 1)];
      const sourceRow = rows[actualRowIndex] || [];
      for (let i = 0; i < computed.columnCount; i += 1) {
        record.push(formatCellValue(sourceRow[i]));
      }
      return record;
    });
    return { headers: previewHeaders, records };
  }, [computed, rows]);

  if (!isVisible) return null;

  if (!totalRows) {
    return (
      <section id="preview-section" className="panel data-panel">
        <div className="panel-heading">
          <div>
            <p className="section-label">Sheet preview</p>
            <h2>Raw worksheet rows</h2>
          </div>
          <p id="preview-meta" className="panel-meta">
            {'The selected sheet is empty. ' + mergeSummary}
          </p>
        </div>
        <div id="preview-empty" className="empty-state">
          This sheet has no rows to preview.
        </div>
        <div id="preview-wrap" className="table-shell is-hidden">
          <table id="preview-table"></table>
        </div>
      </section>
    );
  }

  const meta =
    'Showing rows ' +
    (start + 1) +
    '-' +
    end +
    ' of ' +
    totalRows +
    '. Click a row number to use it as the header. ' +
    mergeSummary;

  const handleClick = (event) => {
    const button = event.target.closest('button.row-selector');
    if (button) {
      const rowIndex = Number.parseInt(button.dataset.rowIndex, 10);
      if (Number.isInteger(rowIndex)) {
        event.stopPropagation();
        onSelectHeaderRow(rowIndex);
      }
      return;
    }
    handleTableClick(event, copyContextRef.current);
  };

  return (
    <section id="preview-section" className="panel data-panel">
      <div className="panel-heading">
        <div>
          <p className="section-label">Sheet preview</p>
          <h2>Raw worksheet rows</h2>
        </div>
        <p id="preview-meta" className="panel-meta">
          {meta}
        </p>
      </div>
      <div id="preview-empty" className="empty-state is-hidden">
        Select a workbook to inspect sheet rows before parsing.
      </div>
      <div id="preview-wrap" className="table-shell" onClick={handleClick}>
        <table id="preview-table">
          <thead>
            <tr>
              <th>Row</th>
              {computed.columnLabels.map((label) => (
                <th key={label}>{label}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {computed.rowItems.map(({ actualRowIndex, isSelectedHeader, cells }) => (
              <tr key={actualRowIndex} className={isSelectedHeader ? 'is-selected' : undefined}>
                <th scope="row">
                  <button
                    type="button"
                    className={'row-selector' + (isSelectedHeader ? ' is-active' : '')}
                    data-row-index={String(actualRowIndex)}
                  >
                    Row {actualRowIndex + 1}
                  </button>
                </th>
                {cells.map((cell) => {
                  const classes = [];
                  if (cell.isEmpty) classes.push('cell-empty');
                  if (cell.anchor) classes.push('cell-merged-anchor');
                  const title = cell.anchor
                    ? 'Merged cell (' + cell.anchor.rowSpan + '×' + cell.anchor.colSpan + ')'
                    : undefined;
                  const tdProps = {
                    className: classes.join(' ') || undefined,
                    title,
                  };
                  if (cell.anchor && cell.anchor.rowSpan > 1) tdProps.rowSpan = cell.anchor.rowSpan;
                  if (cell.anchor && cell.anchor.colSpan > 1) tdProps.colSpan = cell.anchor.colSpan;
                  return (
                    <td key={cell.columnIndex} {...tdProps}>
                      {cell.text || '-'}
                    </td>
                  );
                })}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </section>
  );
}

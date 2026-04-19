import { useMemo, useRef } from 'react';
import { useCellCopy } from '../hooks/useCellCopy.js';

function buildMeta(records, columns, headerIndex, mergeSummary) {
  return (
    'Header: row ' +
    (headerIndex + 1) +
    ' · ' +
    records +
    (records === 1 ? ' record' : ' records') +
    ' · ' +
    columns +
    (columns === 1 ? ' column' : ' columns') +
    ' · ' +
    mergeSummary
  );
}

export default function DataTable({
  isVisible,
  parsedTable,
  selectedColumnIndexes,
  headerIndex,
  mergeSummary,
  showStatus,
}) {
  const { handleTableClick } = useCellCopy({ showStatus });
  const copyContextRef = useRef(null);

  const { headers, records, mergedCarryRows, mergedAnchors } = parsedTable;
  const hasHeaders = headers.length > 0;

  const tableData = useMemo(() => {
    if (!hasHeaders) return { selectedHeaders: [], selectedRows: [] };

    const selectedHeaders = selectedColumnIndexes.map((idx) => headers[idx]);

    const selectedRows = records.map((record, rowIndex) => {
      const flags = mergedCarryRows[rowIndex] || [];
      const anchors = mergedAnchors[rowIndex] || [];
      const cells = [];

      for (const colIdx of selectedColumnIndexes) {
        const isCarry = Boolean(flags[colIdx]);
        const anchor = anchors[colIdx] || null;
        if (isCarry && !anchor) continue;
        cells.push({ colIdx, value: record[colIdx], anchor });
      }

      return cells;
    });

    return { selectedHeaders, selectedRows };
  }, [hasHeaders, headers, records, mergedCarryRows, mergedAnchors, selectedColumnIndexes]);

  copyContextRef.current = useMemo(() => {
    if (!tableData.selectedHeaders.length) return null;
    return {
      headers: tableData.selectedHeaders.slice(),
      records: records.map((row) =>
        selectedColumnIndexes.map((idx) => row[idx])
      ),
    };
  }, [tableData.selectedHeaders, records, selectedColumnIndexes]);

  if (!isVisible) return null;

  if (!hasHeaders) {
    return (
      <div className="ep-empty">
        <p className="ep-empty-title">No columns detected</p>
        <p className="ep-empty-sub">
          The chosen header row produced no column names. Switch to Raw rows
          and click a row button to select a better header.
        </p>
      </div>
    );
  }

  const metaText = buildMeta(
    records.length,
    headers.length,
    headerIndex,
    mergeSummary
  );

  return (
    <div>
      <p className="ep-data-meta">{metaText}</p>
      <div
        className="ep-table-shell"
        onClick={(e) => handleTableClick(e, copyContextRef.current)}
        role="grid"
        aria-label="Parsed table data"
      >
        <table id="ep-parsed-table">
          <thead>
            <tr>
              {tableData.selectedHeaders.map((label, i) => (
                <th key={selectedColumnIndexes[i]} scope="col">
                  {label}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {tableData.selectedRows.length === 0 ? (
              <tr>
                <td colSpan={Math.max(tableData.selectedHeaders.length, 1)}>
                  Header row found, but no non-empty data rows follow it.
                </td>
              </tr>
            ) : (
              tableData.selectedRows.map((cells, rowIndex) => (
                <tr key={rowIndex}>
                  {cells.map((cell) => {
                    const cls = [
                      !cell.value && 'cell-empty',
                      cell.anchor && 'cell-merged-anchor',
                    ]
                      .filter(Boolean)
                      .join(' ');

                    const tdProps = { className: cls || undefined };
                    if (cell.anchor?.rowSpan > 1)
                      tdProps.rowSpan = cell.anchor.rowSpan;

                    return (
                      <td key={cell.colIdx} {...tdProps}>
                        {cell.value || '—'}
                      </td>
                    );
                  })}
                </tr>
              ))
            )}
          </tbody>
        </table>
      </div>
    </div>
  );
}

import { useMemo, useRef } from 'react';
import { useCellCopy } from '../hooks/useCellCopy.js';

function buildParsedMeta({ records, columns, headerIndex, mergeSummary }) {
  return (
    'Using row ' +
    (headerIndex + 1) +
    ' as headers. ' +
    records +
    ' record' +
    (records === 1 ? '' : 's') +
    ' across ' +
    columns +
    ' column' +
    (columns === 1 ? '' : 's') +
    '. ' +
    mergeSummary
  );
}

export default function ParsedTable({
  isVisible,
  parsedTable,
  selectedColumnIndexes,
  headerIndex,
  mergeSummary,
  showStatus,
}) {
  const { handleTableClick } = useCellCopy({ showStatus });
  const copyContextRef = useRef(null);

  const data = useMemo(() => {
    const { headers, records, mergedCarryRows, mergedAnchors } = parsedTable;
    if (!headers.length) {
      return {
        selectedHeaders: [],
        selectedRows: [],
      };
    }

    const selectedHeaders = selectedColumnIndexes.map((idx) => headers[idx]);

    const selectedRows = records.map((record, recordIndex) => {
      const mergedFlags = mergedCarryRows[recordIndex] || [];
      const anchorFlags = mergedAnchors[recordIndex] || [];

      const cells = [];
      selectedColumnIndexes.forEach((columnIndex) => {
        const isCarry = Boolean(mergedFlags[columnIndex]);
        const anchor = anchorFlags[columnIndex] || null;
        if (isCarry && !anchor) return;

        const value = record[columnIndex];
        cells.push({
          columnIndex,
          value,
          anchor,
        });
      });
      return cells;
    });

    return { selectedHeaders, selectedRows };
  }, [parsedTable, selectedColumnIndexes]);

  copyContextRef.current = useMemo(() => {
    if (!data.selectedHeaders.length) return null;
    return {
      headers: data.selectedHeaders.slice(),
      records: parsedTable.records.map((record) =>
        selectedColumnIndexes.map((idx) => record[idx])
      ),
    };
  }, [data, parsedTable.records, selectedColumnIndexes]);

  if (!isVisible) return null;

  const hasHeaders = parsedTable.headers.length > 0;
  const records = parsedTable.records.length;
  const meta = !hasHeaders
    ? 'No columns could be derived from the chosen header row. ' + mergeSummary
    : buildParsedMeta({
        records,
        columns: parsedTable.headers.length,
        headerIndex,
        mergeSummary,
      });

  if (!hasHeaders) {
    return (
      <section id="parsed-section" className="panel data-panel">
        <div className="panel-heading">
          <div>
            <p className="section-label">Parsed output</p>
            <h2>Table data using the chosen header row</h2>
          </div>
          <p id="parsed-meta" className="panel-meta">
            {meta}
          </p>
        </div>
        <div id="parsed-empty" className="empty-state">
          Choose a different header row or verify that the sheet contains data.
        </div>
        <div id="parsed-wrap" className="table-shell is-hidden">
          <table id="parsed-table"></table>
        </div>
      </section>
    );
  }

  return (
    <section id="parsed-section" className="panel data-panel">
      <div className="panel-heading">
        <div>
          <p className="section-label">Parsed output</p>
          <h2>Table data using the chosen header row</h2>
        </div>
        <p id="parsed-meta" className="panel-meta">
          {meta}
        </p>
      </div>
      <div id="parsed-empty" className="empty-state is-hidden">
        Parsed table data will appear here once a sheet is loaded.
      </div>
      <div
        id="parsed-wrap"
        className="table-shell"
        onClick={(event) => handleTableClick(event, copyContextRef.current)}
      >
        <table id="parsed-table">
          <thead>
            <tr>
              {data.selectedHeaders.map((label, index) => (
                <th key={selectedColumnIndexes[index]}>{label}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {!data.selectedRows.length ? (
              <tr>
                <td colSpan={Math.max(data.selectedHeaders.length, 1)}>
                  Header row selected, but there are no non-empty data rows below it.
                </td>
              </tr>
            ) : (
              data.selectedRows.map((cells, rowIndex) => (
                <tr key={rowIndex}>
                  {cells.map((cell) => {
                    const classes = [];
                    if (!cell.value) classes.push('cell-empty');
                    if (cell.anchor) classes.push('cell-merged-anchor');
                    const tdProps = {
                      className: classes.join(' ') || undefined,
                    };
                    if (cell.anchor && cell.anchor.rowSpan > 1) {
                      tdProps.rowSpan = cell.anchor.rowSpan;
                    }
                    return (
                      <td key={cell.columnIndex} {...tdProps}>
                        {cell.value || '-'}
                      </td>
                    );
                  })}
                </tr>
              ))
            )}
          </tbody>
        </table>
      </div>
    </section>
  );
}

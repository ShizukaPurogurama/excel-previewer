import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { useCellCopy } from '../hooks/useCellCopy.js';
import { useColumnFilters } from '../hooks/useColumnFilters.js';
import { getHeaderSignature } from '../utils/columns.js';
import ColumnFilterMenu from './ColumnFilterMenu.jsx';

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

function FilterIcon({ active, open }) {
  const cls = ['ep-col-filter-icon', active && 'is-active', open && 'is-open']
    .filter(Boolean)
    .join(' ');

  return (
    <svg
      className={cls}
      viewBox="0 0 16 16"
      aria-hidden="true"
      focusable="false"
    >
      <path d="M2 3h12L9.5 8v4.4l-3 1.6V8L2 3z" />
    </svg>
  );
}

export default function DataTable({
  isVisible,
  parsedTable,
  selectedColumnIndexes,
  headerIndex,
  mergeSummary,
  showStatus,
  sheetName,
}) {
  const { handleTableClick } = useCellCopy({ showStatus });
  const copyContextRef = useRef(null);
  const filterButtonRefs = useRef({});
  const [openColumnKey, setOpenColumnKey] = useState(null);

  const { headers, records, mergedCarryRows, mergedAnchors } = parsedTable;
  const hasHeaders = headers.length > 0;

  const tableColumns = useMemo(
    () =>
      selectedColumnIndexes.map((columnIndex) => ({
        key: String(columnIndex),
        index: columnIndex,
        label: headers[columnIndex] || 'Column ' + (columnIndex + 1),
      })),
    [headers, selectedColumnIndexes]
  );

  const scopeKey = useMemo(
    () =>
      [sheetName || '', getHeaderSignature(headers), selectedColumnIndexes.join(',')].join(
        '::'
      ),
    [headers, sheetName, selectedColumnIndexes]
  );

  useEffect(() => {
    setOpenColumnKey(null);
  }, [scopeKey]);

  const setFilterButtonRef = useCallback(
    (columnKey) => (element) => {
      if (element) {
        filterButtonRefs.current[columnKey] = element;
        return;
      }
      delete filterButtonRefs.current[columnKey];
    },
    []
  );

  const {
    columns: filteredColumns,
    filteredRowIndexes,
    activeFilterCount,
    setColumnFilter,
    clearColumnFilter,
    clearAllFilters,
  } = useColumnFilters({
    rows: records,
    columns: tableColumns,
    scopeKey,
  });

  const tableData = useMemo(() => {
    if (!hasHeaders) return { selectedHeaders: [], selectedRows: [] };

    const selectedHeaders = filteredColumns.map((column) => column.label);

    const selectedRows = filteredRowIndexes.map((rowIndex) => {
      const record = records[rowIndex] || [];
      const flags = mergedCarryRows[rowIndex] || [];
      const anchors = mergedAnchors[rowIndex] || [];
      const cells = [];

      for (const column of tableColumns) {
        const colIdx = column.index;
        const isCarry = Boolean(flags[colIdx]);
        const anchor = anchors[colIdx] || null;
        if (isCarry && !anchor) continue;
        cells.push({ colIdx, value: record[colIdx], anchor });
      }

      return {
        rowIndex,
        cells,
        record: selectedColumnIndexes.map((idx) => record[idx]),
      };
    });

    return { selectedHeaders, selectedRows };
  }, [
    filteredColumns,
    filteredRowIndexes,
    hasHeaders,
    mergedAnchors,
    mergedCarryRows,
    records,
    selectedColumnIndexes,
    tableColumns,
  ]);

  copyContextRef.current = useMemo(() => {
    if (!tableData.selectedHeaders.length) return null;
    return {
      headers: tableData.selectedHeaders.slice(),
      records: tableData.selectedRows.map((row) => row.record.slice()),
    };
  }, [tableData.selectedHeaders, tableData.selectedRows]);

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

  const metaText = buildMeta(records.length, headers.length, headerIndex, mergeSummary);
  const filteredMetaText =
    activeFilterCount > 0
      ? ' · Showing ' + filteredRowIndexes.length + ' of ' + records.length + ' rows'
      : '';

  const handleToggleColumnMenu = (columnKey) => {
    setOpenColumnKey((current) => (current === columnKey ? null : columnKey));
  };

  const handleClearAllFilters = () => {
    clearAllFilters();
    setOpenColumnKey(null);
  };

  return (
    <div>
      <div className="ep-data-meta-row">
        <p className="ep-data-meta">{metaText + filteredMetaText}</p>
        {activeFilterCount > 0 && (
          <button
            type="button"
            className="ep-clear-filters-btn"
            onClick={handleClearAllFilters}
          >
            Clear filters
          </button>
        )}
      </div>
      <div
        className="ep-table-shell"
        onClick={(e) => handleTableClick(e, copyContextRef.current)}
        role="grid"
        aria-label="Parsed table data"
      >
        <table id="ep-parsed-table">
          <thead>
            <tr>
              {filteredColumns.map((column) => {
                const isOpen = openColumnKey === column.key;
                const triggerRef = setFilterButtonRef(column.key);

                return (
                  <th
                    key={column.key}
                    scope="col"
                    className={
                      'ep-col-th' +
                      (isOpen ? ' ep-col-th-open' : '') +
                      (column.isActive ? ' ep-col-th-active' : '')
                    }
                  >
                    <div className="ep-col-head">
                      <span className="ep-col-label" title={column.label}>
                        {column.label}
                      </span>
                      <button
                        ref={triggerRef}
                        type="button"
                        className={
                          'ep-col-filter-trigger' +
                          (column.isActive ? ' is-active' : '') +
                          (isOpen ? ' is-open' : '')
                        }
                        data-ep-column-filter-trigger
                        aria-label={`Filter ${column.label}`}
                        aria-expanded={isOpen ? 'true' : 'false'}
                        onClick={(event) => {
                          event.stopPropagation();
                          handleToggleColumnMenu(column.key);
                        }}
                      >
                        <FilterIcon active={column.isActive} open={isOpen} />
                        <span className="ep-col-filter-caret" aria-hidden="true">
                          ▾
                        </span>
                      </button>
                    </div>

                    {isOpen && (
                      <ColumnFilterMenu
                        anchorElement={filterButtonRefs.current[column.key] || null}
                        columnLabel={column.label}
                        state={column.state}
                        options={column.options}
                        onChange={(patch) => setColumnFilter(column.key, patch)}
                        onClear={() => clearColumnFilter(column.key)}
                        onClose={() => setOpenColumnKey(null)}
                      />
                    )}
                  </th>
                );
              })}
            </tr>
          </thead>
          <tbody>
            {tableData.selectedRows.length === 0 ? (
              <tr>
                <td colSpan={Math.max(tableData.selectedHeaders.length, 1)}>
                  {activeFilterCount > 0
                    ? 'No rows match the active filters.'
                    : 'Header row found, but no non-empty data rows follow it.'}
                </td>
              </tr>
            ) : (
              tableData.selectedRows.map(({ rowIndex, cells }) => (
                <tr key={rowIndex}>
                  {cells.map((cell) => {
                    const cls = [
                      !cell.value && 'cell-empty',
                      cell.anchor && 'cell-merged-anchor',
                    ]
                      .filter(Boolean)
                      .join(' ');

                    const tdProps = { className: cls || undefined };
                    if (cell.anchor?.rowSpan > 1) tdProps.rowSpan = cell.anchor.rowSpan;

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
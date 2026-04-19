import { useCallback, useMemo, useState } from 'react';
import {
  normalizeCell,
  normalizeColumnPickerSearch,
} from '../utils/columns.js';

const DEFAULT_SORT_DIRECTION = 'asc';
const DEFAULT_TEXT_MODE = 'contains';
const EMPTY_SCOPE_STATE = Object.freeze({});
const VALUE_COLLATOR = new Intl.Collator(undefined, {
  numeric: true,
  sensitivity: 'base',
});

function normalizeColumnKey(columnKey) {
  return String(columnKey);
}

function compareTokenValues(left, right) {
  const leftBlank = left === '';
  const rightBlank = right === '';
  if (leftBlank !== rightBlank) return leftBlank ? 1 : -1;
  return VALUE_COLLATOR.compare(left, right);
}

function sortTokenValues(values) {
  return Array.from(new Set(values)).sort(compareTokenValues);
}

function resolveColumnState(rawState) {
  return {
    selectedValues: Array.isArray(rawState?.selectedValues)
      ? sortTokenValues(rawState.selectedValues.map((value) => normalizeCell(value)))
      : null,
    sortDirection: rawState?.sortDirection === 'desc' ? 'desc' : DEFAULT_SORT_DIRECTION,
    textMode: rawState?.textMode === 'startsWith' ? 'startsWith' : DEFAULT_TEXT_MODE,
    textQuery: String(rawState?.textQuery || '').trim(),
  };
}

function buildColumnOptions(rows, columns) {
  const optionMaps = columns.map(() => new Map());

  rows.forEach((row) => {
    columns.forEach((column, columnPosition) => {
      const columnIndex = Number.isInteger(column.index)
        ? column.index
        : Number.parseInt(String(column.key), 10);
      const value = normalizeCell(row && row[columnIndex]);
      const label = value || '(Blanks)';
      const optionMap = optionMaps[columnPosition];
      const existing = optionMap.get(value);

      if (existing) {
        existing.count += 1;
        return;
      }

      optionMap.set(value, {
        value,
        label,
        count: 1,
        searchText: normalizeColumnPickerSearch(label),
      });
    });
  });

  return columns.map((column, columnPosition) => {
    const columnIndex = Number.isInteger(column.index)
      ? column.index
      : Number.parseInt(String(column.key), 10);
    const options = Array.from(optionMaps[columnPosition].values());

    return {
      key: normalizeColumnKey(column.key),
      index: columnIndex,
      label: column.label || 'Column ' + (columnPosition + 1),
      options,
    };
  });
}

function isDefaultColumnState(state) {
  return (
    state.selectedValues === null &&
    state.sortDirection === DEFAULT_SORT_DIRECTION &&
    state.textMode === DEFAULT_TEXT_MODE &&
    !String(state.textQuery || '').trim()
  );
}

function normalizePatchState(currentState, patch) {
  const base = currentState || resolveColumnState();
  const nextSelectedValues =
    patch && Object.prototype.hasOwnProperty.call(patch, 'selectedValues')
      ? patch.selectedValues === null
        ? null
        : Array.isArray(patch.selectedValues)
          ? sortTokenValues(
              patch.selectedValues.map((value) => normalizeCell(value))
            )
          : base.selectedValues
      : base.selectedValues;

  return {
    selectedValues: nextSelectedValues,
    sortDirection:
      patch && patch.sortDirection === 'desc'
        ? 'desc'
        : patch && patch.sortDirection === 'asc'
          ? 'asc'
          : base.sortDirection,
    textMode:
      patch && patch.textMode === 'startsWith'
        ? 'startsWith'
        : patch && patch.textMode === 'contains'
          ? 'contains'
          : base.textMode,
    textQuery:
      patch && Object.prototype.hasOwnProperty.call(patch, 'textQuery')
        ? String(patch.textQuery || '').trim()
        : base.textQuery,
  };
}

export function useColumnFilters({ rows, columns, scopeKey }) {
  const [scopeStateByKey, setScopeStateByKey] = useState(() => new Map());

  const columnMeta = useMemo(() => {
    if (!Array.isArray(columns) || !columns.length) return [];
    return buildColumnOptions(Array.isArray(rows) ? rows : [], columns);
  }, [rows, columns]);

  const currentScopeState = scopeStateByKey.get(scopeKey) || EMPTY_SCOPE_STATE;

  const resolvedColumnMeta = useMemo(() => {
    if (!columnMeta.length) return [];
    const stateByKey = currentScopeState || {};

    return columnMeta.map((column) => {
      const rawState = stateByKey[column.key] || null;
      const state = resolveColumnState(rawState);
      const active = Boolean(state.selectedValues !== null || state.textQuery);
      const options = column.options.slice().sort((left, right) => {
        if (left.value === '' && right.value !== '') return 1;
        if (right.value === '' && left.value !== '') return -1;
        if (state.sortDirection === 'desc') {
          return VALUE_COLLATOR.compare(right.label, left.label);
        }
        return VALUE_COLLATOR.compare(left.label, right.label);
      });

      return {
        ...column,
        options,
        state,
        isActive: active,
        selectedSet:
          state.selectedValues === null
            ? new Set(options.map((option) => option.value))
            : new Set(state.selectedValues),
        normalizedTextQuery: normalizeColumnPickerSearch(state.textQuery),
      };
    });
  }, [columnMeta, currentScopeState]);

  const filteredRowIndexes = useMemo(() => {
    if (!Array.isArray(rows) || !rows.length) return [];

    const activeColumns = resolvedColumnMeta.filter((column) => column.isActive);
    if (!activeColumns.length) {
      return Array.from({ length: rows.length }, (_, index) => index);
    }

    const visibleIndexes = [];

    rowLoop: for (let rowIndex = 0; rowIndex < rows.length; rowIndex += 1) {
      const row = rows[rowIndex] || [];

      for (const column of activeColumns) {
        const cellValue = normalizeCell(row[column.index]);

        if (
          column.state.selectedValues !== null &&
          !column.selectedSet.has(cellValue)
        ) {
          continue rowLoop;
        }

        if (column.normalizedTextQuery) {
          const haystack = normalizeColumnPickerSearch(cellValue);
          if (
            column.state.textMode === 'startsWith'
              ? !haystack.startsWith(column.normalizedTextQuery)
              : !haystack.includes(column.normalizedTextQuery)
          ) {
            continue rowLoop;
          }
        }
      }

      visibleIndexes.push(rowIndex);
    }

    return visibleIndexes;
  }, [rows, resolvedColumnMeta]);

  const activeFilterCount = useMemo(
    () => resolvedColumnMeta.filter((column) => column.isActive).length,
    [resolvedColumnMeta]
  );

  const setColumnFilter = useCallback(
    (columnKey, patch) => {
      if (!scopeKey) return;
      const normalizedKey = normalizeColumnKey(columnKey);

      setScopeStateByKey((previous) => {
        const next = new Map(previous);
        const scopeState = { ...(next.get(scopeKey) || {}) };
        const currentState = resolveColumnState(scopeState[normalizedKey]);
        const patchState = typeof patch === 'function' ? patch(currentState) : patch;
        const nextState = normalizePatchState(currentState, patchState || {});

        if (isDefaultColumnState(nextState)) {
          delete scopeState[normalizedKey];
        } else {
          scopeState[normalizedKey] = nextState;
        }

        if (Object.keys(scopeState).length) {
          next.set(scopeKey, scopeState);
        } else {
          next.delete(scopeKey);
        }

        return next;
      });
    },
    [scopeKey]
  );

  const clearColumnFilter = useCallback(
    (columnKey) => {
      if (!scopeKey) return;
      const normalizedKey = normalizeColumnKey(columnKey);

      setScopeStateByKey((previous) => {
        const next = new Map(previous);
        const scopeState = { ...(next.get(scopeKey) || {}) };
        if (!scopeState[normalizedKey]) return previous;
        delete scopeState[normalizedKey];

        if (Object.keys(scopeState).length) {
          next.set(scopeKey, scopeState);
        } else {
          next.delete(scopeKey);
        }

        return next;
      });
    },
    [scopeKey]
  );

  const clearAllFilters = useCallback(() => {
    if (!scopeKey) return;
    setScopeStateByKey((previous) => {
      if (!previous.has(scopeKey)) return previous;
      const next = new Map(previous);
      next.delete(scopeKey);
      return next;
    });
  }, [scopeKey]);

  return {
    columns: resolvedColumnMeta,
    filteredRowIndexes,
    activeFilterCount,
    setColumnFilter,
    clearColumnFilter,
    clearAllFilters,
  };
}
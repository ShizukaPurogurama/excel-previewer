import { useEffect, useId, useLayoutEffect, useMemo, useRef, useState } from 'react';
import { normalizeColumnPickerSearch } from '../utils/columns.js';

const MENU_WIDTH = 320;

function sortTokens(values) {
  return Array.from(new Set(values)).sort((left, right) => {
    const leftBlank = left === '';
    const rightBlank = right === '';
    if (leftBlank !== rightBlank) return leftBlank ? 1 : -1;
    return left.localeCompare(right, undefined, {
      numeric: true,
      sensitivity: 'base',
    });
  });
}

function getSelectedSet(state, options) {
  if (state.selectedValues === null) {
    return new Set(options.map((option) => option.value));
  }
  return new Set(state.selectedValues);
}

export default function ColumnFilterMenu({
  anchorElement,
  columnLabel,
  state,
  options,
  onChange,
  onClear,
  onClose,
}) {
  const menuRef = useRef(null);
  const selectAllRef = useRef(null);
  const searchInputId = useId();
  const [searchValue, setSearchValue] = useState('');
  const [alignment, setAlignment] = useState('start');

  const normalizedSearch = normalizeColumnPickerSearch(searchValue);

  const visibleOptions = useMemo(() => {
    if (!normalizedSearch) return options;
    return options.filter((option) => option.searchText.includes(normalizedSearch));
  }, [normalizedSearch, options]);

  const selectedSet = useMemo(() => getSelectedSet(state, options), [state, options]);

  const visibleSelectedCount = useMemo(
    () =>
      visibleOptions.reduce(
        (count, option) => count + (selectedSet.has(option.value) ? 1 : 0),
        0
      ),
    [selectedSet, visibleOptions]
  );

  const allVisibleSelected =
    visibleOptions.length > 0 && visibleOptions.every((option) => selectedSet.has(option.value));
  const someVisibleSelected =
    visibleOptions.length > 0 && visibleOptions.some((option) => selectedSet.has(option.value));
  const selectAllIndeterminate = someVisibleSelected && !allVisibleSelected;
  const isFilterActive = Boolean(state.selectedValues !== null || state.textQuery);

  useEffect(() => {
    if (selectAllRef.current) {
      selectAllRef.current.indeterminate = selectAllIndeterminate;
    }
  }, [selectAllIndeterminate]);

  useLayoutEffect(() => {
    if (!anchorElement || !menuRef.current) return undefined;

    const measureAlignment = () => {
      const anchorRect = anchorElement.getBoundingClientRect();
      const viewportWidth = window.innerWidth || MENU_WIDTH;
      const estimatedMenuWidth = Math.min(MENU_WIDTH, viewportWidth - 20);
      const shouldAlignEnd =
        anchorRect.right + estimatedMenuWidth > viewportWidth - 12 &&
        anchorRect.left > estimatedMenuWidth / 2;

      setAlignment(shouldAlignEnd ? 'end' : 'start');
    };

    const frame = window.requestAnimationFrame(measureAlignment);
    window.addEventListener('resize', measureAlignment);

    return () => {
      window.cancelAnimationFrame(frame);
      window.removeEventListener('resize', measureAlignment);
    };
  }, [anchorElement, options.length, state.sortDirection, state.textQuery]);

  useEffect(() => {
    const handlePointerDown = (event) => {
      const target = event.target;
      if (!(target instanceof Node)) return;
      if (menuRef.current && menuRef.current.contains(target)) return;
      if (anchorElement && anchorElement.contains(target)) return;
      onClose();
    };

    const handleKeyDown = (event) => {
      if (event.key !== 'Escape') return;
      event.preventDefault();
      if (searchValue) {
        setSearchValue('');
        return;
      }
      onClose();
    };

    document.addEventListener('pointerdown', handlePointerDown);
    document.addEventListener('keydown', handleKeyDown);

    return () => {
      document.removeEventListener('pointerdown', handlePointerDown);
      document.removeEventListener('keydown', handleKeyDown);
    };
  }, [anchorElement, onClose, searchValue]);

  useEffect(() => {
    if (!menuRef.current) return;
    const input = menuRef.current.querySelector('input[type="search"]');
    if (input && document.activeElement !== input) {
      input.focus();
      input.select();
    }
  }, []);

  const applySelectedValues = (nextValues) => {
    const uniqueValues = sortTokens(nextValues);
    const allValuesSelected =
      options.length > 0 &&
      uniqueValues.length === options.length &&
      options.every((option) => uniqueValues.indexOf(option.value) !== -1);

    if (allValuesSelected && !String(state.textQuery || '').trim()) {
      onChange({ selectedValues: null });
      return;
    }

    onChange({ selectedValues: uniqueValues });
  };

  const handleToggleAllVisible = () => {
    const nextValues = new Set(selectedSet);
    if (allVisibleSelected) {
      visibleOptions.forEach((option) => nextValues.delete(option.value));
    } else {
      visibleOptions.forEach((option) => nextValues.add(option.value));
    }
    applySelectedValues(Array.from(nextValues));
  };

  const handleToggleValue = (value, checked) => {
    const nextValues = new Set(selectedSet);
    if (checked) nextValues.add(value);
    else nextValues.delete(value);
    applySelectedValues(Array.from(nextValues));
  };

  const handleSortChange = (sortDirection) => {
    if (sortDirection === state.sortDirection) return;
    onChange({ sortDirection });
  };

  const handleTextModeChange = (nextMode) => {
    if (nextMode === state.textMode) return;
    onChange({ textMode: nextMode });
  };

  const handleTextQueryChange = (nextQuery) => {
    onChange({ textQuery: nextQuery });
  };

  const menuCls = [
    'ep-col-filter-menu',
    alignment === 'end' && 'is-end',
    alignment !== 'end' && 'is-start',
  ]
    .filter(Boolean)
    .join(' ');

  const sortAscActive = state.sortDirection === 'asc';
  const sortDescActive = state.sortDirection === 'desc';

  return (
    <div
      ref={menuRef}
      className={menuCls}
      data-ep-column-filter
      role="dialog"
      aria-modal="false"
      aria-label={`Filter ${columnLabel}`}
    >
      <div className="ep-col-filter-header">
        <div className="ep-col-filter-heading">
          <p className="ep-col-filter-title">{columnLabel}</p>
          <p className="ep-col-filter-subtitle">{options.length} unique values</p>
        </div>
        <div className="ep-col-filter-header-actions">
          <button
            type="button"
            className="ep-col-filter-link"
            disabled={!isFilterActive}
            onClick={() => {
              setSearchValue('');
              onClear();
            }}
          >
            Clear
          </button>
          <button
            type="button"
            className="ep-col-filter-link"
            onClick={onClose}
          >
            Close
          </button>
        </div>
      </div>

      <div className="ep-col-filter-sorts" aria-label="Sort column values">
        <button
          type="button"
          className={'ep-col-filter-sort' + (sortAscActive ? ' is-active' : '')}
          aria-pressed={sortAscActive ? 'true' : 'false'}
          onClick={() => handleSortChange('asc')}
        >
          Sort A to Z
        </button>
        <button
          type="button"
          className={'ep-col-filter-sort' + (sortDescActive ? ' is-active' : '')}
          aria-pressed={sortDescActive ? 'true' : 'false'}
          onClick={() => handleSortChange('desc')}
        >
          Sort Z to A
        </button>
      </div>

      <label className="ep-col-filter-search" htmlFor={searchInputId}>
        <span className="ep-col-filter-label">Search values</span>
        <input
          id={searchInputId}
          type="search"
          className="ep-col-filter-input"
          placeholder="Search values"
          autoComplete="off"
          spellCheck="false"
          value={searchValue}
          onChange={(event) => setSearchValue(event.target.value)}
        />
      </label>

      <label className="ep-col-filter-select-all">
        <input
          ref={selectAllRef}
          type="checkbox"
          checked={allVisibleSelected}
          disabled={!visibleOptions.length}
          onChange={handleToggleAllVisible}
        />
        <span>Select all</span>
      </label>

      <div className="ep-col-filter-list" data-ep-column-filter>
        {!visibleOptions.length ? (
          <div className="ep-col-filter-empty">No values match your search.</div>
        ) : (
          visibleOptions.map((option) => {
            const checked = selectedSet.has(option.value);

            return (
              <label
                key={option.value || '(blank)'}
                className={
                  'ep-col-filter-option' + (checked ? ' is-selected' : '')
                }
              >
                <input
                  type="checkbox"
                  checked={checked}
                  onChange={(event) => handleToggleValue(option.value, event.target.checked)}
                />
                <span className="ep-col-filter-option-label">{option.label}</span>
                <span className="ep-col-filter-option-count">{option.count}</span>
              </label>
            );
          })
        )}
      </div>

      <div className="ep-col-filter-text">
        <div className="ep-col-filter-text-head">
          <span className="ep-col-filter-label">Text filter</span>
        </div>
        <div className="ep-col-filter-text-row">
          <select
            className="ep-col-filter-select"
            value={state.textMode}
            onChange={(event) => handleTextModeChange(event.target.value)}
          >
            <option value="contains">Contains</option>
            <option value="startsWith">Starts with</option>
          </select>
          <input
            type="search"
            className="ep-col-filter-input"
            placeholder="Text to match"
            autoComplete="off"
            spellCheck="false"
            value={state.textQuery}
            onChange={(event) => handleTextQueryChange(event.target.value)}
          />
        </div>
      </div>

      <div className="ep-col-filter-footer">
        <span>
          {visibleSelectedCount}/{visibleOptions.length || 0} visible selected
        </span>
      </div>
    </div>
  );
}
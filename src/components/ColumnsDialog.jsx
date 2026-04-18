import { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react';
import { MAX_VISIBLE_COLUMNS } from '../constants.js';
import { normalizeColumnPickerSearch, toSpreadsheetColumn } from '../utils/columns.js';
import { hasAnyMoneyMapping } from '../utils/money.js';
import { useBodyScrollLock } from '../hooks/useBodyScrollLock.js';

function parseSelectIndex(value) {
  if (value === '' || value === null || value === undefined) return null;
  const parsed = Number.parseInt(String(value), 10);
  return Number.isInteger(parsed) && parsed >= 0 ? parsed : null;
}

function readMapping(form) {
  return {
    senderCurrency: parseSelectIndex(form.senderCurrency),
    senderAmount: parseSelectIndex(form.senderAmount),
    receiverCurrency: parseSelectIndex(form.receiverCurrency),
    receiverAmount: parseSelectIndex(form.receiverAmount),
    source: 'manual',
  };
}

function extractMappingForm(mapping) {
  return {
    senderCurrency: Number.isInteger(mapping && mapping.senderCurrency)
      ? String(mapping.senderCurrency)
      : '',
    senderAmount: Number.isInteger(mapping && mapping.senderAmount)
      ? String(mapping.senderAmount)
      : '',
    receiverCurrency: Number.isInteger(mapping && mapping.receiverCurrency)
      ? String(mapping.receiverCurrency)
      : '',
    receiverAmount: Number.isInteger(mapping && mapping.receiverAmount)
      ? String(mapping.receiverAmount)
      : '',
  };
}

function resolveMoneyNote({ headers, moneyMappingState }) {
  const countText = headers.length
    ? headers.length + ' columns available'
    : 'No columns available';

  if (!headers.length) {
    return 'Map sender and receiver amount columns to format money values.';
  }

  const source = moneyMappingState && moneyMappingState.source;
  const mapping = moneyMappingState && moneyMappingState.mapping;

  if (source === 'manual') {
    return 'Using your saved money mapping for this column set. ' + countText + '.';
  }
  if (mapping && (mapping.senderAmount !== null || mapping.receiverAmount !== null)) {
    return 'Auto-detected money columns from the sheet. Adjust them if needed. ' + countText + '.';
  }
  return 'Map sender and receiver amount columns to format money values. ' + countText + '.';
}

export default function ColumnsDialog({
  isOpen,
  onClose,
  hasWorkbook,
  headers,
  selectedColumnIndexes,
  moneyMappingState,
  onToggleColumn,
  onResetColumns,
  onApplyMoneyMapping,
  onAutoDetectMoney,
  onClearMoneyMapping,
}) {
  const dialogRef = useRef(null);
  const optionsRef = useRef(null);
  const searchInputRef = useRef(null);
  const scrollTopRef = useRef(0);
  const restoreScrollRef = useRef(false);
  const focusValueRef = useRef('');
  const restoreFocusRef = useRef(false);

  const [search, setSearch] = useState('');
  const [note, setNote] = useState({ message: '', tone: '' });
  const [mappingForm, setMappingForm] = useState(() =>
    extractMappingForm(moneyMappingState && moneyMappingState.mapping)
  );

  useBodyScrollLock(isOpen);

  useEffect(() => {
    setMappingForm(extractMappingForm(moneyMappingState && moneyMappingState.mapping));
  }, [moneyMappingState]);

  useEffect(() => {
    const dialog = dialogRef.current;
    if (!dialog) return;

    if (isOpen && !dialog.open) {
      setSearch('');
      setNote({
        message:
          'Search the list below and select up to ' + MAX_VISIBLE_COLUMNS + ' columns.',
        tone: '',
      });
      scrollTopRef.current = 0;
      restoreScrollRef.current = false;
      focusValueRef.current = '';
      restoreFocusRef.current = false;

      if (typeof dialog.showModal === 'function') {
        dialog.showModal();
      }
      window.requestAnimationFrame(() => {
        if (searchInputRef.current) {
          searchInputRef.current.focus();
          searchInputRef.current.select();
        }
      });
    } else if (!isOpen && dialog.open) {
      dialog.close();
    }
  }, [isOpen]);

  useEffect(() => {
    const dialog = dialogRef.current;
    if (!dialog) return undefined;
    const handleClose = () => {
      onClose();
    };
    dialog.addEventListener('close', handleClose);
    return () => dialog.removeEventListener('close', handleClose);
  }, [onClose]);

  const visibleOptions = useMemo(() => {
    const selectedSet = new Set(selectedColumnIndexes);
    const query = normalizeColumnPickerSearch(search);

    const items = [];
    headers.forEach((label, index) => {
      const key = toSpreadsheetColumn(index);
      const haystack = (key + ' ' + String(label || '')).toLowerCase();
      if (query && haystack.indexOf(query) === -1) return;

      items.push({
        index,
        key,
        label: label || 'Column ' + (index + 1),
        selected: selectedSet.has(index),
      });
    });

    items.sort((a, b) => {
      if (a.selected !== b.selected) return a.selected ? -1 : 1;
      return a.index - b.index;
    });

    return items;
  }, [headers, selectedColumnIndexes, search]);

  const visibleSelectedCount = visibleOptions.filter((item) => item.selected).length;

  useLayoutEffect(() => {
    const optionsEl = optionsRef.current;
    if (!optionsEl) return;

    if (restoreScrollRef.current) {
      optionsEl.scrollTop = scrollTopRef.current;
      restoreScrollRef.current = false;
    } else {
      optionsEl.scrollTop = 0;
    }

    if (restoreFocusRef.current && focusValueRef.current) {
      const checkbox = optionsEl.querySelector(
        'input[type="checkbox"][value="' + focusValueRef.current + '"]'
      );
      if (checkbox && typeof checkbox.focus === 'function') {
        try {
          checkbox.focus({ preventScroll: true });
        } catch (error) {
          checkbox.focus();
        }
      }
      restoreFocusRef.current = false;
    }
  }, [visibleOptions]);

  const handleSearchChange = (event) => {
    setSearch(event.target.value);
    scrollTopRef.current = 0;
    restoreScrollRef.current = false;
  };

  const handleSearchKeydown = (event) => {
    if (event.key !== 'Escape' || !event.target.value) return;
    event.preventDefault();
    event.stopPropagation();
    setSearch('');
    scrollTopRef.current = 0;
    restoreScrollRef.current = false;
  };

  const handleOptionsScroll = useCallback(() => {
    if (optionsRef.current) {
      scrollTopRef.current = optionsRef.current.scrollTop;
    }
  }, []);

  const handleToggle = (toggledIndex, nextChecked) => {
    scrollTopRef.current = optionsRef.current ? optionsRef.current.scrollTop : 0;
    restoreScrollRef.current = true;
    focusValueRef.current = String(toggledIndex);
    restoreFocusRef.current = true;

    const result = onToggleColumn(toggledIndex, nextChecked);

    if (!result || !result.ok) {
      if (result && result.reason === 'max-reached') {
        setNote({
          message: 'Maximum is ' + MAX_VISIBLE_COLUMNS + ' columns.',
          tone: 'error',
        });
      } else if (result && result.reason === 'min-reached') {
        setNote({ message: 'At least one column must stay selected.', tone: 'error' });
      }
      return;
    }

    setNote({
      message: 'Applied ' + result.nextSelection.length + ' selected columns.',
      tone: 'success',
    });
  };

  const handleReset = () => {
    const defaults = onResetColumns();
    if (!defaults || !defaults.indexes) return;
    setNote({
      message:
        defaults.source === 'preset'
          ? 'Reset to the preset columns for this sheet.'
          : 'Reset to the first ' + defaults.indexes.length + ' columns.',
      tone: 'success',
    });
  };

  const handleMappingFieldChange = (field) => (event) => {
    const nextForm = { ...mappingForm, [field]: event.target.value };
    setMappingForm(nextForm);

    const mapping = readMapping(nextForm);
    const applied = onApplyMoneyMapping(mapping);

    if (applied === true) {
      setNote({ message: 'Saved money mapping for this column set.', tone: 'success' });
    } else if (applied === false) {
      setNote({ message: 'Cleared the money mapping for this column set.', tone: '' });
    }
  };

  const handleAutoDetect = () => {
    onAutoDetectMoney();
    setNote({
      message: 'Applied auto-detected money mapping. Review it before continuing.',
      tone: 'success',
    });
  };

  const handleClear = () => {
    onClearMoneyMapping();
    setMappingForm({
      senderCurrency: '',
      senderAmount: '',
      receiverCurrency: '',
      receiverAmount: '',
    });
    setNote({
      message: 'Cleared saved money mapping. Auto-detection is active again.',
      tone: '',
    });
  };

  const mappingNoteText = resolveMoneyNote({ headers, moneyMappingState });

  const hasHeaders = headers.length > 0;
  const mappingValues = readMapping(mappingForm);
  const manualMappingExists = hasAnyMoneyMapping({
    ...mappingValues,
    source: 'manual',
  });

  const noteClasses = ['dialog-note'];
  if (note.tone === 'error') noteClasses.push('is-error');
  else if (note.tone === 'success') noteClasses.push('is-success');

  const searchMetaText = !hasHeaders
    ? 'No columns available'
    : visibleSelectedCount
      ? visibleOptions.length + ' of ' + headers.length + ' columns | ' + visibleSelectedCount + ' selected'
      : visibleOptions.length + ' of ' + headers.length + ' columns';

  const fieldOptions = headers.map((label, index) => ({
    value: String(index),
    label: toSpreadsheetColumn(index) + ' - ' + (label || 'Column ' + (index + 1)),
  }));

  return (
    <dialog id="columns-dialog" className="columns-dialog" ref={dialogRef}>
      <form method="dialog" className="columns-dialog-shell">
        <header className="columns-dialog-header">
          <div>
            <p className="section-label">Columns</p>
            <h2>Choose visible columns</h2>
          </div>
          <button id="close-columns-dialog" className="ghost-button" type="submit">
            Close
          </button>
        </header>

        <div className="columns-toolbar">
          <div className="columns-toolbar-copy">
            <span id="columns-dialog-count" className="column-count">
              {selectedColumnIndexes.length}/{MAX_VISIBLE_COLUMNS} selected
            </span>
            <span id="columns-search-meta" className="columns-search-meta">
              {searchMetaText}
            </span>
          </div>
          <button
            id="reset-columns-button"
            className="ghost-button"
            type="button"
            disabled={!hasWorkbook || !hasHeaders}
            onClick={handleReset}
          >
            Reset defaults
          </button>
        </div>

        <section className="columns-search-panel money-mapping-panel">
          <div className="money-mapping-header">
            <div>
              <p className="section-label">Money mapping</p>
              <h3>Map sender and receiver money columns</h3>
            </div>
            <div className="money-mapping-actions">
              <button
                id="auto-detect-money-button"
                className="ghost-button"
                type="button"
                disabled={!hasWorkbook || !hasHeaders}
                onClick={handleAutoDetect}
              >
                Auto-detect
              </button>
              <button
                id="clear-money-mapping-button"
                className="ghost-button"
                type="button"
                disabled={!hasWorkbook || !manualMappingExists}
                onClick={handleClear}
              >
                Clear
              </button>
            </div>
          </div>

          <p id="money-mapping-note" className="dialog-note">
            {mappingNoteText}
          </p>

          <div className="money-mapping-grid">
            {[
              { id: 'sender-currency-select', field: 'senderCurrency', label: 'Sender currency' },
              { id: 'sender-amount-select', field: 'senderAmount', label: 'Sender amount' },
              {
                id: 'receiver-currency-select',
                field: 'receiverCurrency',
                label: 'Receiver currency',
              },
              { id: 'receiver-amount-select', field: 'receiverAmount', label: 'Receiver amount' },
            ].map(({ id, field, label }) => (
              <label key={id} className="money-mapping-field" htmlFor={id}>
                <span>{label}</span>
                <select
                  id={id}
                  className="money-mapping-select"
                  disabled={!hasWorkbook || !hasHeaders}
                  value={mappingForm[field]}
                  onChange={handleMappingFieldChange(field)}
                >
                  <option value="">Not mapped</option>
                  {fieldOptions.map((option) => (
                    <option key={option.value} value={option.value}>
                      {option.label}
                    </option>
                  ))}
                </select>
              </label>
            ))}
          </div>
        </section>

        <p id="columns-dialog-note" className={noteClasses.join(' ')}>
          {note.message ||
            'Search the list below and select up to ' + MAX_VISIBLE_COLUMNS + ' columns for the parsed table.'}
        </p>

        <label className="columns-search-field" htmlFor="columns-search-input">
          <span className="columns-search-label">Search columns</span>
          <input
            id="columns-search-input"
            ref={searchInputRef}
            className="columns-search-input"
            type="search"
            placeholder="Search by column letter or name"
            autoComplete="off"
            spellCheck="false"
            value={search}
            onChange={handleSearchChange}
            onKeyDown={handleSearchKeydown}
          />
        </label>

        <div
          id="columns-options"
          ref={optionsRef}
          className="columns-options"
          data-view="list"
          onScroll={handleOptionsScroll}
        >
          {!visibleOptions.length ? (
            <div className="columns-options-empty">
              {search ? 'No columns match your search.' : 'No columns are available.'}
            </div>
          ) : (
            visibleOptions.map((item) => (
              <label
                key={item.index}
                className={'columns-option' + (item.selected ? ' is-selected' : '')}
              >
                <input
                  type="checkbox"
                  value={String(item.index)}
                  checked={item.selected}
                  disabled={!hasWorkbook}
                  onChange={(event) => handleToggle(item.index, event.target.checked)}
                />
                <span className="columns-option-key">{item.key}</span>
                <span className="columns-option-name">{item.label}</span>
              </label>
            ))
          )}
        </div>
      </form>
    </dialog>
  );
}

import { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react';
import { MAX_VISIBLE_COLUMNS } from '../../../constants.js';
import { normalizeColumnPickerSearch, toSpreadsheetColumn } from '../utils/columns.js';
import { hasAnyMoneyMapping } from '../utils/money.js';
import { useBodyScrollLock } from '../hooks/useBodyScrollLock.js';

function parseIdx(val) {
  if (val === '' || val == null) return null;
  const n = Number.parseInt(String(val), 10);
  return Number.isInteger(n) && n >= 0 ? n : null;
}

function extractForm(mapping) {
  return {
    senderCurrency: Number.isInteger(mapping?.senderCurrency)
      ? String(mapping.senderCurrency)
      : '',
    senderAmount: Number.isInteger(mapping?.senderAmount)
      ? String(mapping.senderAmount)
      : '',
    receiverCurrency: Number.isInteger(mapping?.receiverCurrency)
      ? String(mapping.receiverCurrency)
      : '',
    receiverAmount: Number.isInteger(mapping?.receiverAmount)
      ? String(mapping.receiverAmount)
      : '',
  };
}

function readForm(form) {
  return {
    senderCurrency: parseIdx(form.senderCurrency),
    senderAmount: parseIdx(form.senderAmount),
    receiverCurrency: parseIdx(form.receiverCurrency),
    receiverAmount: parseIdx(form.receiverAmount),
    source: 'manual',
  };
}

const MONEY_FIELDS = [
  { field: 'senderCurrency', label: 'Sender currency' },
  { field: 'senderAmount', label: 'Sender amount' },
  { field: 'receiverCurrency', label: 'Receiver currency' },
  { field: 'receiverAmount', label: 'Receiver amount' },
];

export default function ColumnManager({
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
  const listRef = useRef(null);
  const searchRef = useRef(null);
  const scrollRef = useRef(0);

  const [search, setSearch] = useState('');
  const [feedback, setFeedback] = useState({ msg: '', tone: '' });
  const [moneyForm, setMoneyForm] = useState(() =>
    extractForm(moneyMappingState?.mapping)
  );

  useBodyScrollLock(isOpen);

  useEffect(() => {
    setMoneyForm(extractForm(moneyMappingState?.mapping));
  }, [moneyMappingState]);

  useEffect(() => {
    const el = dialogRef.current;
    if (!el) return undefined;

    if (isOpen && !el.open) {
      setSearch('');
      setFeedback({
        msg: 'Select up to ' + MAX_VISIBLE_COLUMNS + ' columns.',
        tone: '',
      });
      scrollRef.current = 0;
      el.showModal?.();
      window.requestAnimationFrame(() => {
        searchRef.current?.focus();
      });
    } else if (!isOpen && el.open) {
      el.close();
    }
  }, [isOpen]);

  useEffect(() => {
    const el = dialogRef.current;
    if (!el) return undefined;
    el.addEventListener('close', onClose);
    return () => el.removeEventListener('close', onClose);
  }, [onClose]);

  const hasHeaders = headers.length > 0;
  const selectedSet = useMemo(
    () => new Set(selectedColumnIndexes),
    [selectedColumnIndexes]
  );

  const visibleOptions = useMemo(() => {
    const q = normalizeColumnPickerSearch(search);
    return headers
      .map((label, index) => {
        const key = toSpreadsheetColumn(index);
        const hay = (key + ' ' + String(label || '')).toLowerCase();
        if (q && !hay.includes(q)) return null;
        return {
          index,
          key,
          label: label || 'Column ' + (index + 1),
          selected: selectedSet.has(index),
        };
      })
      .filter(Boolean)
      .sort((a, b) => {
        if (a.selected !== b.selected) return a.selected ? -1 : 1;
        return a.index - b.index;
      });
  }, [headers, selectedSet, search]);

  useLayoutEffect(() => {
    if (listRef.current) listRef.current.scrollTop = scrollRef.current;
  }, [visibleOptions]);

  const handleToggle = (idx, checked) => {
    scrollRef.current = listRef.current?.scrollTop ?? 0;
    const result = onToggleColumn(idx, checked);

    if (!result?.ok) {
      if (result?.reason === 'max-reached') {
        setFeedback({
          msg: 'Maximum is ' + MAX_VISIBLE_COLUMNS + ' columns.',
          tone: 'error',
        });
      } else if (result?.reason === 'min-reached') {
        setFeedback({
          msg: 'At least one column must stay selected.',
          tone: 'error',
        });
      }
      return;
    }
    setFeedback({
      msg: result.nextSelection.length + ' columns selected.',
      tone: 'success',
    });
  };

  const handleReset = () => {
    const res = onResetColumns();
    if (!res?.indexes) return;
    setFeedback({
      msg:
        res.source === 'preset'
          ? 'Reset to preset columns for this sheet.'
          : 'Reset to first ' + res.indexes.length + ' columns.',
      tone: 'success',
    });
  };

  const handleMoneyChange = useCallback(
    (field) => (e) => {
      const next = { ...moneyForm, [field]: e.target.value };
      setMoneyForm(next);
      const applied = onApplyMoneyMapping(readForm(next));
      if (applied === true)
        setFeedback({ msg: 'Money mapping saved.', tone: 'success' });
      else if (applied === false)
        setFeedback({ msg: 'Money mapping cleared.', tone: '' });
    },
    [moneyForm, onApplyMoneyMapping]
  );

  const handleAutoDetect = () => {
    onAutoDetectMoney();
    setFeedback({
      msg: 'Auto-detected money columns applied. Review before saving.',
      tone: 'success',
    });
  };

  const handleClear = () => {
    onClearMoneyMapping();
    setMoneyForm({
      senderCurrency: '',
      senderAmount: '',
      receiverCurrency: '',
      receiverAmount: '',
    });
    setFeedback({ msg: 'Money mapping cleared.', tone: '' });
  };

  const handleSearchKeyDown = (e) => {
    if (e.key === 'Escape' && e.target.value) {
      e.preventDefault();
      e.stopPropagation();
      setSearch('');
      scrollRef.current = 0;
    }
  };

  const fieldOptions = headers.map((label, idx) => ({
    value: String(idx),
    label: toSpreadsheetColumn(idx) + ' — ' + (label || 'Column ' + (idx + 1)),
  }));

  const manualMappingExists = hasAnyMoneyMapping({
    ...readForm(moneyForm),
    source: 'manual',
  });

  const feedbackCls = [
    'dialog-note',
    feedback.tone === 'error' && 'is-error',
    feedback.tone === 'success' && 'is-success',
  ]
    .filter(Boolean)
    .join(' ');

  const metaText = !hasHeaders
    ? 'No columns available'
    : visibleOptions.length + ' of ' + headers.length + ' columns';

  return (
    <dialog className="columns-dialog" ref={dialogRef}>
      <form method="dialog" className="columns-dialog-shell">
        {/* Header */}
        <header className="columns-dialog-header">
          <h2>Manage columns</h2>
          <button type="submit" className="ghost-button">
            Close
          </button>
        </header>

        {/* Toolbar row */}
        <div className="columns-toolbar">
          <div className="columns-toolbar-copy">
            <span className="column-count">
              {selectedColumnIndexes.length}/{MAX_VISIBLE_COLUMNS} selected
            </span>
            <span className="columns-search-meta">{metaText}</span>
          </div>
          <button
            type="button"
            className="ghost-button"
            disabled={!hasWorkbook || !hasHeaders}
            onClick={handleReset}
          >
            Reset defaults
          </button>
        </div>

        {/* Money mapping */}
        <section className="columns-search-panel money-mapping-panel">
          <div className="money-mapping-header">
            <h3>Money mapping</h3>
            <div className="money-mapping-actions">
              <button
                type="button"
                className="ghost-button"
                disabled={!hasWorkbook || !hasHeaders}
                onClick={handleAutoDetect}
              >
                Auto-detect
              </button>
              <button
                type="button"
                className="ghost-button"
                disabled={!hasWorkbook || !manualMappingExists}
                onClick={handleClear}
              >
                Clear
              </button>
            </div>
          </div>

          <div className="money-mapping-grid">
            {MONEY_FIELDS.map(({ field, label }) => (
              <label key={field} className="money-mapping-field">
                <span>{label}</span>
                <select
                  className="money-mapping-select"
                  disabled={!hasWorkbook || !hasHeaders}
                  value={moneyForm[field]}
                  onChange={handleMoneyChange(field)}
                >
                  <option value="">Not mapped</option>
                  {fieldOptions.map((opt) => (
                    <option key={opt.value} value={opt.value}>
                      {opt.label}
                    </option>
                  ))}
                </select>
              </label>
            ))}
          </div>
        </section>

        {/* Feedback note */}
        <p className={feedbackCls}>{feedback.msg}</p>

        {/* Column search */}
        <label className="columns-search-field" htmlFor="ep-col-search">
          <span className="columns-search-label">Search columns</span>
          <input
            id="ep-col-search"
            ref={searchRef}
            type="search"
            className="columns-search-input"
            placeholder="Column letter or name"
            autoComplete="off"
            spellCheck="false"
            value={search}
            onChange={(e) => {
              setSearch(e.target.value);
              scrollRef.current = 0;
            }}
            onKeyDown={handleSearchKeyDown}
          />
        </label>

        {/* Column list */}
        <div
          ref={listRef}
          className="columns-options"
          onScroll={() => {
            scrollRef.current = listRef.current?.scrollTop ?? 0;
          }}
        >
          {!visibleOptions.length ? (
            <div className="columns-options-empty">
              {search
                ? 'No columns match your search.'
                : 'No columns available.'}
            </div>
          ) : (
            visibleOptions.map((item) => (
              <label
                key={item.index}
                className={
                  'columns-option' + (item.selected ? ' is-selected' : '')
                }
              >
                <input
                  type="checkbox"
                  value={String(item.index)}
                  checked={item.selected}
                  disabled={!hasWorkbook}
                  onChange={(e) => handleToggle(item.index, e.target.checked)}
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

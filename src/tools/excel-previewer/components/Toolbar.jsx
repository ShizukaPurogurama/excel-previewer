import { useEffect, useState } from 'react';
import { RELOAD_INTERVAL_OPTIONS } from '../../../constants.js';
import { formatInterval } from '../utils/columns.js';

const VIEWS = [
  { id: 'preview', label: 'Raw rows' },
  { id: 'parsed', label: 'Parsed table' },
  { id: 'changelog', label: "What's new" },
];

export default function Toolbar({
  fileName,
  sheetNames,
  currentSheetName,
  headerIndex,
  reloadMode,
  reloadInterval,
  isBusy,
  activeView,
  onPickFile,
  onSheetChange,
  onApplyHeader,
  onReloadModeChange,
  onReloadIntervalChange,
  onViewChange,
}) {
  const [headerValue, setHeaderValue] = useState(String((headerIndex || 0) + 1));

  useEffect(() => {
    setHeaderValue(String((headerIndex || 0) + 1));
  }, [headerIndex, currentSheetName, fileName]);

  const disabled = isBusy;

  const handleHeaderKeyDown = (e) => {
    if (e.key === 'Enter') {
      e.preventDefault();
      onApplyHeader(headerValue);
    }
  };

  const toggleReload = () => {
    onReloadModeChange(reloadMode === 'auto' ? 'manual' : 'auto');
  };

  return (
    <div className="ep-toolbar" role="toolbar" aria-label="Workbook controls">
      {/* File */}
      <div className="ep-toolbar-group">
        <button
          type="button"
          className="ep-file-pill"
          disabled={isBusy}
          onClick={onPickFile}
          title="Open a different file"
          aria-label={`File: ${fileName}. Click to open a different file.`}
        >
          {fileName}
        </button>
      </div>

      <div className="ep-toolbar-divider" aria-hidden="true" />

      {/* Sheet */}
      <div className="ep-toolbar-group">
        <label className="ep-ctrl-label" htmlFor="ep-sheet-sel">Sheet</label>
        <select
          id="ep-sheet-sel"
          className="ep-ctrl-select"
          value={currentSheetName}
          disabled={disabled || sheetNames.length <= 1}
          onChange={(e) => onSheetChange(e.target.value)}
          aria-label="Select sheet"
        >
          {sheetNames.map((name) => (
            <option key={name} value={name}>{name}</option>
          ))}
        </select>
      </div>

      <div className="ep-toolbar-divider" aria-hidden="true" />

      {/* Header row */}
      <div className="ep-toolbar-group">
        <label className="ep-ctrl-label" htmlFor="ep-header-in">Header</label>
        <div className="ep-header-ctrl">
          <input
            id="ep-header-in"
            type="number"
            min="1"
            step="1"
            inputMode="numeric"
            className="ep-header-input"
            value={headerValue}
            disabled={disabled}
            onChange={(e) => setHeaderValue(e.target.value)}
            onKeyDown={handleHeaderKeyDown}
            aria-label="Header row number (1-based)"
          />
          <button
            type="button"
            className="ep-apply-btn"
            disabled={disabled}
            onClick={() => onApplyHeader(headerValue)}
            aria-label="Apply header row"
          >
            Use
          </button>
        </div>
      </div>

      <div className="ep-toolbar-divider" aria-hidden="true" />

      {/* Reload mode */}
      <div className="ep-toolbar-group">
        <button
          type="button"
          className={'ep-reload-btn' + (reloadMode === 'auto' ? ' ep-auto' : '')}
          disabled={disabled}
          onClick={toggleReload}
          aria-pressed={reloadMode === 'auto' ? 'true' : 'false'}
          title={
            reloadMode === 'auto'
              ? 'Auto-reload active. Click to switch to manual.'
              : 'Manual reload. Click to enable auto-reload.'
          }
        >
          ⟳{' '}
          {reloadMode === 'auto'
            ? 'Auto ' + formatInterval(reloadInterval)
            : 'Manual'}
        </button>

        {reloadMode === 'auto' && (
          <select
            className="ep-ctrl-select"
            value={String(reloadInterval)}
            disabled={disabled}
            onChange={(e) => onReloadIntervalChange(Number(e.target.value))}
            aria-label="Auto-reload interval"
          >
            {RELOAD_INTERVAL_OPTIONS.map((opt) => (
              <option key={opt.value} value={String(opt.value)}>
                {opt.label}
              </option>
            ))}
          </select>
        )}
      </div>

      <div className="ep-toolbar-spacer" />

      {/* View tabs */}
      <div className="ep-tabs" role="tablist" aria-label="Switch view">
        {VIEWS.map((v) => (
          <button
            key={v.id}
            type="button"
            role="tab"
            aria-selected={activeView === v.id ? 'true' : 'false'}
            className={'ep-tab' + (activeView === v.id ? ' ep-tab-active' : '')}
            onClick={() => onViewChange(v.id)}
          >
            {v.label}
          </button>
        ))}
      </div>
    </div>
  );
}

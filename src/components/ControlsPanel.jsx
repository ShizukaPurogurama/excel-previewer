import { useEffect, useState } from 'react';
import { FILE_INPUT_ACCEPT, RELOAD_INTERVAL_OPTIONS } from '../constants.js';

export default function ControlsPanel({
  hasWorkbook,
  isBusy,
  fileName,
  sheetNames,
  currentSheetName,
  headerIndex,
  reloadMode,
  reloadInterval,
  onFileInputClick,
  onFileInputChange,
  onSheetChange,
  onApplyHeaderRow,
  onReloadModeChange,
  onReloadIntervalChange,
}) {
  const [headerInputValue, setHeaderInputValue] = useState('1');

  useEffect(() => {
    setHeaderInputValue(String((headerIndex || 0) + 1));
  }, [headerIndex, currentSheetName, fileName]);

  const sheetSelectDisabled = isBusy || !hasWorkbook;
  const headerInputDisabled = isBusy || !hasWorkbook;
  const reloadModeDisabled = isBusy || !hasWorkbook;
  const reloadIntervalDisabled = reloadModeDisabled || reloadMode !== 'auto';

  const handleHeaderApply = () => {
    onApplyHeaderRow(headerInputValue);
  };

  const handleHeaderKeydown = (event) => {
    if (event.key === 'Enter') {
      event.preventDefault();
      handleHeaderApply();
    }
  };

  return (
    <section className="controls-panel">
      <div className="controls-grid">
        <label className="field">
          <span className="field-label">Workbook</span>
          <input
            id="file-input"
            className="file-input"
            type="file"
            accept={FILE_INPUT_ACCEPT}
            disabled={isBusy}
            onClick={onFileInputClick}
            onChange={onFileInputChange}
          />
          <small>Your file stays local to this browser session.</small>
        </label>

        <label className="field">
          <span className="field-label">Sheet</span>
          <select
            id="sheet-select"
            disabled={sheetSelectDisabled}
            value={currentSheetName || ''}
            onChange={(event) => onSheetChange(event.target.value)}
          >
            {!sheetNames.length ? (
              <option value="">Upload a workbook first</option>
            ) : (
              sheetNames.map((name) => (
                <option key={name} value={name}>
                  {name}
                </option>
              ))
            )}
          </select>
        </label>

        <div className="field">
          <span className="field-label">Header row</span>
          <div className="header-picker">
            <input
              id="header-row-input"
              type="number"
              min="1"
              step="1"
              inputMode="numeric"
              disabled={headerInputDisabled}
              value={headerInputValue}
              onChange={(event) => setHeaderInputValue(event.target.value)}
              onKeyDown={handleHeaderKeydown}
            />
            <button
              id="apply-header-row"
              className="action-button"
              type="button"
              disabled={headerInputDisabled}
              onClick={handleHeaderApply}
            >
              Use row
            </button>
          </div>
        </div>

        <div className="field">
          <span className="field-label">Reload mode</span>
          <div className="reload-mode-picker">
            <select
              id="reload-mode-select"
              disabled={reloadModeDisabled}
              value={reloadMode}
              onChange={(event) => onReloadModeChange(event.target.value)}
            >
              <option value="manual">Manual</option>
              <option value="auto">Auto reload</option>
            </select>
            <select
              id="reload-interval-select"
              disabled={reloadIntervalDisabled}
              value={String(reloadInterval)}
              onChange={(event) => onReloadIntervalChange(Number(event.target.value))}
            >
              {RELOAD_INTERVAL_OPTIONS.map((option) => (
                <option key={option.value} value={String(option.value)}>
                  {option.label}
                </option>
              ))}
            </select>
          </div>
        </div>
      </div>
    </section>
  );
}

import { useCallback, useRef, useState } from 'react';
import { FILE_INPUT_ACCEPT } from '../../constants.js';
import { useToasts } from './hooks/useToasts.js';
import { useWorkbook } from './hooks/useWorkbook.js';
import './previewer.css';

import ChangelogPanel from './components/ChangelogPanel.jsx';
import ColumnManager from './components/ColumnManager.jsx';
import DataTable from './components/DataTable.jsx';
import DropZone from './components/DropZone.jsx';
import FloatingFab from './components/FloatingFab.jsx';
import RawTable from './components/RawTable.jsx';
import StatusStrip from './components/StatusStrip.jsx';
import Toast from './components/Toast.jsx';
import Toolbar from './components/Toolbar.jsx';

export default function ExcelPreviewerTool() {
  const [activeView, setActiveView] = useState('parsed');
  const [isColumnsOpen, setIsColumnsOpen] = useState(false);
  const [isDragOver, setIsDragOver] = useState(false);
  const [statusMessage, setStatusMessage] = useState('Upload a workbook to begin.');
  const [statusTone, setStatusTone] = useState('');

  const fileInputRef = useRef(null);
  const { toasts, showToast } = useToasts();

  const showStatus = useCallback(
    (message, tone, hideToast) => {
      if (!message) return;
      setStatusMessage(message);
      setStatusTone(tone || '');
      if (!hideToast) showToast(message, tone);
    },
    [showToast]
  );

  const workbook = useWorkbook({ showStatus });
  const { handleFileInputChange, handleFileInputClick } = workbook;
  const hasWorkbook = Boolean(workbook.workbook);

  const triggerFileInput = useCallback(() => {
    fileInputRef.current?.click();
  }, []);

  // Accepted by handleFileInputChange which only reads event.target.files[0]
  const handleDroppedFile = useCallback(
    (file) => {
      handleFileInputChange({ target: { files: [file] } });
    },
    [handleFileInputChange]
  );

  const handleDragOver = useCallback((e) => {
    e.preventDefault();
    setIsDragOver(true);
  }, []);

  const handleDragLeave = useCallback((e) => {
    if (!e.currentTarget.contains(e.relatedTarget)) {
      setIsDragOver(false);
    }
  }, []);

  const handleDrop = useCallback(
    (e) => {
      e.preventDefault();
      setIsDragOver(false);
      const file = e.dataTransfer.files[0];
      if (file) handleDroppedFile(file);
    },
    [handleDroppedFile]
  );

  const handleOpenColumns = useCallback(() => {
    if (hasWorkbook && workbook.parsedTable.headers.length) {
      setIsColumnsOpen(true);
    }
  }, [hasWorkbook, workbook.parsedTable.headers.length]);

  const isStatusRelevant =
    statusMessage && statusMessage !== 'Upload a workbook to begin.';

  return (
    <div
      className={'ep-root' + (isDragOver ? ' ep-drag-active' : '')}
      onDragOver={handleDragOver}
      onDragLeave={handleDragLeave}
      onDrop={handleDrop}
    >
      {/* Hidden file input — FSA intercepts click; fallback uses onChange */}
      <input
        ref={fileInputRef}
        type="file"
        accept={FILE_INPUT_ACCEPT}
        className="ep-hidden-input"
        aria-hidden="true"
        tabIndex={-1}
        onClick={handleFileInputClick}
        onChange={handleFileInputChange}
      />

      {isDragOver && (
        <div className="ep-drag-overlay" aria-hidden="true">
          Drop to open workbook
        </div>
      )}

      {!hasWorkbook && (
        <DropZone isBusy={workbook.isBusy} onPickFile={triggerFileInput} />
      )}

      {hasWorkbook && (
        <div className="ep-workspace">
          <Toolbar
            fileName={workbook.fileName}
            sheetNames={workbook.sheetNames}
            currentSheetName={workbook.currentSheetName}
            headerIndex={workbook.headerIndex}
            reloadMode={workbook.reloadMode}
            reloadInterval={workbook.reloadInterval}
            isBusy={workbook.isBusy}
            activeView={activeView}
            onPickFile={triggerFileInput}
            onSheetChange={workbook.selectSheet}
            onApplyHeader={workbook.applyManualHeaderRow}
            onReloadModeChange={workbook.setReloadMode}
            onReloadIntervalChange={workbook.setReloadInterval}
            onViewChange={setActiveView}
          />

          {isStatusRelevant && (
            <StatusStrip message={statusMessage} tone={statusTone} />
          )}

          <div className="ep-view">
            <RawTable
              isVisible={activeView === 'preview'}
              rows={workbook.rows}
              headerIndex={workbook.headerIndex}
              mergedCovered={workbook.mergedCovered}
              mergeAnchors={workbook.mergeAnchors}
              mergeSummary={workbook.mergeSummary}
              onSelectHeaderRow={workbook.applyHeaderIndex}
              showStatus={showStatus}
            />
            <DataTable
              isVisible={activeView === 'parsed'}
              parsedTable={workbook.parsedTable}
              selectedColumnIndexes={workbook.selectedColumnIndexes}
              headerIndex={workbook.headerIndex}
              mergeSummary={workbook.mergeSummary}
              sheetName={workbook.currentSheetName}
              showStatus={showStatus}
            />
            <ChangelogPanel isVisible={activeView === 'changelog'} />
          </div>
        </div>
      )}

      <FloatingFab
        hasWorkbook={hasWorkbook}
        isBusy={workbook.isBusy}
        reloadMode={workbook.reloadMode}
        reloadInterval={workbook.reloadInterval}
        fileHandle={workbook.fileHandle}
        supportsFSA={workbook.supportsFSA}
        selectedColumnCount={workbook.selectedColumnIndexes.length}
        onOpenChangelog={() => setActiveView('changelog')}
        onOpenColumns={handleOpenColumns}
        onReload={workbook.reselectAndReload}
      />

      <ColumnManager
        isOpen={isColumnsOpen}
        onClose={() => setIsColumnsOpen(false)}
        hasWorkbook={hasWorkbook}
        headers={workbook.parsedTable.headers}
        selectedColumnIndexes={workbook.selectedColumnIndexes}
        moneyMappingState={workbook.parsedTable.moneyMappingState}
        onToggleColumn={workbook.updateColumnSelection}
        onResetColumns={workbook.resetColumnSelection}
        onApplyMoneyMapping={workbook.applyMoneyMapping}
        onAutoDetectMoney={workbook.autoDetectMoneyMapping}
        onClearMoneyMapping={workbook.clearMoneyMapping}
      />

      <Toast toasts={toasts} />
    </div>
  );
}

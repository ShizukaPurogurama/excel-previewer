import { useCallback, useState } from 'react';
import { MAX_VISIBLE_COLUMNS } from '../../constants.js';
import { useToasts } from './hooks/useToasts.js';
import { useWorkbook } from './hooks/useWorkbook.js';
import ControlsPanel from './components/ControlsPanel.jsx';
import StatusNote from './components/StatusNote.jsx';
import StatsGrid from './components/StatsGrid.jsx';
import ViewSwitcher from './components/ViewSwitcher.jsx';
import PreviewTable from './components/PreviewTable.jsx';
import ParsedTable from './components/ParsedTable.jsx';
import Changelog from './components/Changelog.jsx';
import FloatingSettings from './components/FloatingSettings.jsx';
import ColumnsDialog from './components/ColumnsDialog.jsx';
import ToastContainer from './components/ToastContainer.jsx';

export default function ExcelPreviewerTool() {
  const { toasts, showToast } = useToasts();

  const [statusMessage, setStatusMessage] = useState('Upload a workbook to begin.');
  const [statusTone, setStatusTone] = useState('');
  const [activeView, setActiveView] = useState('parsed');
  const [isColumnsDialogOpen, setIsColumnsDialogOpen] = useState(false);

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

  const hasWorkbook = Boolean(workbook.workbook);
  const recordCount = workbook.parsedTable.records.length;
  const headerRowLabel = workbook.rows.length ? 'Row ' + (workbook.headerIndex + 1) : '-';

  const handleApplyMoneyMapping = (mapping) => workbook.applyMoneyMapping(mapping);

  return (
    <>
      <div className="app-shell">
        <ControlsPanel
          hasWorkbook={hasWorkbook}
          isBusy={workbook.isBusy}
          fileName={workbook.fileName}
          sheetNames={workbook.sheetNames}
          currentSheetName={workbook.currentSheetName}
          headerIndex={workbook.headerIndex}
          reloadMode={workbook.reloadMode}
          reloadInterval={workbook.reloadInterval}
          onFileInputClick={workbook.handleFileInputClick}
          onFileInputChange={workbook.handleFileInputChange}
          onSheetChange={workbook.selectSheet}
          onApplyHeaderRow={workbook.applyManualHeaderRow}
          onReloadModeChange={workbook.setReloadMode}
          onReloadIntervalChange={workbook.setReloadInterval}
        />

        <StatusNote message={statusMessage} tone={statusTone} />

        <StatsGrid
          fileName={workbook.fileName}
          sheetName={workbook.currentSheetName}
          headerRowLabel={headerRowLabel}
          recordCount={recordCount}
        />

        <ViewSwitcher activeView={activeView} onChange={setActiveView} />

        <PreviewTable
          isVisible={activeView === 'preview'}
          rows={workbook.rows}
          headerIndex={workbook.headerIndex}
          mergedCovered={workbook.mergedCovered}
          mergeAnchors={workbook.mergeAnchors}
          mergeSummary={workbook.mergeSummary}
          onSelectHeaderRow={workbook.applyHeaderIndex}
          showStatus={showStatus}
        />

        <ParsedTable
          isVisible={activeView === 'parsed'}
          parsedTable={workbook.parsedTable}
          selectedColumnIndexes={workbook.selectedColumnIndexes}
          headerIndex={workbook.headerIndex}
          mergeSummary={workbook.mergeSummary}
          showStatus={showStatus}
        />

        <Changelog isVisible={activeView === 'changelog'} />
      </div>

      <ColumnsDialog
        isOpen={isColumnsDialogOpen}
        onClose={() => setIsColumnsDialogOpen(false)}
        hasWorkbook={hasWorkbook}
        headers={workbook.parsedTable.headers}
        selectedColumnIndexes={workbook.selectedColumnIndexes}
        moneyMappingState={workbook.parsedTable.moneyMappingState}
        onToggleColumn={workbook.updateColumnSelection}
        onResetColumns={workbook.resetColumnSelection}
        onApplyMoneyMapping={handleApplyMoneyMapping}
        onAutoDetectMoney={workbook.autoDetectMoneyMapping}
        onClearMoneyMapping={workbook.clearMoneyMapping}
      />

      <FloatingSettings
        hasWorkbook={hasWorkbook}
        isBusy={workbook.isBusy}
        reloadMode={workbook.reloadMode}
        reloadInterval={workbook.reloadInterval}
        fileHandle={workbook.fileHandle}
        supportsFSA={workbook.supportsFSA}
        selectedColumnCount={workbook.selectedColumnIndexes.length}
        maxVisibleColumns={MAX_VISIBLE_COLUMNS}
        onOpenChangelog={() => setActiveView('changelog')}
        onOpenColumns={() => {
          if (hasWorkbook && workbook.parsedTable.headers.length) {
            setIsColumnsDialogOpen(true);
          }
        }}
        onReload={workbook.reselectAndReload}
      />

      <ToastContainer toasts={toasts} />
    </>
  );
}

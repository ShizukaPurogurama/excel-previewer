import { useCallback, useEffect, useRef, useState } from 'react';
import { APP_VERSION, MAX_VISIBLE_COLUMNS } from '../../../constants.js';
import { formatInterval } from '../utils/columns.js';

function reloadSub({ hasWorkbook, isBusy, reloadMode, reloadInterval, fileHandle, supportsFSA }) {
  if (!hasWorkbook) return 'Load a file first';
  if (isBusy) return 'Working…';
  if (reloadMode === 'auto') return 'Auto (' + formatInterval(reloadInterval) + ')';
  if (fileHandle) return 'One-click reload ready';
  if (supportsFSA) return 'Click to connect local reload';
  return 'Click to re-select the updated file';
}

export default function FloatingFab({
  hasWorkbook,
  isBusy,
  reloadMode,
  reloadInterval,
  fileHandle,
  supportsFSA,
  selectedColumnCount,
  onOpenChangelog,
  onOpenColumns,
  onReload,
}) {
  const [isOpen, setIsOpen] = useState(false);
  const wrapRef = useRef(null);
  const close = useCallback(() => setIsOpen(false), []);

  useEffect(() => {
    if (!isOpen) return undefined;
    const handler = (e) => {
      if (wrapRef.current && !wrapRef.current.contains(e.target)) close();
    };
    document.addEventListener('mousedown', handler);
    return () => document.removeEventListener('mousedown', handler);
  }, [isOpen, close]);

  const sub = reloadSub({
    hasWorkbook,
    isBusy,
    reloadMode,
    reloadInterval,
    fileHandle,
    supportsFSA,
  });

  return (
    <div className="ep-fab-wrap" ref={wrapRef}>
      <button
        type="button"
        className="ep-fab-toggle"
        aria-expanded={isOpen ? 'true' : 'false'}
        aria-label="Quick actions"
        onClick={() => setIsOpen((p) => !p)}
      >
        ⚙
      </button>

      {isOpen && (
        <div className="ep-fab-menu" role="menu" aria-label="Quick actions">
          <button
            type="button"
            role="menuitem"
            className="ep-fab-item"
            onClick={() => {
              close();
              onOpenChangelog();
            }}
          >
            <span className="ep-fab-label">What&rsquo;s new</span>
            <span className="ep-fab-sub">See changes in {APP_VERSION}</span>
          </button>

          <button
            type="button"
            role="menuitem"
            className="ep-fab-item"
            disabled={!hasWorkbook || !selectedColumnCount}
            onClick={() => {
              close();
              onOpenColumns();
            }}
          >
            <span className="ep-fab-label">Manage columns</span>
            <span className="ep-fab-sub">
              {selectedColumnCount}/{MAX_VISIBLE_COLUMNS} selected
            </span>
          </button>

          <button
            type="button"
            role="menuitem"
            className="ep-fab-item ep-fab-primary"
            disabled={!hasWorkbook || isBusy}
            onClick={() => {
              close();
              onReload();
            }}
          >
            <span className="ep-fab-label">Reload workbook</span>
            <span className="ep-fab-sub">{sub}</span>
          </button>
        </div>
      )}
    </div>
  );
}

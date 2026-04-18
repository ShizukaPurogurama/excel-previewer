import { useCallback, useEffect, useRef, useState } from 'react';
import { APP_VERSION } from '../../../constants.js';
import { formatInterval } from '../utils/columns.js';

function resolveReloadMeta({ hasWorkbook, isBusy, reloadMode, reloadInterval, fileHandle, supportsFSA }) {
  if (!hasWorkbook) return 'Load a file first';
  if (isBusy) return 'Working...';
  if (reloadMode === 'auto') return 'Auto-reload active (' + formatInterval(reloadInterval) + ')';
  if (fileHandle) return 'One-click reload is ready';
  if (supportsFSA) return 'Click once to connect local reload';
  return 'Click to re-select the updated file';
}

export default function FloatingSettings({
  hasWorkbook,
  isBusy,
  reloadMode,
  reloadInterval,
  fileHandle,
  supportsFSA,
  selectedColumnCount,
  maxVisibleColumns,
  onOpenChangelog,
  onOpenColumns,
  onReload,
}) {
  const [isOpen, setIsOpen] = useState(false);
  const containerRef = useRef(null);

  const close = useCallback(() => setIsOpen(false), []);

  useEffect(() => {
    if (!isOpen) return undefined;
    const handler = (event) => {
      if (containerRef.current && !containerRef.current.contains(event.target)) {
        setIsOpen(false);
      }
    };
    document.addEventListener('click', handler);
    return () => document.removeEventListener('click', handler);
  }, [isOpen]);

  const reloadMeta = resolveReloadMeta({
    hasWorkbook,
    isBusy,
    reloadMode,
    reloadInterval,
    fileHandle,
    supportsFSA,
  });

  const reloadDisabled = !hasWorkbook || isBusy;
  const columnsDisabled = !hasWorkbook || !selectedColumnCount;

  return (
    <div id="floating-settings" className="floating-settings" ref={containerRef}>
      <button
        id="floating-settings-toggle"
        className="floating-settings-toggle"
        type="button"
        aria-expanded={isOpen ? 'true' : 'false'}
        aria-controls="floating-settings-menu"
        onClick={() => setIsOpen((prev) => !prev)}
      >
        ⚙
      </button>

      <div
        id="floating-settings-menu"
        className={'floating-settings-menu' + (isOpen ? '' : ' is-hidden')}
        role="menu"
        aria-label="Quick actions"
      >
        <button
          id="floating-changelog-button"
          className="floating-menu-item"
          type="button"
          role="menuitem"
          onClick={() => {
            close();
            onOpenChangelog();
          }}
        >
          <span className="floating-reload-title">View changelog</span>
          <span id="floating-changelog-meta" className="floating-reload-meta">
            See what is new in {APP_VERSION}
          </span>
        </button>

        <button
          id="floating-columns-button"
          className="floating-menu-item"
          type="button"
          role="menuitem"
          disabled={columnsDisabled}
          onClick={() => {
            close();
            onOpenColumns();
          }}
        >
          <span className="floating-reload-title">Manage columns</span>
          <span id="column-count" className="floating-reload-meta">
            {selectedColumnCount}/{maxVisibleColumns} selected
          </span>
        </button>

        <button
          id="floating-reload-button"
          className="floating-menu-item is-primary"
          type="button"
          role="menuitem"
          disabled={reloadDisabled}
          onClick={() => {
            close();
            onReload();
          }}
        >
          <span className="floating-reload-title">Reload workbook</span>
          <span id="floating-reload-meta" className="floating-reload-meta">
            {reloadMeta}
          </span>
        </button>
      </div>
    </div>
  );
}

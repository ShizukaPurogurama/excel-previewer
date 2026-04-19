// Kept: click-to-copy for cells and entire columns, with flash feedback — wires clipboard utils to table click events.
import { useCallback, useRef } from 'react';
import {
  buildClipboardTableHtml,
  copyHtmlToClipboard,
  getClipboardCellText,
} from '../utils/clipboard.js';

export function useCellCopy({ showStatus }) {
  const feedbackTimersRef = useRef(new WeakMap());

  const flashCell = useCallback((cell) => {
    if (!cell) return;
    const map = feedbackTimersRef.current;
    const existing = map.get(cell);
    if (existing) clearTimeout(existing);

    cell.classList.add('is-copied');
    const timer = setTimeout(() => {
      cell.classList.remove('is-copied');
      map.delete(cell);
    }, 650);
    map.set(cell, timer);
  }, []);

  const copyColumn = useCallback(
    (cell, columnIndex, copyContext) => {
      if (!copyContext || !Array.isArray(copyContext.records) || columnIndex >= copyContext.headers.length) {
        return;
      }
      if (!copyContext.records.length) {
        showStatus('No rows to copy from this column.', 'error');
        return;
      }

      const columnDataText = copyContext.records.map((row) => getClipboardCellText(row[columnIndex]));
      const plain = columnDataText.join('\n');
      const html = buildClipboardTableHtml(columnDataText.map((value) => [value]));

      copyHtmlToClipboard(plain, html)
        .then(() => {
          flashCell(cell);
          showStatus('Copied entire column', 'success');
        })
        .catch(() => {
          showStatus('Copy failed. Please copy manually.', 'error');
        });
    },
    [flashCell, showStatus]
  );

  const copyCellText = useCallback(
    (cell, text) => {
      const html = buildClipboardTableHtml([[text]]);
      copyHtmlToClipboard(text, html)
        .then(() => {
          flashCell(cell);
          const preview = text ? text.slice(0, 42) + (text.length > 42 ? '...' : '') : '(empty)';
          showStatus('Copied "' + preview + '"', 'success');
        })
        .catch(() => {
          showStatus('Copy failed. Please copy manually.', 'error');
        });
    },
    [flashCell, showStatus]
  );

  const handleTableClick = useCallback(
    (event, copyContext) => {
      if (event.target.closest('button')) return;
      if (event.target.closest('a')) return;
      if (event.target.closest('[data-ep-column-filter]')) return;

      const cell = event.target.closest('td, th');
      if (!cell) return;

      if (cell.tagName === 'TH' && cell.closest('thead')) {
        const ths = Array.from(cell.parentElement.children);
        const colIndex = ths.indexOf(cell);
        if (colIndex !== -1) {
          copyColumn(cell, colIndex, copyContext);
          return;
        }
      }

      const text =
        cell.dataset.formula ||
        (cell.classList.contains('cell-empty') ? '' : cell.textContent.trim());
      copyCellText(cell, text);
    },
    [copyCellText, copyColumn]
  );

  return { handleTableClick };
}

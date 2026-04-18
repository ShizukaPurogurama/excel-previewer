import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import {
  COLUMN_PRESETS,
  COLUMN_STORAGE_KEY,
  FILE_PICKER_OPTIONS,
  MAX_VISIBLE_COLUMNS,
} from '../../../constants.js';
import { clampHeaderIndex, getHeaderSignature } from '../utils/columns.js';
import { detectHeaderRow } from '../utils/headerDetection.js';
import {
  extractSheetData,
  formatMergeSummary,
  readWorkbookFromFile,
  supportsFileSystemAccess,
} from '../utils/xlsx.js';
import {
  buildParsedTable,
  loadColumnSelectionsFromStorage,
  resolveDefaultColumnSelection,
  saveColumnSelectionsToStorage,
} from '../utils/parsedTable.js';
import {
  detectMoneyMapping,
  hasAnyMoneyMapping,
  normalizeMoneyMapping,
  removeMoneyMappingFromStorage,
  saveMoneyMappingToStorage,
} from '../utils/money.js';

function pickPreferredSheet(sheetNames, preferredSheetName) {
  if (preferredSheetName && sheetNames.includes(preferredSheetName)) return preferredSheetName;
  return sheetNames[0] || '';
}

export function useWorkbook({ showStatus }) {
  const [workbook, setWorkbook] = useState(null);
  const [fileName, setFileName] = useState('');
  const [fileHandle, setFileHandle] = useState(null);
  const [currentSheetName, setCurrentSheetName] = useState('');
  const [headerSelections, setHeaderSelections] = useState(() => new Map());
  const [columnSelections, setColumnSelections] = useState(() => new Map());
  const [moneyMappings, setMoneyMappings] = useState(() => new Map());
  const [reloadMode, setReloadMode] = useState('manual');
  const [reloadInterval, setReloadInterval] = useState(1000);
  const [isBusy, setIsBusy] = useState(false);

  const sheetDataCacheRef = useRef(new Map());
  const reloadTimerRef = useRef(null);

  const sheetNames = workbook ? workbook.SheetNames : [];

  const currentSheetData = useMemo(() => {
    if (!workbook || !currentSheetName) return null;
    const cache = sheetDataCacheRef.current;
    if (cache.has(currentSheetName)) return cache.get(currentSheetName);
    const data = extractSheetData(workbook, currentSheetName);
    cache.set(currentSheetName, data);
    return data;
  }, [workbook, currentSheetName]);

  const rows = currentSheetData ? currentSheetData.rows : [];
  const mergedCovered = currentSheetData ? currentSheetData.mergedCovered : null;
  const mergeAnchors = currentSheetData ? currentSheetData.mergeAnchors : null;
  const mergeStats = currentSheetData ? currentSheetData.mergeStats : null;
  const mergeSummary = formatMergeSummary(mergeStats);

  const headerIndex = useMemo(() => {
    if (!currentSheetName) return 0;
    const stored = headerSelections.get(currentSheetName);
    if (typeof stored === 'number') return clampHeaderIndex(stored, rows.length);
    return clampHeaderIndex(detectHeaderRow(rows), rows.length);
  }, [currentSheetName, headerSelections, rows]);

  useEffect(() => {
    if (!currentSheetName) return;
    if (headerSelections.has(currentSheetName)) return;
    if (!rows.length) return;
    const detected = detectHeaderRow(rows);
    setHeaderSelections((prev) => {
      if (prev.has(currentSheetName)) return prev;
      const next = new Map(prev);
      next.set(currentSheetName, detected);
      return next;
    });
  }, [currentSheetName, headerSelections, rows]);

  const parsedTable = useMemo(
    () =>
      buildParsedTable({
        rows,
        headerIndex,
        mergedCovered,
        mergeAnchors,
        moneyMappings,
      }),
    [rows, headerIndex, mergedCovered, mergeAnchors, moneyMappings]
  );

  const selectedColumnIndexes = useMemo(() => {
    const headers = parsedTable.headers;
    const safeTotal = headers.length;
    if (!safeTotal || !currentSheetName) return [];

    const stored = columnSelections.get(currentSheetName) || [];
    const validStored = stored.filter(
      (index) => Number.isInteger(index) && index >= 0 && index < safeTotal
    );
    if (validStored.length) return validStored.slice(0, MAX_VISIBLE_COLUMNS);

    const saved = loadColumnSelectionsFromStorage(COLUMN_STORAGE_KEY, headers);
    if (saved && saved.length) {
      const validSaved = saved
        .filter((index) => Number.isInteger(index) && index >= 0 && index < safeTotal)
        .slice(0, MAX_VISIBLE_COLUMNS);
      if (validSaved.length) return validSaved;
    }

    const defaults = resolveDefaultColumnSelection(
      headers,
      currentSheetName,
      COLUMN_PRESETS,
      MAX_VISIBLE_COLUMNS
    );
    return defaults.indexes;
  }, [parsedTable.headers, currentSheetName, columnSelections]);

  const stopAutoReload = useCallback(() => {
    if (reloadTimerRef.current) {
      clearInterval(reloadTimerRef.current);
      reloadTimerRef.current = null;
    }
  }, []);

  const loadFromFile = useCallback(
    async (file, options = {}) => {
      const {
        fileHandle: handle = null,
        preferredSheetName = '',
        preserveSelections = false,
        statusMessage = 'Reading workbook in your browser...',
        successMessage,
        resetOnError = true,
        hideToast = false,
        isAutoReload = false,
      } = options;

      if (!isAutoReload) setIsBusy(true);
      showStatus(statusMessage, '', hideToast);

      const savedHeaderSelections = preserveSelections
        ? new Map(headerSelections)
        : new Map();
      const savedColumnSelections = preserveSelections
        ? new Map(columnSelections)
        : new Map();
      const savedMoneyMappings = preserveSelections ? new Map(moneyMappings) : new Map();

      try {
        const loaded = await readWorkbookFromFile(file);
        const nextSheetName = pickPreferredSheet(loaded.SheetNames, preferredSheetName);
        const validSheetNames = new Set(loaded.SheetNames);

        const filteredHeaderSelections = new Map();
        savedHeaderSelections.forEach((value, key) => {
          if (validSheetNames.has(key) && typeof value === 'number') {
            filteredHeaderSelections.set(key, value);
          }
        });

        const filteredColumnSelections = new Map();
        savedColumnSelections.forEach((value, key) => {
          if (!validSheetNames.has(key) || !Array.isArray(value)) return;
          const safeIndexes = value
            .filter((idx) => Number.isInteger(idx) && idx >= 0)
            .slice(0, MAX_VISIBLE_COLUMNS);
          if (safeIndexes.length) filteredColumnSelections.set(key, safeIndexes);
        });

        sheetDataCacheRef.current = new Map();
        setWorkbook(loaded);
        setFileName(file.name);
        setFileHandle(handle);
        setHeaderSelections(filteredHeaderSelections);
        setColumnSelections(filteredColumnSelections);
        setMoneyMappings(savedMoneyMappings);
        setCurrentSheetName(nextSheetName);

        showStatus(
          successMessage || 'Loaded ' + file.name + '.',
          'success',
          hideToast
        );
      } catch (error) {
        if (resetOnError) {
          sheetDataCacheRef.current = new Map();
          setWorkbook(null);
          setFileName('');
          setFileHandle(null);
          setHeaderSelections(new Map());
          setColumnSelections(new Map());
          setMoneyMappings(new Map());
          setCurrentSheetName('');
          stopAutoReload();
          setReloadMode('manual');
          setReloadInterval(1000);
        }
        showStatus(error.message || 'The workbook could not be parsed.', 'error');
      } finally {
        if (!isAutoReload) setIsBusy(false);
      }
    },
    [columnSelections, headerSelections, moneyMappings, showStatus, stopAutoReload]
  );

  const handleFileInputChange = useCallback(
    async (event) => {
      const file = event.target.files && event.target.files[0];
      if (!file) return;
      await loadFromFile(file, {
        fileHandle: null,
        preferredSheetName: '',
        preserveSelections: false,
        statusMessage: 'Reading workbook in your browser...',
        successMessage:
          'Loaded ' + file.name + '. Use the floating reload button after saving a new version locally.',
        resetOnError: true,
      });
    },
    [loadFromFile]
  );

  const handleFileInputClick = useCallback(
    async (event) => {
      if (isBusy) return;
      if (!supportsFileSystemAccess()) {
        event.target.value = '';
        return;
      }

      event.preventDefault();

      try {
        const handles = await window.showOpenFilePicker(FILE_PICKER_OPTIONS);
        const handle = handles && handles[0];
        if (!handle) return;

        const file = await handle.getFile();
        await loadFromFile(file, {
          fileHandle: handle,
          preferredSheetName: '',
          preserveSelections: false,
          statusMessage: 'Reading workbook in your browser...',
          successMessage:
            'Loaded ' + file.name + '. Reload and auto-reload are ready.',
          resetOnError: true,
        });
      } catch (error) {
        if (error && error.name === 'AbortError') return;
        showStatus('Could not open the workbook. Please try again.', 'error');
      }
    },
    [isBusy, loadFromFile, showStatus]
  );

  const reselectAndReload = useCallback(async () => {
    if (!workbook) {
      showStatus('Load a workbook first.', 'error');
      return;
    }

    const wasAuto = reloadMode === 'auto';
    if (wasAuto) stopAutoReload();

    if (supportsFileSystemAccess()) {
      try {
        const handles = await window.showOpenFilePicker(FILE_PICKER_OPTIONS);
        const handle = handles && handles[0];
        if (!handle) {
          if (wasAuto) {
            // Effect will restart the timer
          }
          return;
        }
        const file = await handle.getFile();
        await loadFromFile(file, {
          fileHandle: handle,
          preferredSheetName: currentSheetName,
          preserveSelections: true,
          statusMessage: 'Reloading the saved workbook...',
          successMessage: 'Reloaded the latest version of ' + file.name + '.',
          resetOnError: false,
        });
      } catch (error) {
        if (!error || error.name !== 'AbortError') {
          showStatus('Could not reload the workbook. Please try again.', 'error');
        }
      }
      return;
    }

    showStatus(
      'Your browser requires re-selecting the updated file. Choose the saved workbook again.',
      ''
    );
  }, [currentSheetName, loadFromFile, reloadMode, showStatus, stopAutoReload, workbook]);

  useEffect(() => {
    if (reloadMode !== 'auto') {
      stopAutoReload();
      return undefined;
    }
    if (!workbook || !fileHandle || !supportsFileSystemAccess()) {
      stopAutoReload();
      return undefined;
    }

    const timer = setInterval(async () => {
      try {
        const file = await fileHandle.getFile();
        await loadFromFile(file, {
          fileHandle,
          preferredSheetName: currentSheetName,
          preserveSelections: true,
          statusMessage: 'Auto-reloading workbook...',
          successMessage: 'Auto-reloaded the latest version of ' + file.name + '.',
          resetOnError: false,
          hideToast: true,
          isAutoReload: true,
        });
      } catch (err) {
        // eslint-disable-next-line no-console
        console.error('Auto reload failed:', err);
      }
    }, reloadInterval);

    reloadTimerRef.current = timer;

    return () => {
      clearInterval(timer);
      if (reloadTimerRef.current === timer) reloadTimerRef.current = null;
    };
  }, [reloadMode, reloadInterval, workbook, fileHandle, currentSheetName, loadFromFile, stopAutoReload]);

  const selectSheet = useCallback((sheetName) => {
    if (!sheetName) return;
    setCurrentSheetName(sheetName);
  }, []);

  const applyHeaderIndex = useCallback(
    (rawIndex) => {
      if (!currentSheetName) return;
      const safe = clampHeaderIndex(rawIndex, rows.length);
      setHeaderSelections((prev) => {
        const next = new Map(prev);
        next.set(currentSheetName, safe);
        return next;
      });
      showStatus(
        'Using row ' + (safe + 1) + ' as the header for ' + currentSheetName + '.',
        'success'
      );
    },
    [currentSheetName, rows.length, showStatus]
  );

  const applyManualHeaderRow = useCallback(
    (rawValueString) => {
      if (!currentSheetName) return;
      const trimmed = (rawValueString || '').trim();
      const parsed = Number.parseInt(trimmed, 10);
      if (!trimmed || !/^\d+$/.test(trimmed) || !Number.isFinite(parsed)) {
        showStatus('Enter a whole-number header row before applying it.', 'error');
        return;
      }
      if (parsed < 1) {
        showStatus('Header row numbers start at 1.', 'error');
        return;
      }
      applyHeaderIndex(parsed - 1);
    },
    [applyHeaderIndex, currentSheetName, showStatus]
  );

  const updateColumnSelection = useCallback(
    (toggledIndex, nextChecked) => {
      if (!currentSheetName) return { ok: false, reason: 'no-sheet' };
      if (!parsedTable.headers.length) return { ok: false, reason: 'no-headers' };
      if (!Number.isInteger(toggledIndex) || toggledIndex < 0) {
        return { ok: false, reason: 'invalid-index' };
      }

      const current = selectedColumnIndexes;
      const nextSet = new Set(current);

      if (nextChecked) {
        if (!nextSet.has(toggledIndex) && nextSet.size >= MAX_VISIBLE_COLUMNS) {
          return { ok: false, reason: 'max-reached' };
        }
        nextSet.add(toggledIndex);
      } else {
        nextSet.delete(toggledIndex);
        if (!nextSet.size) return { ok: false, reason: 'min-reached' };
      }

      const nextSelection = Array.from(nextSet).sort((a, b) => a - b);
      setColumnSelections((prev) => {
        const next = new Map(prev);
        next.set(currentSheetName, nextSelection);
        return next;
      });
      saveColumnSelectionsToStorage(COLUMN_STORAGE_KEY, parsedTable.headers, nextSelection);
      return { ok: true, nextSelection };
    },
    [currentSheetName, parsedTable.headers, selectedColumnIndexes]
  );

  const resetColumnSelection = useCallback(() => {
    if (!currentSheetName || !parsedTable.headers.length) {
      return { source: 'none', indexes: [] };
    }
    const defaults = resolveDefaultColumnSelection(
      parsedTable.headers,
      currentSheetName,
      COLUMN_PRESETS,
      MAX_VISIBLE_COLUMNS
    );
    setColumnSelections((prev) => {
      const next = new Map(prev);
      next.set(currentSheetName, defaults.indexes);
      return next;
    });
    saveColumnSelectionsToStorage(COLUMN_STORAGE_KEY, parsedTable.headers, defaults.indexes);
    return defaults;
  }, [currentSheetName, parsedTable.headers]);

  const applyMoneyMapping = useCallback(
    (mapping) => {
      if (!currentSheetName) return;
      const headers = parsedTable.headers;
      if (!headers.length) return;

      const signature = getHeaderSignature(headers);
      if (!signature) return;

      const normalized = normalizeMoneyMapping(
        { ...mapping, source: 'manual' },
        headers.length
      );

      setMoneyMappings((prev) => {
        const next = new Map(prev);
        if (hasAnyMoneyMapping(normalized)) {
          next.set(signature, normalized);
          saveMoneyMappingToStorage(headers, normalized);
        } else {
          next.delete(signature);
          removeMoneyMappingFromStorage(headers);
        }
        return next;
      });

      return hasAnyMoneyMapping(normalized);
    },
    [currentSheetName, parsedTable.headers]
  );

  const autoDetectMoneyMapping = useCallback(() => {
    if (!currentSheetName) return;
    const headers = parsedTable.headers;
    if (!headers.length) return;
    const signature = getHeaderSignature(headers);
    if (!signature) return;

    const dataRows = rows.slice(headerIndex + 1);
    const detected = detectMoneyMapping(headers, dataRows);

    setMoneyMappings((prev) => {
      const next = new Map(prev);
      next.set(signature, detected);
      return next;
    });
    saveMoneyMappingToStorage(headers, detected);
  }, [currentSheetName, headerIndex, parsedTable.headers, rows]);

  const clearMoneyMapping = useCallback(() => {
    if (!currentSheetName) return;
    const headers = parsedTable.headers;
    if (!headers.length) return;
    const signature = getHeaderSignature(headers);
    if (!signature) return;

    setMoneyMappings((prev) => {
      if (!prev.has(signature)) return prev;
      const next = new Map(prev);
      next.delete(signature);
      return next;
    });
    removeMoneyMappingFromStorage(headers);
  }, [currentSheetName, parsedTable.headers]);

  const resetViewer = useCallback(() => {
    stopAutoReload();
    sheetDataCacheRef.current = new Map();
    setWorkbook(null);
    setFileName('');
    setFileHandle(null);
    setCurrentSheetName('');
    setHeaderSelections(new Map());
    setColumnSelections(new Map());
    setMoneyMappings(new Map());
    setReloadMode('manual');
    setReloadInterval(1000);
    setIsBusy(false);
  }, [stopAutoReload]);

  return {
    workbook,
    fileName,
    fileHandle,
    currentSheetName,
    sheetNames,
    rows,
    mergedCovered,
    mergeAnchors,
    mergeStats,
    mergeSummary,
    headerIndex,
    parsedTable,
    selectedColumnIndexes,
    moneyMappings,
    reloadMode,
    reloadInterval,
    isBusy,
    supportsFSA: supportsFileSystemAccess(),
    handleFileInputChange,
    handleFileInputClick,
    reselectAndReload,
    selectSheet,
    applyHeaderIndex,
    applyManualHeaderRow,
    updateColumnSelection,
    resetColumnSelection,
    applyMoneyMapping,
    autoDetectMoneyMapping,
    clearMoneyMapping,
    setReloadMode,
    setReloadInterval,
    resetViewer,
  };
}

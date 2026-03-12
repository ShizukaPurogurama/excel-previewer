(function () {
  const PREVIEW_ROW_LIMIT = 24;
  const MAX_VISIBLE_COLUMNS = 7;
  const THEME_STORAGE_KEY = "excel-viewer-theme";
  const COLUMN_STORAGE_KEY = "excel-viewer-columns";
  
  // ---------------------------------------------------------------------------
  // COLUMN_PRESETS — default column display configuration
  //
  // Each key is a sheet name (case-insensitive) or "*" for all sheets.
  // Each value is an ordered list of column header names to show by default.
  //
  // Resolution order (first match wins):
  //   1. User's saved selection in memory / localStorage  ← unchanged
  //   2. Exact sheet-name key in COLUMN_PRESETS           ← new
  //   3. Wildcard "*" key in COLUMN_PRESETS              ← new (replaces old PRIORITY_COLUMNS)
  //   4. First MAX_VISIBLE_COLUMNS columns               ← unchanged fallback
  //
  // Names are matched case-insensitively against the actual header row.
  // Columns not found in the file are silently skipped.
  // ---------------------------------------------------------------------------
  const COLUMN_PRESETS = {
    // Global default — applies to every sheet that has no sheet-specific entry.
    "*": [
      "TestID",
      "Service Name",
      "Sender Account",
      "Sender Amount",
      "Fee",
      "Receiver Account",
      "Receiver Amount",
      "TID",
      "Status",
    ],

    // Per-sheet overrides — uncomment or add entries as needed:
    // "Sheet1": ["ID", "Name", "Amount", "Date"],
    // "Summary": ["Category", "Total"],
  };

  const themeMediaQuery = window.matchMedia("(prefers-color-scheme: dark)");
  const FILE_PICKER_OPTIONS = {
    multiple: false,
    excludeAcceptAllOption: false,
    types: [
      {
        description: "Excel and CSV files",
        accept: {
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"],
          "application/vnd.ms-excel": [".xls", ".xlsb"],
          "application/vnd.oasis.opendocument.spreadsheet": [".ods"],
          "text/csv": [".csv"],
        },
      },
    ],
  };

  const elements = {
    fileInput: document.getElementById("file-input"),
    sheetSelect: document.getElementById("sheet-select"),
    headerRowInput: document.getElementById("header-row-input"),
    applyHeaderRowButton: document.getElementById("apply-header-row"),
    reloadModeSelect: document.getElementById("reload-mode-select"),
    reloadIntervalSelect: document.getElementById("reload-interval-select"),
    settingsToggleButton: document.getElementById("floating-settings-toggle"),
    settingsMenu: document.getElementById("floating-settings-menu"),
    openColumnsDialogButton: document.getElementById("floating-columns-button"),
    columnsDialog: document.getElementById("columns-dialog"),
    closeColumnsDialogButton: document.getElementById("close-columns-dialog"),
    columnsOptions: document.getElementById("columns-options"),
    columnsViewListButton: document.getElementById("columns-view-list"),
    columnsViewGridButton: document.getElementById("columns-view-grid"),
    columnCount: document.getElementById("column-count"),
    columnsDialogCount: document.getElementById("columns-dialog-count"),
    columnsDialogNote: document.getElementById("columns-dialog-note"),
    resetColumnsButton: document.getElementById("reset-columns-button"),
    floatingReloadButton: document.getElementById("floating-reload-button"),
    floatingReloadMeta: document.getElementById("floating-reload-meta"),
    themeToggle: document.getElementById("theme-toggle"),
    themeToggleLabel: document.getElementById("theme-toggle-label"),
    statusNote: document.getElementById("status-note"),
    toastContainer: document.getElementById("toast-container"),
    previewMeta: document.getElementById("preview-meta"),
    previewEmpty: document.getElementById("preview-empty"),
    previewWrap: document.getElementById("preview-wrap"),
    previewTable: document.getElementById("preview-table"),
    parsedMeta: document.getElementById("parsed-meta"),
    parsedEmpty: document.getElementById("parsed-empty"),
    parsedWrap: document.getElementById("parsed-wrap"),
    parsedTable: document.getElementById("parsed-table"),
    viewPreviewTab: document.getElementById("view-preview-tab"),
    viewParsedTab: document.getElementById("view-parsed-tab"),
    previewSection: document.getElementById("preview-section"),
    parsedSection: document.getElementById("parsed-section"),
    statWorkbook: document.getElementById("stat-workbook"),
    statSheet: document.getElementById("stat-sheet"),
    statHeaderRow: document.getElementById("stat-header-row"),
    statRecords: document.getElementById("stat-records"),
  };

  const state = {
    workbook: null,
    fileName: "",
    fileHandle: null,
    currentSheetName: "",
    sheetRows: new Map(),
    sheetMergeStats: new Map(),
    sheetMergedCells: new Map(),
    sheetMergeAnchors: new Map(),
    headerSelections: new Map(),
    columnSelections: new Map(),
    reloadTimer: null,
    isAutoReloading: false,
    autoReloadInterval: 1000,
    activeView: "parsed",
    columnPickerHeaders: [],
    columnPickerView: "list",
  };

  const copyFeedbackTimers = new WeakMap();
  let toastTimerSeed = 0;

  elements.fileInput.addEventListener("change", handleFileUpload);
  elements.fileInput.addEventListener("click", prepareFileInputForReload);
  elements.sheetSelect.addEventListener("change", handleSheetChange);
  elements.applyHeaderRowButton.addEventListener("click", applyManualHeaderRow);
  elements.floatingReloadButton.addEventListener("click", handleFloatingReload);
  elements.themeToggle.addEventListener("click", toggleTheme);
  elements.headerRowInput.addEventListener("keydown", function (event) {
    if (event.key === "Enter") {
      applyManualHeaderRow();
    }
  });
  elements.previewTable.addEventListener("click", handlePreviewClick);
  elements.previewTable.addEventListener("click", handleTableCellCopyClick);
  elements.parsedTable.addEventListener("click", handleTableCellCopyClick);
  elements.reloadModeSelect.addEventListener("change", handleReloadModeChange);
  elements.reloadIntervalSelect.addEventListener("change", handleReloadIntervalChange);
  elements.openColumnsDialogButton.addEventListener("click", openColumnsDialog);
  elements.columnsOptions.addEventListener("change", handleColumnOptionToggle);
  elements.columnsViewListButton.addEventListener("click", function () {
    setColumnPickerView("list");
  });
  elements.columnsViewGridButton.addEventListener("click", function () {
    setColumnPickerView("grid");
  });
  elements.resetColumnsButton.addEventListener("click", handleResetColumns);
  elements.viewPreviewTab.addEventListener("click", handleViewTabClick);
  elements.viewParsedTab.addEventListener("click", handleViewTabClick);
  elements.settingsToggleButton.addEventListener("click", toggleSettingsMenu);
  document.addEventListener("click", handleOutsideSettingsClick);

  if (typeof themeMediaQuery.addEventListener === "function") {
    themeMediaQuery.addEventListener("change", handleSystemThemeChange);
  } else if (typeof themeMediaQuery.addListener === "function") {
    themeMediaQuery.addListener(handleSystemThemeChange);
  }

  function handleSheetChange(event) {
    const sheetName = event.target.value;
    if (!sheetName) {
      return;
    }

    renderSheet(sheetName);
  }

  function handleViewTabClick(event) {
    const view = event.currentTarget.dataset.view;
    setActiveView(view);
  }

  function setActiveView(view) {
    const safeView = view === "parsed" ? "parsed" : "preview";
    state.activeView = safeView;

    const isPreview = safeView === "preview";
    elements.previewSection.classList.toggle("is-hidden", !isPreview);
    elements.parsedSection.classList.toggle("is-hidden", isPreview);
    elements.viewPreviewTab.classList.toggle("is-active", isPreview);
    elements.viewParsedTab.classList.toggle("is-active", !isPreview);
    elements.viewPreviewTab.setAttribute("aria-pressed", String(isPreview));
    elements.viewParsedTab.setAttribute("aria-pressed", String(!isPreview));
  }

  function toggleSettingsMenu() {
    const isOpen = !elements.settingsMenu.classList.contains("is-hidden");
    setSettingsMenuOpen(!isOpen);
  }

  function setSettingsMenuOpen(isOpen) {
    elements.settingsMenu.classList.toggle("is-hidden", !isOpen);
    elements.settingsToggleButton.setAttribute("aria-expanded", String(isOpen));
  }

  function handleOutsideSettingsClick(event) {
    if (elements.settingsMenu.classList.contains("is-hidden")) {
      return;
    }

    const container = document.getElementById("floating-settings");
    if (container && !container.contains(event.target)) {
      setSettingsMenuOpen(false);
    }
  }

  async function prepareFileInputForReload(event) {
    if (elements.fileInput.disabled) {
      return;
    }

    if (!supportsFileSystemAccess()) {
      elements.fileInput.value = "";
      return;
    }

    event.preventDefault();

    try {
      const handles = await window.showOpenFilePicker(FILE_PICKER_OPTIONS);
      const handle = handles && handles[0];

      if (!handle) {
        return;
      }

      const file = await handle.getFile();
      await loadWorkbookFile(file, {
        fileHandle: handle,
        preferredSheetName: "",
        preserveHeaderSelections: false,
        statusMessage: "Reading workbook in your browser...",
        successMessage:
          "Loaded " +
          file.name +
          ". Reload and auto-reload are ready.",
        resetOnError: true,
      });
    } catch (error) {
      if (error && error.name === "AbortError") {
        return;
      }

      showStatus("Could not open the workbook. Please try again.", "error");
    }
  }

  function handleSystemThemeChange(event) {
    if (getStoredTheme()) {
      return;
    }

    applyTheme(event.matches ? "dark" : "light", false);
  }

  async function handleFileUpload(event) {
    const [file] = event.target.files;

    if (!file) {
      return;
    }

    await loadWorkbookFile(file, {
      fileHandle: null,
      preferredSheetName: "",
      preserveHeaderSelections: false,
      statusMessage: "Reading workbook in your browser...",
      successMessage:
        "Loaded " + file.name + ". Use the floating reload button after saving a new version locally.",
      resetOnError: true,
    });
  }

  async function handleFloatingReload() {
    if (!state.workbook) {
      showStatus("Load a workbook first.", "error");
      return;
    }

    // If auto-reload is active, pause it temporarily for manual reload
    const wasAutoReloading = state.isAutoReloading;
    if (wasAutoReloading) {
      stopAutoReload();
    }

    // Always re-select the file to ensure we get the latest version
    // This fixes the issue where the web view was showing cached content
    if (supportsFileSystemAccess()) {
      try {
        const handles = await window.showOpenFilePicker(FILE_PICKER_OPTIONS);
        const handle = handles && handles[0];

        if (!handle) {
          if (wasAutoReloading) {
            startAutoReload();
          }
          return;
        }

        const file = await handle.getFile();
        await loadWorkbookFile(file, {
          fileHandle: handle,
          preferredSheetName: state.currentSheetName,
          preserveHeaderSelections: true,
          statusMessage: "Reloading the saved workbook...",
          successMessage: "Reloaded the latest version of " + file.name + ".",
          resetOnError: false,
        });

        if (wasAutoReloading) {
          startAutoReload();
        }
      } catch (error) {
        if (error && error.name === "AbortError") {
          if (wasAutoReloading) {
            startAutoReload();
          }
          return;
        }

        showStatus("Could not reload the workbook. Please try again.", "error");
        if (wasAutoReloading) {
          startAutoReload();
        }
      }

      return;
    }

    // Fallback for browsers that don't support File System Access API
    showStatus(
      "Your browser requires re-selecting the updated file. Choose the saved workbook again.",
      ""
    );
    elements.fileInput.click();
    if (wasAutoReloading) {
      startAutoReload();
    }
  }

  async function loadWorkbookFile(file, options) {
    if (typeof XLSX === "undefined") {
      showStatus("SheetJS could not be loaded. Check the CDN connection and refresh the page.", "error");
      return;
    }

    const preferredSheetName = options.preferredSheetName || "";
    const preserveHeaderSelections = Boolean(options.preserveHeaderSelections);
    const savedHeaderSelections = preserveHeaderSelections
      ? new Map(state.headerSelections)
      : new Map();
    const savedColumnSelections = preserveHeaderSelections
      ? new Map(state.columnSelections)
      : new Map();

    if (!options.isAutoReload) {
      setBusy(true);
    }
    showStatus(options.statusMessage || "Reading workbook in your browser...", "", options.hideToast);

    try {
      const data = await file.arrayBuffer();
      const workbook = XLSX.read(data, {
        type: "array",
        cellDates: true,
      });

      if (!workbook.SheetNames.length) {
        throw new Error("This workbook does not contain any readable sheets.");
      }

      state.workbook = workbook;
      state.fileName = file.name;
      state.fileHandle = options.fileHandle || null;
      state.currentSheetName = "";
      state.sheetRows.clear();
      state.sheetMergeStats.clear();
      state.sheetMergedCells.clear();
      restoreHeaderSelections(savedHeaderSelections, workbook.SheetNames);
      restoreColumnSelections(savedColumnSelections, workbook.SheetNames);

      populateSheetSelect(workbook.SheetNames);
      renderSheet(
        preferredSheetName && workbook.SheetNames.includes(preferredSheetName)
          ? preferredSheetName
          : workbook.SheetNames[0]
      );
      syncReloadControls(true);
      showStatus(options.successMessage || "Loaded " + file.name + ".", "success", options.hideToast);
    } catch (error) {
      if (options.resetOnError) {
        resetViewer();
      }
      showStatus(error.message || "The workbook could not be parsed.", "error");
    } finally {
      if (!options.isAutoReload) {
        setBusy(false);
      }
    }
  }

  function restoreHeaderSelections(savedHeaderSelections, sheetNames) {
    const validSheetNames = new Set(sheetNames);

    state.headerSelections.clear();
    savedHeaderSelections.forEach(function (headerIndex, sheetName) {
      if (validSheetNames.has(sheetName) && typeof headerIndex === "number") {
        state.headerSelections.set(sheetName, headerIndex);
      }
    });
  }

  function restoreColumnSelections(savedColumnSelections, sheetNames) {
    const validSheetNames = new Set(sheetNames);

    state.columnSelections.clear();
    savedColumnSelections.forEach(function (columnIndexes, sheetName) {
      if (!validSheetNames.has(sheetName) || !Array.isArray(columnIndexes)) {
        return;
      }

      const safeIndexes = columnIndexes
        .filter(function (index) {
          return Number.isInteger(index) && index >= 0;
        })
        .slice(0, MAX_VISIBLE_COLUMNS);

      if (safeIndexes.length) {
        state.columnSelections.set(sheetName, safeIndexes);
      }
    });
  }

  function populateSheetSelect(sheetNames) {
    elements.sheetSelect.innerHTML = "";

    sheetNames.forEach(function (sheetName) {
      const option = document.createElement("option");
      option.value = sheetName;
      option.textContent = sheetName;
      elements.sheetSelect.append(option);
    });

    elements.sheetSelect.disabled = false;
  }

  function renderSheet(sheetName) {
    const rows = getRowsForSheet(sheetName);
    const storedHeaderIndex = state.headerSelections.get(sheetName);
    const headerIndex =
      typeof storedHeaderIndex === "number" ? storedHeaderIndex : detectHeaderRow(rows);

    state.currentSheetName = sheetName;
    state.headerSelections.set(sheetName, headerIndex);
    elements.sheetSelect.value = sheetName;

    const mergeSummary = formatMergeSummary(getMergeStatsForSheet(sheetName));

    syncHeaderControls(rows.length, headerIndex);
    renderPreview(rows, headerIndex, mergeSummary);
    renderParsedTable(rows, headerIndex, mergeSummary);
    setActiveView(state.activeView);
  }

  function getRowsForSheet(sheetName) {
    if (state.sheetRows.has(sheetName)) {
      return state.sheetRows.get(sheetName);
    }

    const sheet = state.workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      raw: true,
      defval: "",
      blankrows: true,
    });
    const alignedRows = alignRowsToExcelRowNumbers(sheet, rows);
    const mergeMaps = expandMergedCellsForDisplay(sheet, alignedRows);

    state.sheetMergedCells.set(sheetName, mergeMaps.covered);
    state.sheetMergeAnchors.set(sheetName, mergeMaps.anchors);
    state.sheetRows.set(sheetName, alignedRows);
    return alignedRows;
  }

  function expandMergedCellsForDisplay(sheet, rows) {
    const covered = new Set();
    const anchors = new Map();
    const merges = sheet && Array.isArray(sheet["!merges"]) ? sheet["!merges"] : [];

    merges.forEach(function (range) {
      if (!range || !range.s || !range.e) {
        return;
      }

      const startRow = range.s.r;
      const endRow = range.e.r;
      const startCol = range.s.c;
      const endCol = range.e.c;

      const anchorRow = rows[startRow] || [];
      const anchorValue = anchorRow[startCol];

      for (let rowIndex = startRow; rowIndex <= endRow; rowIndex += 1) {
        if (!rows[rowIndex]) {
          rows[rowIndex] = [];
        }

        for (let colIndex = startCol; colIndex <= endCol; colIndex += 1) {
          if (rowIndex === startRow && colIndex === startCol) {
            continue;
          }

          const cellValue = rows[rowIndex][colIndex];
          if (cellValue === "" || cellValue === null || cellValue === undefined) {
            rows[rowIndex][colIndex] = anchorValue;
          }

          covered.add(rowIndex + ":" + colIndex);
        }
      }

      anchors.set(startRow + ":" + startCol, {
        rowSpan: endRow - startRow + 1,
        colSpan: endCol - startCol + 1,
      });
    });

    return { covered: covered, anchors: anchors };
  }

  function isMergedCarryCell(sheetName, rowIndex, colIndex) {
    const mergeMap = state.sheetMergedCells.get(sheetName);
    if (!mergeMap) {
      return false;
    }

    return mergeMap.has(rowIndex + ":" + colIndex);
  }

  function getMergeAnchorCell(sheetName, rowIndex, colIndex) {
    const anchors = state.sheetMergeAnchors.get(sheetName);
    if (!anchors) {
      return null;
    }

    return anchors.get(rowIndex + ":" + colIndex) || null;
  }

  function getMergeStatsForSheet(sheetName) {
    if (state.sheetMergeStats.has(sheetName)) {
      return state.sheetMergeStats.get(sheetName);
    }

    const sheet = state.workbook && state.workbook.Sheets[sheetName];
    const merges = sheet && Array.isArray(sheet["!merges"]) ? sheet["!merges"] : [];

    let rowMerged = 0;
    let columnMerged = 0;
    let bothMerged = 0;

    merges.forEach(function (range) {
      if (!range || !range.s || !range.e) {
        return;
      }

      const spansRows = range.e.r > range.s.r;
      const spansColumns = range.e.c > range.s.c;

      if (spansColumns) {
        rowMerged += 1;
      }

      if (spansRows) {
        columnMerged += 1;
      }

      if (spansRows && spansColumns) {
        bothMerged += 1;
      }
    });

    const stats = {
      total: merges.length,
      rowMerged: rowMerged,
      columnMerged: columnMerged,
      bothMerged: bothMerged,
    };

    state.sheetMergeStats.set(sheetName, stats);
    return stats;
  }

  function formatMergeSummary(stats) {
    if (!stats || !stats.total) {
      return "No merged cells detected.";
    }

    return (
      "Merged regions: " +
      stats.total +
      " (row merges: " +
      stats.rowMerged +
      ", column merges: " +
      stats.columnMerged +
      ")."
    );
  }

  function alignRowsToExcelRowNumbers(sheet, rows) {
    const reference = sheet["!ref"];

    if (!reference) {
      return rows;
    }

    const range = XLSX.utils.decode_range(reference);

    if (!range.s.r) {
      return rows;
    }

    return Array.from({ length: range.s.r }, function () {
      return [];
    }).concat(rows);
  }

  function syncHeaderControls(rowCount, headerIndex) {
    const safeRowCount = Math.max(rowCount, 1);
    const safeHeaderIndex = clampHeaderIndex(headerIndex, rowCount);
    const hasRows = rowCount > 0;

    elements.headerRowInput.disabled = !hasRows;
    elements.applyHeaderRowButton.disabled = !hasRows;
    elements.headerRowInput.min = "1";
    elements.headerRowInput.max = String(safeRowCount);
    elements.headerRowInput.value = String(safeHeaderIndex + 1);
  }

  function syncReloadControls(hasWorkbook) {
    elements.reloadModeSelect.disabled = !hasWorkbook;
    elements.reloadIntervalSelect.disabled = !hasWorkbook || elements.reloadModeSelect.value !== "auto";
  }

  function handleReloadModeChange(event) {
    const mode = event.target.value;
    const isAuto = mode === "auto";
    state.isAutoReloading = isAuto;
    
    syncReloadControls(!!state.workbook);
    
    if (isAuto) {
      startAutoReload();
    } else {
      stopAutoReload();
    }
  }

  function handleReloadIntervalChange(event) {
    state.autoReloadInterval = Number(event.target.value);
    
    if (state.isAutoReloading) {
      stopAutoReload();
      startAutoReload();
    }
  }

  function startAutoReload() {
    if (!state.workbook || !supportsFileSystemAccess()) {
      return;
    }

    stopAutoReload();
    
    state.reloadTimer = setInterval(async function () {
      try {
        await performAutoReload();
      } catch (error) {
        console.error("Auto reload failed:", error);
      }
    }, state.autoReloadInterval);
    
    updateFloatingReloadButtonState(false);
  }

  function stopAutoReload() {
    if (state.reloadTimer) {
      clearInterval(state.reloadTimer);
      state.reloadTimer = null;
    }
    updateFloatingReloadButtonState(false);
  }

  async function performAutoReload() {
    if (!state.workbook || !state.fileHandle) {
      return;
    }

    try {
      const file = await state.fileHandle.getFile();
      await loadWorkbookFile(file, {
        fileHandle: state.fileHandle,
        preferredSheetName: state.currentSheetName,
        preserveHeaderSelections: true,
        statusMessage: "Auto-reloading workbook...",
        successMessage: "Auto-reloaded the latest version of " + file.name + ".",
        resetOnError: false,
        hideToast: true,
        isAutoReload: true,
      });
    } catch (error) {
      console.error("Auto reload failed:", error);
    }
  }

  function renderPreview(rows, selectedHeaderIndex, mergeSummary) {
    const totalRows = rows.length;

    if (!totalRows) {
      elements.previewMeta.textContent = "The selected sheet is empty. " + mergeSummary;
      showEmpty(elements.previewEmpty, elements.previewWrap);
      elements.previewEmpty.textContent = "This sheet has no rows to preview.";
      return;
    }

    const windowStart =
      selectedHeaderIndex >= PREVIEW_ROW_LIMIT
        ? Math.max(0, selectedHeaderIndex - Math.floor(PREVIEW_ROW_LIMIT / 3))
        : 0;
    const windowEnd = Math.min(totalRows, windowStart + PREVIEW_ROW_LIMIT);
    const visibleRows = rows.slice(windowStart, windowEnd);
    const columnCount = Math.max(
      1,
      rows[selectedHeaderIndex] ? rows[selectedHeaderIndex].length : 0,
      ...visibleRows.map(function (row) {
        return row.length;
      })
    );

    elements.previewMeta.textContent =
      "Showing rows " +
      (windowStart + 1) +
      "-" +
      windowEnd +
      " of " +
      totalRows +
      ". Click a row number to use it as the header. " +
      mergeSummary;

    elements.previewTable.innerHTML = "";

    const thead = document.createElement("thead");
    const headerRow = document.createElement("tr");
    headerRow.append(createTextCell("th", "Row"));

    for (let columnIndex = 0; columnIndex < columnCount; columnIndex += 1) {
      headerRow.append(createTextCell("th", toSpreadsheetColumn(columnIndex)));
    }

    thead.append(headerRow);

    const tbody = document.createElement("tbody");

    visibleRows.forEach(function (row, visibleRowIndex) {
      const actualRowIndex = windowStart + visibleRowIndex;
      const tr = document.createElement("tr");

      if (actualRowIndex === selectedHeaderIndex) {
        tr.classList.add("is-selected");
      }

      const rowButtonCell = document.createElement("th");
      rowButtonCell.scope = "row";

      const rowButton = document.createElement("button");
      rowButton.type = "button";
      rowButton.className = "row-selector";
      rowButton.dataset.rowIndex = String(actualRowIndex);
      rowButton.textContent = "Row " + (actualRowIndex + 1);

      if (actualRowIndex === selectedHeaderIndex) {
        rowButton.classList.add("is-active");
      }

      rowButtonCell.append(rowButton);
      tr.append(rowButtonCell);

      for (let columnIndex = 0; columnIndex < columnCount; columnIndex += 1) {
        if (isMergedCarryCell(state.currentSheetName, actualRowIndex, columnIndex)) {
          continue;
        }

        const value = row[columnIndex];
        const text = formatCellValue(value);
        const td = createTextCell("td", text || "-");
        const anchor = getMergeAnchorCell(state.currentSheetName, actualRowIndex, columnIndex);

        if (!text) {
          td.classList.add("cell-empty");
        }

        if (anchor) {
          if (anchor.rowSpan > 1) {
            td.rowSpan = anchor.rowSpan;
          }
          if (anchor.colSpan > 1) {
            td.colSpan = anchor.colSpan;
          }
          td.classList.add("cell-merged-anchor");
          td.title = "Merged cell (" + anchor.rowSpan + "×" + anchor.colSpan + ")";
        }

        tr.append(td);
      }

      tbody.append(tr);
    });

    elements.previewTable.append(thead, tbody);
    showTable(elements.previewEmpty, elements.previewWrap);
  }

  function openColumnsDialog() {
    if (!state.workbook || !state.currentSheetName || !state.columnPickerHeaders.length) {
      return;
    }

    setSettingsMenuOpen(false);

    if (typeof elements.columnsDialog.showModal === "function") {
      showColumnsDialogNote("Pick up to " + MAX_VISIBLE_COLUMNS + " columns to display.", "");
      elements.columnsDialog.showModal();
    }
  }

  function setColumnPickerView(view) {
    const safeView = view === "grid" ? "grid" : "list";
    state.columnPickerView = safeView;
    elements.columnsOptions.dataset.view = safeView;
    elements.columnsViewListButton.classList.toggle("is-active", safeView === "list");
    elements.columnsViewGridButton.classList.toggle("is-active", safeView === "grid");
  }

  function handleColumnOptionToggle(event) {
    const target = event.target;
    if (!(target instanceof HTMLInputElement) || target.type !== "checkbox") {
      return;
    }

    if (!state.currentSheetName) {
      return;
    }

    const toggledIndex = Number.parseInt(target.value, 10);
    if (!Number.isInteger(toggledIndex) || toggledIndex < 0) {
      return;
    }

    const current = getSelectedColumnIndexesForCurrentSheet(state.columnPickerHeaders.length);
    const nextSet = new Set(current);

    if (target.checked) {
      if (!nextSet.has(toggledIndex) && nextSet.size >= MAX_VISIBLE_COLUMNS) {
        target.checked = false;
        showColumnsDialogNote("Maximum is " + MAX_VISIBLE_COLUMNS + " columns.", "error");
        return;
      }
      nextSet.add(toggledIndex);
    } else {
      nextSet.delete(toggledIndex);
      if (!nextSet.size) {
        target.checked = true;
        showColumnsDialogNote("At least one column must stay selected.", "error");
        return;
      }
    }

    const nextSelection = Array.from(nextSet).sort(function (a, b) {
      return a - b;
    });
    state.columnSelections.set(state.currentSheetName, nextSelection);
    saveColumnSelectionsToStorage(state.columnPickerHeaders, nextSelection);
    renderSheet(state.currentSheetName);
    showColumnsDialogNote("Applied " + nextSelection.length + " selected columns.", "success");
  }

  function renderParsedTable(rows, headerIndex, mergeSummary) {
    const parsed = buildParsedTable(rows, headerIndex);

    elements.statWorkbook.textContent = state.fileName || "None loaded";
    elements.statSheet.textContent = state.currentSheetName || "-";
    elements.statHeaderRow.textContent = rows.length ? "Row " + (headerIndex + 1) : "-";
    elements.statRecords.textContent = String(parsed.records.length);

    if (!parsed.headers.length) {
      syncColumnPicker([], []);
      elements.parsedMeta.textContent =
        "No columns could be derived from the chosen header row. " + mergeSummary;
      elements.parsedEmpty.textContent =
        "Choose a different header row or verify that the sheet contains data.";
      showEmpty(elements.parsedEmpty, elements.parsedWrap);
      return;
    }

    elements.parsedMeta.textContent =
      "Using row " +
      (headerIndex + 1) +
      " as headers. " +
      parsed.records.length +
      " record" +
      (parsed.records.length === 1 ? "" : "s") +
      " across " +
      parsed.headers.length +
      " column" +
      (parsed.headers.length === 1 ? "" : "s") +
      ". " +
      mergeSummary;

    const selectedColumnIndexes = getSelectedColumnIndexesForCurrentSheet(parsed.headers.length);
    const selectedHeaders = selectedColumnIndexes.map(function (index) {
      return parsed.headers[index];
    });
    const selectedRecords = parsed.records.map(function (record) {
      return selectedColumnIndexes.map(function (index) {
        return record[index];
      });
    });
    const selectedMergedRecords = parsed.mergedCarryRows.map(function (rowFlags) {
      return selectedColumnIndexes.map(function (index) {
        return Boolean(rowFlags[index]);
      });
    });
    const selectedAnchorRecords = parsed.mergedAnchors.map(function (anchorRow) {
      return selectedColumnIndexes.map(function (index) {
        return anchorRow[index];
      });
    });

    syncColumnPicker(parsed.headers, selectedColumnIndexes);

    elements.parsedTable.innerHTML = "";

    const thead = document.createElement("thead");
    const headerRow = document.createElement("tr");
    selectedHeaders.forEach(function (label) {
      headerRow.append(createTextCell("th", label));
    });
    thead.append(headerRow);

    const tbody = document.createElement("tbody");

    if (!selectedRecords.length) {
      const emptyRow = document.createElement("tr");
      const cell = createTextCell(
        "td",
        "Header row selected, but there are no non-empty data rows below it."
      );
      cell.colSpan = Math.max(selectedHeaders.length, 1);
      emptyRow.append(cell);
      tbody.append(emptyRow);
    } else {
      selectedRecords.forEach(function (record, recordIndex) {
        const tr = document.createElement("tr");
        const mergedFlags = selectedMergedRecords[recordIndex] || [];
        const anchorFlags = selectedAnchorRecords[recordIndex] || [];

        record.forEach(function (value, cellIndex) {
          if (mergedFlags[cellIndex] && !anchorFlags[cellIndex]) {
            return; // Skip rendering carried cells, rely on anchor's rowSpan
          }

          const td = createTextCell("td", value || "-");
          if (!value) {
            td.classList.add("cell-empty");
          }

          const anchor = anchorFlags[cellIndex];
          if (anchor) {
            if (anchor.rowSpan > 1) {
              td.rowSpan = anchor.rowSpan;
            }
            td.classList.add("cell-merged-anchor");
          }

          tr.append(td);
        });

        tbody.append(tr);
      });
    }

    elements.parsedTable.append(thead, tbody);
    showTable(elements.parsedEmpty, elements.parsedWrap);
  }

  function buildParsedTable(rows, headerIndex) {
    if (!rows.length) {
      return { headers: [], records: [], mergedCarryRows: [] };
    }

    const safeHeaderIndex = clampHeaderIndex(headerIndex, rows.length);
    const headerRow = rows[safeHeaderIndex] || [];
    const dataRows = rows.slice(safeHeaderIndex + 1);
    const maxColumnCount = Math.max(
      headerRow.length,
      ...dataRows.map(function (row) {
        return row.length;
      }),
      0
    );

    if (!maxColumnCount) {
      return { headers: [], records: [], mergedCarryRows: [] };
    }

    const activeColumns = [];

    for (let columnIndex = 0; columnIndex < maxColumnCount; columnIndex += 1) {
      const headerText = normalizeCell(headerRow[columnIndex]);
      const hasDataBelow = dataRows.some(function (row) {
        return !isEmptyCell(row[columnIndex]);
      });

      if (headerText || hasDataBelow) {
        activeColumns.push(columnIndex);
      }
    }

    if (!activeColumns.length) {
      return { headers: [], records: [], mergedCarryRows: [] };
    }

    const seenHeaders = new Map();
    const headers = activeColumns.map(function (columnIndex) {
      return makeUniqueHeaderName(headerRow[columnIndex], columnIndex, seenHeaders);
    });

    const rowsWithMeta = dataRows.reduce(function (acc, row, dataRowIndex) {
      const hasData = activeColumns.some(function (columnIndex) {
        return !isEmptyCell(row[columnIndex]);
      });

      if (!hasData) {
        return acc;
      }

      const sourceRowIndex = safeHeaderIndex + 1 + dataRowIndex;

      acc.records.push(
        activeColumns.map(function (columnIndex) {
          return formatCellValue(row[columnIndex]);
        })
      );

      acc.mergedCarryRows.push(
        activeColumns.map(function (columnIndex) {
          return isMergedCarryCell(state.currentSheetName, sourceRowIndex, columnIndex);
        })
      );
      
      acc.mergedAnchors.push(
        activeColumns.map(function (columnIndex) {
          return getMergeAnchorCell(state.currentSheetName, sourceRowIndex, columnIndex);
        })
      );

      return acc;
    }, { records: [], mergedCarryRows: [], mergedAnchors: [] });

    return { 
      headers: headers, 
      records: rowsWithMeta.records, 
      mergedCarryRows: rowsWithMeta.mergedCarryRows,
      mergedAnchors: rowsWithMeta.mergedAnchors
    };
  }

  function getSelectedColumnIndexesForCurrentSheet(totalColumns) {
    const safeTotal = Math.max(0, totalColumns);
    if (!safeTotal || !state.currentSheetName) {
      return [];
    }

    const stored = state.columnSelections.get(state.currentSheetName) || [];
    const validStored = stored.filter(function (index) {
      return Number.isInteger(index) && index >= 0 && index < safeTotal;
    });

    if (validStored.length) {
      const limitedStored = validStored.slice(0, MAX_VISIBLE_COLUMNS);
      state.columnSelections.set(state.currentSheetName, limitedStored);
      return limitedStored;
    }

    // Attempt to restore from localStorage based on header signature
    const savedFromStorage = loadColumnSelectionsFromStorage(state.columnPickerHeaders);
    if (savedFromStorage && savedFromStorage.length) {
      const validSaved = savedFromStorage.filter(function(index) {
        return Number.isInteger(index) && index >= 0 && index < safeTotal;
      }).slice(0, MAX_VISIBLE_COLUMNS);
      
      if (validSaved.length) {
        state.columnSelections.set(state.currentSheetName, validSaved);
        return validSaved;
      }
    }

    // Fallback: resolve from COLUMN_PRESETS (sheet-specific, then wildcard)
    var presetNames = resolveColumnPreset(state.currentSheetName);
    if (presetNames && presetNames.length) {
      var presetIndexes = [];
      state.columnPickerHeaders.forEach(function(header, index) {
        if (presetNames.indexOf(header.toLowerCase()) !== -1) {
          presetIndexes.push(index);
        }
      });

      if (presetIndexes.length) {
        // Re-apply the preset order defined by the user in COLUMN_PRESETS
        presetIndexes.sort(function(a, b) {
          return presetNames.indexOf(state.columnPickerHeaders[a].toLowerCase()) -
                 presetNames.indexOf(state.columnPickerHeaders[b].toLowerCase());
        });
        var limitedPreset = presetIndexes.slice(0, MAX_VISIBLE_COLUMNS);
        state.columnSelections.set(state.currentSheetName, limitedPreset);
        return limitedPreset;
      }
    }

    const defaults = Array.from(
      { length: Math.min(MAX_VISIBLE_COLUMNS, safeTotal) },
      function (_, index) {
        return index;
      }
    );
    state.columnSelections.set(state.currentSheetName, defaults);
    return defaults;
  }

  // Resolve the applicable COLUMN_PRESETS entry for a given sheet name.
  // Returns a lowercase-normalised array of column names, or null if no preset matches.
  function resolveColumnPreset(sheetName) {
    if (!COLUMN_PRESETS || typeof COLUMN_PRESETS !== "object") {
      return null;
    }

    // Sheet-specific lookup (case-insensitive key match)
    var sheetKey = Object.keys(COLUMN_PRESETS).find(function(key) {
      return key !== "*" && key.toLowerCase() === (sheetName || "").toLowerCase();
    });

    var names = sheetKey
      ? COLUMN_PRESETS[sheetKey]
      : COLUMN_PRESETS["*"];

    if (!Array.isArray(names) || !names.length) {
      return null;
    }

    return names.map(function(n) { return String(n).toLowerCase(); });
  }

  function getHeaderSignature(headers) {
    if (!headers || !headers.length) return "";
    return headers.join("|");
  }

  function loadColumnSelectionsFromStorage(headers) {
    try {
      const sig = getHeaderSignature(headers);
      if (!sig) return null;
      
      const raw = window.localStorage.getItem(COLUMN_STORAGE_KEY);
      if (!raw) return null;
      
      const data = JSON.parse(raw);
      return data[sig] || null;
    } catch (e) {
      return null;
    }
  }

  function saveColumnSelectionsToStorage(headers, selectedIndexes) {
    try {
      const sig = getHeaderSignature(headers);
      if (!sig) return;
      
      const raw = window.localStorage.getItem(COLUMN_STORAGE_KEY);
      const data = raw ? JSON.parse(raw) : {};
      
      data[sig] = Array.from(selectedIndexes);
      
      // Prevent unbounded growth by keeping only the last 50 signatures
      const keys = Object.keys(data);
      if (keys.length > 50) {
        delete data[keys[0]];
      }
      
      window.localStorage.setItem(COLUMN_STORAGE_KEY, JSON.stringify(data));
    } catch (e) {
      // Ignore if localStorage is unavailable
    }
  }

  function syncColumnPicker(headers, selectedIndexes) {
    state.columnPickerHeaders = headers.slice();

    const countText = selectedIndexes.length + "/" + MAX_VISIBLE_COLUMNS + " selected";
    elements.columnCount.textContent = countText;
    elements.columnsDialogCount.textContent = countText;

    if (!headers.length) {
      elements.columnsOptions.innerHTML = "";
      elements.openColumnsDialogButton.disabled = true;
      elements.resetColumnsButton.disabled = true;
      return;
    }

    renderColumnPickerOptions(headers, selectedIndexes);
    elements.openColumnsDialogButton.disabled = !state.workbook;
    elements.resetColumnsButton.disabled = !state.workbook;
  }

  function renderColumnPickerOptions(headers, selectedIndexes) {
    const selectedSet = new Set(selectedIndexes);
    const fragment = document.createDocumentFragment();

    headers.forEach(function (label, index) {
      const optionLabel = document.createElement("label");
      optionLabel.className = "columns-option";

      const checkbox = document.createElement("input");
      checkbox.type = "checkbox";
      checkbox.value = String(index);
      checkbox.checked = selectedSet.has(index);
      checkbox.disabled = !state.workbook;

      const text = document.createElement("span");
      text.textContent = toSpreadsheetColumn(index) + " · " + label;

      optionLabel.append(checkbox, text);
      fragment.append(optionLabel);
    });

    elements.columnsOptions.innerHTML = "";
    elements.columnsOptions.append(fragment);
    setColumnPickerView(state.columnPickerView);
  }

  function handleResetColumns() {
    if (!state.currentSheetName) {
      return;
    }

    const optionValues = state.columnPickerHeaders.map(function (_, index) {
      return index;
    });

    if (!optionValues.length) {
      return;
    }

    const defaults = optionValues.slice(0, MAX_VISIBLE_COLUMNS);
    state.columnSelections.set(state.currentSheetName, defaults);
    saveColumnSelectionsToStorage(state.columnPickerHeaders, defaults);
    renderSheet(state.currentSheetName);
    showColumnsDialogNote("Reset to first " + defaults.length + " columns.", "success");
  }

  function makeUniqueHeaderName(value, columnIndex, seenHeaders) {
    const baseName = normalizeCell(value) || "Column " + (columnIndex + 1);
    const key = baseName.toLowerCase();
    const count = seenHeaders.get(key) || 0;

    seenHeaders.set(key, count + 1);

    if (!count) {
      return baseName;
    }

    return baseName + " (" + (count + 1) + ")";
  }

  function detectHeaderRow(rows) {
    if (!rows.length) {
      return 0;
    }

    const limit = Math.min(rows.length, 25);
    let bestIndex = 0;
    let bestScore = Number.NEGATIVE_INFINITY;
    let firstNonEmptyIndex = 0;
    let foundNonEmptyRow = false;

    for (let rowIndex = 0; rowIndex < limit; rowIndex += 1) {
      const row = rows[rowIndex] || [];
      const values = row.map(normalizeCell).filter(Boolean);

      if (!values.length) {
        continue;
      }

      if (!foundNonEmptyRow) {
        firstNonEmptyIndex = rowIndex;
        foundNonEmptyRow = true;
      }

      const uniqueCount = new Set(
        values.map(function (value) {
          return value.toLowerCase();
        })
      ).size;
      const textLikeCount = values.filter(isTextLike).length;
      const numericLikeCount = values.filter(isNumericLike).length;
      const nextRows = rows.slice(rowIndex + 1, rowIndex + 4);
      const belowDensity =
        nextRows.reduce(function (sum, nextRow) {
          return sum + countNonEmpty(nextRow);
        }, 0) / Math.max(nextRows.length, 1);
      const supportingRows = nextRows.filter(function (nextRow) {
        return countNonEmpty(nextRow) > 0;
      }).length;

      let score = values.length * 2;
      score += uniqueCount * 1.25;
      score += textLikeCount * 2;
      score += Math.min(belowDensity, values.length);
      score += supportingRows;
      score -= Math.max(0, numericLikeCount - textLikeCount) * 1.5;

      if (values.length === 1) {
        score -= 5;
      }

      if (rowIndex > 0 && countNonEmpty(rows[rowIndex - 1]) === 0) {
        score += 1;
      }

      if (score > bestScore) {
        bestScore = score;
        bestIndex = rowIndex;
      }
    }

    return foundNonEmptyRow ? bestIndex : firstNonEmptyIndex;
  }

  function applyManualHeaderRow(event) {
    if (event) {
      event.preventDefault();
    }

    if (!state.currentSheetName) {
      return;
    }

    const rows = getRowsForSheet(state.currentSheetName);
    const rawValue = elements.headerRowInput.value.trim();
    const requestedRowNumber = Number.parseInt(rawValue, 10);

    if (!rawValue || !/^\d+$/.test(rawValue) || !Number.isFinite(requestedRowNumber)) {
      showStatus("Enter a whole-number header row before applying it.", "error");
      return;
    }

    if (requestedRowNumber < 1) {
      showStatus("Header row numbers start at 1.", "error");
      return;
    }

    applyHeaderSelection(requestedRowNumber - 1);
  }

  function handlePreviewClick(event) {
    const button = event.target.closest(".row-selector");

    if (!button || !state.currentSheetName) {
      return;
    }

    applyHeaderSelection(Number(button.dataset.rowIndex));
  }

  function applyHeaderSelection(headerIndex) {
    if (!state.currentSheetName) {
      return;
    }

    const rows = getRowsForSheet(state.currentSheetName);
    const safeHeaderIndex = clampHeaderIndex(headerIndex, rows.length);

    state.headerSelections.set(state.currentSheetName, safeHeaderIndex);
    renderSheet(state.currentSheetName);
    setActiveView("parsed");
    showStatus(
      "Using row " + (safeHeaderIndex + 1) + " as the header for " + state.currentSheetName + ".",
      "success"
    );
  }

  function toggleTheme() {
    const nextTheme = document.documentElement.dataset.theme === "dark" ? "light" : "dark";
    applyTheme(nextTheme, true);
  }

  function handleTableCellCopyClick(event) {
    if (event.target.closest("button")) {
      return;
    }

    const cell = event.target.closest("td, th");
    if (!cell) {
      return;
    }

    if (cell.querySelector("button")) {
      return;
    }

    if (cell.tagName === "TH" && cell.closest("thead")) {
      const table = cell.closest("table");
      const ths = Array.from(cell.parentElement.children);
      const colIndex = ths.indexOf(cell);
      
      if (colIndex !== -1) {
        const rows = Array.from(table.querySelectorAll("tbody tr"));
        const columnDataText = rows.map(function(row) {
          const td = row.children[colIndex];
          if (!td) return "";
          return td.dataset.formula || (td.classList.contains("cell-empty") ? "" : td.textContent.trim());
        });
        
        const columnDataPlain = columnDataText.join("\n");
        const columnDataHtml = "<table>" + rows.map(function(row) {
          const td = row.children[colIndex];
          if (!td) return "<tr><td></td></tr>";
          const inner = td.querySelector("a") ? td.innerHTML : (td.classList.contains("cell-empty") ? "" : td.textContent.trim());
          return "<tr><td style=\"mso-number-format:'\@'\">" + inner + "</td></tr>";
        }).join("") + "</table>";
        
        copyHtmlToClipboard(columnDataPlain, columnDataHtml)
          .then(function () {
            flashCopiedCell(cell);
            showStatus('Copied entire column', "success");
          })
          .catch(function () {
            showStatus("Copy failed. Please copy manually.", "error");
          });
        return;
      }
    }

    if (event.target.closest("a")) {
      return; // Let the link be clicked without triggering cell copy
    }

    let text = cell.dataset.formula || (cell.classList.contains("cell-empty") ? "" : cell.textContent.trim());
    let innerContent = cell.querySelector("a") ? cell.innerHTML : (cell.classList.contains("cell-empty") ? "" : cell.textContent.trim());
    let html = "<table><tr><td style=\"mso-number-format:'\@'\">" + innerContent + "</td></tr></table>";

    copyHtmlToClipboard(text, html)
      .then(function () {
        flashCopiedCell(cell);
        const preview = text ? text.slice(0, 42) + (text.length > 42 ? "..." : "") : "(empty)";
        showStatus('Copied "' + preview + '"', "success");
      })
      .catch(function () {
        showStatus("Copy failed. Please copy manually.", "error");
      });
  }

  function copyHtmlToClipboard(plainText, htmlText) {
    if (navigator.clipboard && window.ClipboardItem) {
      const typeText = "text/plain";
      const typeHtml = "text/html";
      try {
        const item = new ClipboardItem({
          [typeText]: new Blob([plainText], { type: typeText }),
          [typeHtml]: new Blob([htmlText], { type: typeHtml })
        });
        return navigator.clipboard.write([item]);
      } catch (err) {
        // Fallback for browsers passing error throwing if ClipboardItem isn't fully supported
      }
    }

    // Fallback approach
    return new Promise(function (resolve, reject) {
      try {
        // For HTML, the execCommand approach requires selecting actual HTML elements.
        // It's very difficult to reliably inject custom 'mso-number-format' HTML via execCommand.
        // If Clipboard API fails, we fallback to just writing the plain text so at least they have the data.
        const helper = document.createElement("textarea");
        helper.value = plainText;
        helper.setAttribute("readonly", "");
        helper.style.position = "fixed";
        helper.style.opacity = "0";
        document.body.append(helper);
        helper.select();
        const success = document.execCommand("copy");
        helper.remove();

        if (!success) {
          reject(new Error("Copy command failed"));
          return;
        }

        resolve();
      } catch (error) {
        reject(error);
      }
    });
  }

  function flashCopiedCell(cell) {
    const activeTimer = copyFeedbackTimers.get(cell);
    if (activeTimer) {
      clearTimeout(activeTimer);
    }

    cell.classList.add("is-copied");

    const timer = setTimeout(function () {
      cell.classList.remove("is-copied");
      copyFeedbackTimers.delete(cell);
    }, 650);

    copyFeedbackTimers.set(cell, timer);
  }

  function initializeTheme() {
    const storedTheme = getStoredTheme();
    const initialTheme = storedTheme || (themeMediaQuery.matches ? "dark" : "light");

    applyTheme(initialTheme, false);
  }

  function applyTheme(theme, shouldPersist) {
    const resolvedTheme = theme === "dark" ? "dark" : "light";

    document.documentElement.dataset.theme = resolvedTheme;
    document.documentElement.style.colorScheme = resolvedTheme;
    elements.themeToggle.setAttribute("aria-pressed", String(resolvedTheme === "dark"));
    elements.themeToggleLabel.textContent = resolvedTheme === "dark" ? "Light mode" : "Dark mode";

    if (shouldPersist) {
      setStoredTheme(resolvedTheme);
    }
  }

  function getStoredTheme() {
    try {
      const value = window.localStorage.getItem(THEME_STORAGE_KEY);
      return value === "light" || value === "dark" ? value : "";
    } catch (error) {
      return "";
    }
  }

  function setStoredTheme(theme) {
    try {
      window.localStorage.setItem(THEME_STORAGE_KEY, theme);
    } catch (error) {
      return;
    }
  }

  function clampHeaderIndex(index, rowCount) {
    if (rowCount <= 0) {
      return 0;
    }

    const safeIndex = Number.isFinite(index) ? Math.floor(index) : 0;
    return Math.min(Math.max(safeIndex, 0), rowCount - 1);
  }

  function createTextCell(tagName, text) {
    const cell = document.createElement(tagName);
    
    if (typeof text === 'string') {
      const trimmedText = text.trim();
      const hyperMatch = trimmedText.match(/^=HYPERLINK\(\s*"([^"]+)"\s*(?:,\s*"([^"]+)")?\s*\)$/i);
      
      if (hyperMatch) {
        const url = hyperMatch[1];
        const label = hyperMatch[2] || url;
        
        const link = document.createElement("a");
        link.href = url;
        link.textContent = label;
        link.className = "cell-link";
        link.target = "_blank";
        link.rel = "noopener noreferrer";
        
        // Prevent row selection when clicking the link
        link.addEventListener("click", function(e) {
          e.stopPropagation();
        });
        
        cell.dataset.formula = trimmedText;
        cell.append(link);
        return cell;
      }
    }

    cell.textContent = text;
    return cell;
  }

  function countNonEmpty(row) {
    return (row || []).reduce(function (count, value) {
      return count + (isEmptyCell(value) ? 0 : 1);
    }, 0);
  }

  function isEmptyCell(value) {
    return normalizeCell(value) === "";
  }

  function normalizeCell(value) {
    if (value === null || value === undefined) {
      return "";
    }

    if (value instanceof Date) {
      return value.toISOString().slice(0, 10);
    }

    if (typeof value === "string") {
      return value.trim();
    }

    return String(value).trim();
  }

  function formatCellValue(value) {
    if (value === null || value === undefined || value === "") {
      return "";
    }

    if (value instanceof Date) {
      return value.toLocaleDateString(undefined, {
        year: "numeric",
        month: "short",
        day: "numeric",
      });
    }

    return String(value);
  }

  function isNumericLike(value) {
    const text = typeof value === "string" ? value.trim() : normalizeCell(value);
    return /^[-+]?[$]?\d[\d,]*(\.\d+)?%?$/.test(text);
  }

  function isTextLike(value) {
    return /[A-Za-z]/.test(value);
  }

  function toSpreadsheetColumn(index) {
    let label = "";
    let value = index + 1;

    while (value > 0) {
      const remainder = (value - 1) % 26;
      label = String.fromCharCode(65 + remainder) + label;
      value = Math.floor((value - 1) / 26);
    }

    return label;
  }

  function showTable(emptyElement, tableElement) {
    emptyElement.classList.add("is-hidden");
    tableElement.classList.remove("is-hidden");
  }

  function showEmpty(emptyElement, tableElement) {
    emptyElement.classList.remove("is-hidden");
    tableElement.classList.add("is-hidden");
  }

  function showStatus(message, tone, hideToast) {
    elements.statusNote.textContent = message;
    elements.statusNote.classList.remove("is-error", "is-success");

    if (tone === "error") {
      elements.statusNote.classList.add("is-error");
    }

    if (tone === "success") {
      elements.statusNote.classList.add("is-success");
    }

    if (!hideToast) {
      showToast(message, tone);
    }
  }

  function showToast(message, tone) {
    if (!elements.toastContainer || !message) {
      return;
    }

    const toast = document.createElement("div");
    toast.className = "toast";
    if (tone === "error") {
      toast.classList.add("is-error");
    } else if (tone === "success") {
      toast.classList.add("is-success");
    }

    const id = ++toastTimerSeed;
    toast.dataset.toastId = String(id);
    toast.textContent = message;

    elements.toastContainer.append(toast);

    requestAnimationFrame(function () {
      toast.classList.add("is-visible");
    });

    const timeoutMs = tone === "error" ? 3200 : 2200;
    setTimeout(function () {
      if (!toast.isConnected) {
        return;
      }

      toast.classList.remove("is-visible");
      setTimeout(function () {
        if (toast.isConnected) {
          toast.remove();
        }
      }, 180);
    }, timeoutMs);
  }

  function showColumnsDialogNote(message, tone) {
    elements.columnsDialogNote.textContent = message;
    elements.columnsDialogNote.classList.remove("is-error", "is-success");

    if (tone === "error") {
      elements.columnsDialogNote.classList.add("is-error");
    }

    if (tone === "success") {
      elements.columnsDialogNote.classList.add("is-success");
    }
  }

  function setBusy(isBusy) {
    elements.fileInput.disabled = isBusy;
    elements.sheetSelect.disabled = isBusy || !state.workbook;
    elements.headerRowInput.disabled = isBusy || !state.workbook;
    elements.applyHeaderRowButton.disabled = isBusy || !state.workbook;
    elements.reloadModeSelect.disabled = isBusy || !state.workbook;
    elements.reloadIntervalSelect.disabled = isBusy || !state.workbook || elements.reloadModeSelect.value !== "auto";
    elements.openColumnsDialogButton.disabled =
      isBusy || !state.workbook || !state.columnPickerHeaders.length;
    elements.resetColumnsButton.disabled =
      isBusy || !state.workbook || !state.columnPickerHeaders.length;
    updateFloatingReloadButtonState(isBusy);
  }

  function resetViewer() {
    state.workbook = null;
    state.fileName = "";
    state.fileHandle = null;
    state.currentSheetName = "";
    state.sheetRows.clear();
    state.sheetMergeStats.clear();
    state.sheetMergedCells.clear();
    state.sheetMergeAnchors.clear();
    state.headerSelections.clear();
    state.columnSelections.clear();
    stopAutoReload();
    state.isAutoReloading = false;
    state.autoReloadInterval = 1000;

    elements.sheetSelect.innerHTML = '<option value="">Upload a workbook first</option>';
    elements.sheetSelect.disabled = true;
    elements.headerRowInput.value = "1";
    elements.headerRowInput.disabled = true;
    elements.applyHeaderRowButton.disabled = true;
    elements.reloadModeSelect.value = "manual";
    elements.reloadModeSelect.disabled = true;
    elements.reloadIntervalSelect.value = "1000";
    elements.reloadIntervalSelect.disabled = true;
    state.columnPickerHeaders = [];
    elements.columnsOptions.innerHTML = "";
    elements.columnCount.textContent = "0/" + MAX_VISIBLE_COLUMNS + " selected";
    elements.columnsDialogCount.textContent = "0/" + MAX_VISIBLE_COLUMNS + " selected";
    elements.resetColumnsButton.disabled = true;
    elements.openColumnsDialogButton.disabled = true;
    showColumnsDialogNote("Pick up to " + MAX_VISIBLE_COLUMNS + " columns to display in the parsed table.", "");

    elements.previewTable.innerHTML = "";
    elements.parsedTable.innerHTML = "";
    elements.previewMeta.textContent = "The raw preview will appear here after upload.";
    elements.parsedMeta.textContent = "Choose a header row to build the parsed table.";
    elements.previewEmpty.textContent = "Select a workbook to inspect sheet rows before parsing.";
    elements.parsedEmpty.textContent = "Parsed table data will appear here once a sheet is loaded.";
    showEmpty(elements.previewEmpty, elements.previewWrap);
    showEmpty(elements.parsedEmpty, elements.parsedWrap);

    elements.statWorkbook.textContent = "None loaded";
    elements.statSheet.textContent = "-";
    elements.statHeaderRow.textContent = "-";
    elements.statRecords.textContent = "0";
    setActiveView("parsed");
    setSettingsMenuOpen(false);
    updateFloatingReloadButtonState(false);
  }

  function updateFloatingReloadButtonState(isBusy) {
    const hasWorkbook = Boolean(state.workbook);

    elements.floatingReloadButton.disabled = !hasWorkbook || isBusy;

    if (!hasWorkbook) {
      elements.floatingReloadMeta.textContent = "Load a file first";
      return;
    }

    if (isBusy) {
      elements.floatingReloadMeta.textContent = "Working...";
      return;
    }

    if (state.isAutoReloading) {
      const interval = Number(elements.reloadIntervalSelect.value);
      const intervalText = formatInterval(interval);
      elements.floatingReloadMeta.textContent = "Auto-reload active (" + intervalText + ")";
      return;
    }

    if (state.fileHandle) {
      elements.floatingReloadMeta.textContent = "One-click reload is ready";
      return;
    }

    if (supportsFileSystemAccess()) {
      elements.floatingReloadMeta.textContent = "Click once to connect local reload";
      return;
    }

    elements.floatingReloadMeta.textContent = "Click to re-select the updated file";
  }

  function formatInterval(ms) {
    if (ms < 1000) {
      return ms + "ms";
    } else if (ms < 60000) {
      return Math.floor(ms / 1000) + "s";
    } else {
      return Math.floor(ms / 60000) + "min";
    }
  }

  function supportsFileSystemAccess() {
    return typeof window.showOpenFilePicker === "function";
  }

  initializeTheme();
  resetViewer();
})();

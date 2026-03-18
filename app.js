(function () {
  const PREVIEW_ROW_LIMIT = 24;
  const MAX_VISIBLE_COLUMNS = 10;
  const THEME_STORAGE_KEY = "excel-viewer-theme";
  const COLUMN_STORAGE_KEY = "excel-viewer-columns";
  const MONEY_MAPPING_STORAGE_KEY = "excel-viewer-money-mappings";
  const CHANGELOG_ENTRIES = Object.freeze([
    Object.freeze({
      version: "1.3.0",
      dateLabel: "March 17, 2026",
      stage: "Current release",
      headline: "Cleaner money display, steadier column scrolling, and a built-in update page.",
      summary:
        "This release focuses on everyday workbook review. It makes payment amounts easier to read, keeps long column lists more stable while you browse them, and adds a plain-language place to understand what changed.",
      audience:
        "Operations, finance, QA, and support teammates who need quick spreadsheet review without technical release notes.",
      focus:
        "Clarity first: fewer surprises while scrolling, clearer amount fields, and easier communication for non-technical users.",
      highlights: [
        "Amount columns can automatically show $ for USD and ៛ for KHR when the sheet includes a matching currency column.",
        "The column picker keeps its place more reliably, reducing the jump-back effect in long lists.",
        "A new What's new page explains updates in simple language and shows the app version in a visible place.",
      ],
      impact: [
        "Payment and settlement files are easier to scan because the currency sign is visible at a glance.",
        "Long column lists feel more dependable during review and cleanup work.",
        "Anyone on the team can quickly see what improved without reading code or commit history.",
      ],
      notes: [
        "Your files still stay local in the browser unless you choose to reopen or reload them from your own device.",
        "The parsed table remains the main working view. The changelog page is read-only and there to keep updates easy to understand.",
      ],
    }),
  ]);
  const APP_VERSION = "v" + CHANGELOG_ENTRIES[0].version;
  
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
  const SUPPORTED_CURRENCY_CODES = Object.freeze({
    USD: true,
    KHR: true,
  });
  const currencyFormatterCache = new Map();

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
    floatingChangelogButton: document.getElementById("floating-changelog-button"),
    floatingChangelogMeta: document.getElementById("floating-changelog-meta"),
    openColumnsDialogButton: document.getElementById("floating-columns-button"),
    columnsDialog: document.getElementById("columns-dialog"),
    closeColumnsDialogButton: document.getElementById("close-columns-dialog"),
    columnsOptions: document.getElementById("columns-options"),
    columnCount: document.getElementById("column-count"),
    columnsDialogCount: document.getElementById("columns-dialog-count"),
    columnsDialogNote: document.getElementById("columns-dialog-note"),
    columnsSearchInput: document.getElementById("columns-search-input"),
    columnsSearchMeta: document.getElementById("columns-search-meta"),
    resetColumnsButton: document.getElementById("reset-columns-button"),
    autoDetectMoneyButton: document.getElementById("auto-detect-money-button"),
    clearMoneyMappingButton: document.getElementById("clear-money-mapping-button"),
    senderCurrencySelect: document.getElementById("sender-currency-select"),
    senderAmountSelect: document.getElementById("sender-amount-select"),
    receiverCurrencySelect: document.getElementById("receiver-currency-select"),
    receiverAmountSelect: document.getElementById("receiver-amount-select"),
    moneyMappingNote: document.getElementById("money-mapping-note"),
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
    viewChangelogTab: document.getElementById("view-changelog-tab"),
    previewSection: document.getElementById("preview-section"),
    parsedSection: document.getElementById("parsed-section"),
    changelogSection: document.getElementById("changelog-section"),
    changelogSummary: document.getElementById("changelog-summary"),
    changelogList: document.getElementById("changelog-list"),
    statWorkbook: document.getElementById("stat-workbook"),
    statSheet: document.getElementById("stat-sheet"),
    statHeaderRow: document.getElementById("stat-header-row"),
    statRecords: document.getElementById("stat-records"),
    appVersionBadge: document.getElementById("app-version-badge"),
    footerVersion: document.getElementById("footer-version"),
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
    moneyMappings: new Map(),
    reloadTimer: null,
    isAutoReloading: false,
    autoReloadInterval: 1000,
    activeView: "parsed",
    previewCopyContext: null,
    parsedCopyContext: null,
    columnPickerHeaders: [],
    columnPickerSelectedIndexes: [],
    columnPickerSearch: "",
    columnPickerScrollTop: 0,
    restoreColumnPickerScroll: false,
    columnPickerFocusValue: "",
    restoreColumnPickerFocus: false,
    columnPickerView: "list",
  };

  const copyFeedbackTimers = new WeakMap();
  let toastTimerSeed = 0;
  let lockedPageScrollTop = 0;

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
  elements.floatingChangelogButton.addEventListener("click", handleOpenChangelogView);
  elements.openColumnsDialogButton.addEventListener("click", openColumnsDialog);
  elements.columnsDialog.addEventListener("close", unlockPageScroll);
  elements.columnsDialog.addEventListener("cancel", unlockPageScroll);
  elements.columnsOptions.addEventListener("change", handleColumnOptionToggle);
  elements.columnsOptions.addEventListener("scroll", handleColumnsOptionsScroll, {
    passive: true,
  });
  elements.columnsSearchInput.addEventListener("input", handleColumnSearchInput);
  elements.columnsSearchInput.addEventListener("keydown", handleColumnSearchKeydown);
  elements.resetColumnsButton.addEventListener("click", handleResetColumns);
  elements.autoDetectMoneyButton.addEventListener("click", handleAutoDetectMoneyMapping);
  elements.clearMoneyMappingButton.addEventListener("click", handleClearMoneyMapping);
  elements.senderCurrencySelect.addEventListener("change", handleMoneyMappingChange);
  elements.senderAmountSelect.addEventListener("change", handleMoneyMappingChange);
  elements.receiverCurrencySelect.addEventListener("change", handleMoneyMappingChange);
  elements.receiverAmountSelect.addEventListener("change", handleMoneyMappingChange);
  elements.viewPreviewTab.addEventListener("click", handleViewTabClick);
  elements.viewParsedTab.addEventListener("click", handleViewTabClick);
  elements.viewChangelogTab.addEventListener("click", handleViewTabClick);
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
    const safeView =
      view === "preview" || view === "changelog" ? view : "parsed";
    state.activeView = safeView;

    const isPreview = safeView === "preview";
    const isParsed = safeView === "parsed";
    const isChangelog = safeView === "changelog";
    elements.previewSection.classList.toggle("is-hidden", !isPreview);
    elements.parsedSection.classList.toggle("is-hidden", !isParsed);
    elements.changelogSection.classList.toggle("is-hidden", !isChangelog);
    elements.viewPreviewTab.classList.toggle("is-active", isPreview);
    elements.viewParsedTab.classList.toggle("is-active", isParsed);
    elements.viewChangelogTab.classList.toggle("is-active", isChangelog);
    elements.viewPreviewTab.setAttribute("aria-pressed", String(isPreview));
    elements.viewParsedTab.setAttribute("aria-pressed", String(isParsed));
    elements.viewChangelogTab.setAttribute("aria-pressed", String(isChangelog));
  }

  function handleOpenChangelogView() {
    setSettingsMenuOpen(false);
    setActiveView("changelog");
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
    const savedMoneyMappings = preserveHeaderSelections
      ? new Map(state.moneyMappings)
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
      state.moneyMappings.clear();
      savedMoneyMappings.forEach(function (mapping, signature) {
        if (signature && mapping) {
          state.moneyMappings.set(signature, mapping);
        }
      });

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
      state.previewCopyContext = null;
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

    const previewHeaders = ["Row"];
    for (let columnIndex = 0; columnIndex < columnCount; columnIndex += 1) {
      previewHeaders.push(toSpreadsheetColumn(columnIndex));
    }

    state.previewCopyContext = {
      headers: previewHeaders,
      records: visibleRows.map(function (row, visibleRowIndex) {
        const actualRowIndex = windowStart + visibleRowIndex;
        const record = ["Row " + (actualRowIndex + 1)];

        for (let columnIndex = 0; columnIndex < columnCount; columnIndex += 1) {
          record.push(formatCellValue(row[columnIndex]));
        }

        return record;
      }),
    };

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
    resetColumnPickerSearchState();

    if (typeof elements.columnsDialog.showModal === "function") {
      lockPageScroll();
      showColumnsDialogNote("Search the list below and select up to " + MAX_VISIBLE_COLUMNS + " columns.", "");
      elements.columnsDialog.showModal();
      window.requestAnimationFrame(function () {
        elements.columnsSearchInput.focus();
        elements.columnsSearchInput.select();
      });
    }
  }

  function handleColumnSearchInput(event) {
    state.columnPickerSearch = normalizeColumnPickerSearch(event.target.value);
    state.columnPickerScrollTop = 0;
    renderColumnPickerSurface(
      state.columnPickerHeaders,
      getSelectedColumnIndexesForCurrentSheet(state.columnPickerHeaders),
      { resetScroll: true }
    );
  }

  function handleColumnSearchKeydown(event) {
    if (event.key !== "Escape" || !elements.columnsSearchInput.value) {
      return;
    }

    elements.columnsSearchInput.value = "";
    state.columnPickerSearch = "";
    state.columnPickerScrollTop = 0;
    renderColumnPickerSurface(
      state.columnPickerHeaders,
      getSelectedColumnIndexesForCurrentSheet(state.columnPickerHeaders),
      { resetScroll: true }
    );
  }

  function resetColumnPickerSearchState() {
    state.columnPickerSearch = "";
    state.columnPickerScrollTop = 0;
    elements.columnsSearchInput.value = "";
    renderColumnPickerSurface(
      state.columnPickerHeaders,
      getSelectedColumnIndexesForCurrentSheet(state.columnPickerHeaders),
      { resetScroll: true }
    );
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

    const current = getSelectedColumnIndexesForCurrentSheet(state.columnPickerHeaders);
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
    state.columnPickerScrollTop = elements.columnsOptions.scrollTop;
    state.restoreColumnPickerScroll = true;
    state.columnPickerFocusValue = String(toggledIndex);
    state.restoreColumnPickerFocus = true;
    state.columnSelections.set(state.currentSheetName, nextSelection);
    saveColumnSelectionsToStorage(state.columnPickerHeaders, nextSelection);
    renderSheet(state.currentSheetName);
    showColumnsDialogNote("Applied " + nextSelection.length + " selected columns.", "success");
  }

  function handleColumnsOptionsScroll() {
    state.columnPickerScrollTop = elements.columnsOptions.scrollTop;
  }

  function renderParsedTable(rows, headerIndex, mergeSummary) {
    const parsed = buildParsedTable(rows, headerIndex);

    elements.statWorkbook.textContent = state.fileName || "None loaded";
    elements.statSheet.textContent = state.currentSheetName || "-";
    elements.statHeaderRow.textContent = rows.length ? "Row " + (headerIndex + 1) : "-";
    elements.statRecords.textContent = String(parsed.records.length);

    if (!parsed.headers.length) {
      state.parsedCopyContext = null;
      syncColumnPicker([], []);
      syncMoneyMappingControls([], null);
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

    syncMoneyMappingControls(parsed.headers, parsed.moneyMappingState);
    const selectedColumnIndexes = getSelectedColumnIndexesForCurrentSheet(parsed.headers);
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

    state.parsedCopyContext = {
      headers: selectedHeaders.slice(),
      records: selectedRecords.map(function (record) {
        return record.slice();
      }),
    };

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
      return { headers: [], records: [], mergedCarryRows: [], moneyMappingState: null };
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
      return { headers: [], records: [], mergedCarryRows: [], moneyMappingState: null };
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
      return { headers: [], records: [], mergedCarryRows: [], moneyMappingState: null };
    }

    const seenHeaders = new Map();
    const headers = activeColumns.map(function (columnIndex) {
      return makeUniqueHeaderName(headerRow[columnIndex], columnIndex, seenHeaders);
    });
    const moneyMappingState = resolveMoneyMappingForHeaders(headers, dataRows);
    const amountCurrencyColumns = buildAmountCurrencyColumnLookup(headers, moneyMappingState.mapping);

    const rowsWithMeta = dataRows.reduce(function (acc, row, dataRowIndex) {
      const hasData = activeColumns.some(function (columnIndex) {
        return !isEmptyCell(row[columnIndex]);
      });

      if (!hasData) {
        return acc;
      }

      const sourceRowIndex = safeHeaderIndex + 1 + dataRowIndex;

      acc.records.push(
        activeColumns.map(function (columnIndex, activeHeaderIndex) {
          return formatCellValue(row[columnIndex], {
            activeColumns: activeColumns,
            activeHeaderIndex: activeHeaderIndex,
            amountCurrencyColumns: amountCurrencyColumns,
            headers: headers,
            row: row,
            sourceColumnIndex: columnIndex,
            moneyMapping: moneyMappingState.mapping,
          });
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
      mergedAnchors: rowsWithMeta.mergedAnchors,
      moneyMappingState: moneyMappingState,
    };
  }

  function getSelectedColumnIndexesForCurrentSheet(headers) {
    const activeHeaders = Array.isArray(headers) ? headers.slice() : state.columnPickerHeaders.slice();
    const safeTotal = Math.max(0, activeHeaders.length);
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
    const savedFromStorage = loadColumnSelectionsFromStorage(activeHeaders);
    if (savedFromStorage && savedFromStorage.length) {
      const validSaved = savedFromStorage.filter(function(index) {
        return Number.isInteger(index) && index >= 0 && index < safeTotal;
      }).slice(0, MAX_VISIBLE_COLUMNS);
      
      if (validSaved.length) {
        state.columnSelections.set(state.currentSheetName, validSaved);
        return validSaved;
      }
    }

    const defaults = resolveDefaultColumnSelection(activeHeaders, state.currentSheetName);
    state.columnSelections.set(state.currentSheetName, defaults.indexes);
    return defaults.indexes;
  }

  function resolveDefaultColumnSelection(headers, sheetName) {
    const activeHeaders = Array.isArray(headers) ? headers.slice() : [];
    const safeTotal = Math.max(0, activeHeaders.length);

    if (!safeTotal) {
      return { indexes: [], source: "empty" };
    }

    // Resolve from COLUMN_PRESETS first (sheet-specific, then wildcard).
    var presetNames = resolveColumnPreset(sheetName);
    if (presetNames && presetNames.length) {
      var presetIndexes = [];
      activeHeaders.forEach(function(header, index) {
        if (presetNames.indexOf(String(header).toLowerCase()) !== -1) {
          presetIndexes.push(index);
        }
      });

      if (presetIndexes.length) {
        // Re-apply the preset order defined by the user in COLUMN_PRESETS.
        presetIndexes.sort(function(a, b) {
          return presetNames.indexOf(String(activeHeaders[a]).toLowerCase()) -
                 presetNames.indexOf(String(activeHeaders[b]).toLowerCase());
        });

        return {
          indexes: presetIndexes.slice(0, MAX_VISIBLE_COLUMNS),
          source: "preset"
        };
      }
    }

    return {
      indexes: Array.from(
        { length: Math.min(MAX_VISIBLE_COLUMNS, safeTotal) },
        function (_, index) {
          return index;
        }
      ),
      source: "first-columns"
    };
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
    const previousHeaders = state.columnPickerHeaders.slice();
    const previousSelectedIndexes = state.columnPickerSelectedIndexes.slice();
    const preserveOpenDialogScroll =
      elements.columnsDialog.open &&
      arraysEqual(previousHeaders, headers) &&
      arraysEqual(previousSelectedIndexes, selectedIndexes);

    state.columnPickerHeaders = headers.slice();
    state.columnPickerSelectedIndexes = selectedIndexes.slice();

    const countText = selectedIndexes.length + "/" + MAX_VISIBLE_COLUMNS + " selected";
    elements.columnCount.textContent = countText;
    elements.columnsDialogCount.textContent = countText;

    if (!headers.length) {
      elements.columnsOptions.innerHTML = "";
      elements.columnsSearchMeta.textContent = "Showing 0 columns";
      elements.openColumnsDialogButton.disabled = true;
      elements.resetColumnsButton.disabled = true;
      return;
    }

    renderColumnPickerSurface(headers, selectedIndexes, {
      preserveScroll: state.restoreColumnPickerScroll || preserveOpenDialogScroll,
      focusValue: state.restoreColumnPickerFocus ? state.columnPickerFocusValue : "",
    });
    state.restoreColumnPickerScroll = false;
    state.restoreColumnPickerFocus = false;
    elements.openColumnsDialogButton.disabled = !state.workbook;
    elements.resetColumnsButton.disabled = !state.workbook;
  }

  function setColumnPickerView(view) {
    state.columnPickerView = view === "grid" ? "grid" : "list";
    elements.columnsOptions.dataset.view = state.columnPickerView;
  }

  function syncMoneyMappingControls(headers, mappingState) {
    const activeHeaders = Array.isArray(headers) ? headers.slice() : [];
    const mapping = mappingState && mappingState.mapping ? mappingState.mapping : null;
    const source = mappingState && mappingState.source ? mappingState.source : "auto";
    const countText = activeHeaders.length ? activeHeaders.length + " columns available" : "No columns available";

    if (!activeHeaders.length) {
      elements.senderCurrencySelect.innerHTML = "";
      elements.senderAmountSelect.innerHTML = "";
      elements.receiverCurrencySelect.innerHTML = "";
      elements.receiverAmountSelect.innerHTML = "";
      elements.autoDetectMoneyButton.disabled = true;
      elements.clearMoneyMappingButton.disabled = true;
      elements.senderCurrencySelect.disabled = true;
      elements.senderAmountSelect.disabled = true;
      elements.receiverCurrencySelect.disabled = true;
      elements.receiverAmountSelect.disabled = true;
      elements.moneyMappingNote.textContent = "Map sender and receiver amount columns to format money values.";
      return;
    }

    const previousSenderCurrency = elements.senderCurrencySelect.value;
    const previousSenderAmount = elements.senderAmountSelect.value;
    const previousReceiverCurrency = elements.receiverCurrencySelect.value;
    const previousReceiverAmount = elements.receiverAmountSelect.value;

    const optionsHtml = buildMoneyMappingOptions(activeHeaders);
    elements.senderCurrencySelect.innerHTML = optionsHtml;
    elements.senderAmountSelect.innerHTML = optionsHtml;
    elements.receiverCurrencySelect.innerHTML = optionsHtml;
    elements.receiverAmountSelect.innerHTML = optionsHtml;

    applyMoneyMappingValue(elements.senderCurrencySelect, mapping && mapping.senderCurrency);
    applyMoneyMappingValue(elements.senderAmountSelect, mapping && mapping.senderAmount);
    applyMoneyMappingValue(elements.receiverCurrencySelect, mapping && mapping.receiverCurrency);
    applyMoneyMappingValue(elements.receiverAmountSelect, mapping && mapping.receiverAmount);

    if (mappingState && mappingState.preserveCurrentValues) {
      elements.senderCurrencySelect.value = previousSenderCurrency;
      elements.senderAmountSelect.value = previousSenderAmount;
      elements.receiverCurrencySelect.value = previousReceiverCurrency;
      elements.receiverAmountSelect.value = previousReceiverAmount;
    }

    const hasAnyManualMapping =
      Number.isInteger(mapping && mapping.senderCurrency) ||
      Number.isInteger(mapping && mapping.senderAmount) ||
      Number.isInteger(mapping && mapping.receiverCurrency) ||
      Number.isInteger(mapping && mapping.receiverAmount);

    elements.autoDetectMoneyButton.disabled = !state.workbook || !activeHeaders.length;
    elements.clearMoneyMappingButton.disabled = !state.workbook || !hasAnyManualMapping;
    elements.senderCurrencySelect.disabled = !state.workbook;
    elements.senderAmountSelect.disabled = !state.workbook;
    elements.receiverCurrencySelect.disabled = !state.workbook;
    elements.receiverAmountSelect.disabled = !state.workbook;

    if (source === "manual") {
      elements.moneyMappingNote.textContent =
        "Using your saved money mapping for this column set. " + countText + ".";
    } else if (mapping && (mapping.senderAmount !== null || mapping.receiverAmount !== null)) {
      elements.moneyMappingNote.textContent =
        "Auto-detected money columns from the sheet. Adjust them if needed. " + countText + ".";
    } else {
      elements.moneyMappingNote.textContent =
        "Map sender and receiver amount columns to format money values. " + countText + ".";
    }
  }

  function buildMoneyMappingOptions(headers) {
    const fragment = document.createDocumentFragment();
    const placeholder = document.createElement("option");
    placeholder.value = "";
    placeholder.textContent = "Not mapped";
    fragment.append(placeholder);

    headers.forEach(function (label, index) {
      const option = document.createElement("option");
      option.value = String(index);
      option.textContent = toSpreadsheetColumn(index) + " - " + (label || "Column " + (index + 1));
      fragment.append(option);
    });

    const wrapper = document.createElement("div");
    wrapper.append(fragment);
    return wrapper.innerHTML;
  }

  function applyMoneyMappingValue(select, index) {
    select.value = Number.isInteger(index) && index >= 0 ? String(index) : "";
  }

  function handleMoneyMappingChange() {
    if (!state.currentSheetName || !state.columnPickerHeaders.length) {
      return;
    }

    const mapping = readMoneyMappingFromControls();
    const signature = getHeaderSignature(state.columnPickerHeaders);

    if (!signature) {
      return;
    }

    if (hasAnyMoneyMapping(mapping)) {
      state.moneyMappings.set(signature, mapping);
      saveMoneyMappingToStorage(state.columnPickerHeaders, mapping);
    } else {
      state.moneyMappings.delete(signature);
      removeMoneyMappingFromStorage(state.columnPickerHeaders);
    }

    renderSheet(state.currentSheetName);
    showColumnsDialogNote(
      hasAnyMoneyMapping(mapping)
        ? "Saved money mapping for this column set."
        : "Cleared the money mapping for this column set.",
      hasAnyMoneyMapping(mapping) ? "success" : ""
    );
  }

  function handleAutoDetectMoneyMapping() {
    if (!state.currentSheetName || !state.columnPickerHeaders.length) {
      return;
    }

    const rows = getRowsForSheet(state.currentSheetName);
    const parsed = detectMoneyMapping(state.columnPickerHeaders, rows);
    const signature = getHeaderSignature(state.columnPickerHeaders);

    if (!signature) {
      return;
    }

    state.moneyMappings.set(signature, parsed);
    saveMoneyMappingToStorage(state.columnPickerHeaders, parsed);
    renderSheet(state.currentSheetName);
    showColumnsDialogNote("Applied auto-detected money mapping. Review it before continuing.", "success");
  }

  function handleClearMoneyMapping() {
    if (!state.currentSheetName || !state.columnPickerHeaders.length) {
      return;
    }

    const signature = getHeaderSignature(state.columnPickerHeaders);
    if (!signature) {
      return;
    }

    state.moneyMappings.delete(signature);
    removeMoneyMappingFromStorage(state.columnPickerHeaders);
    renderSheet(state.currentSheetName);
    showColumnsDialogNote("Cleared saved money mapping. Auto-detection is active again.", "");
  }

  function readMoneyMappingFromControls() {
    return normalizeMoneyMapping({
      senderCurrency: parseMoneyMappingSelectValue(elements.senderCurrencySelect.value),
      senderAmount: parseMoneyMappingSelectValue(elements.senderAmountSelect.value),
      receiverCurrency: parseMoneyMappingSelectValue(elements.receiverCurrencySelect.value),
      receiverAmount: parseMoneyMappingSelectValue(elements.receiverAmountSelect.value),
      source: "manual",
    }, state.columnPickerHeaders.length);
  }

  function parseMoneyMappingSelectValue(value) {
    if (value === "" || value === null || value === undefined) {
      return null;
    }

    const parsed = Number.parseInt(String(value), 10);
    return Number.isInteger(parsed) && parsed >= 0 ? parsed : null;
  }

  function hasAnyMoneyMapping(mapping) {
    return Boolean(
      mapping &&
        (
          Number.isInteger(mapping.senderCurrency) ||
          Number.isInteger(mapping.senderAmount) ||
          Number.isInteger(mapping.receiverCurrency) ||
          Number.isInteger(mapping.receiverAmount)
        )
    );
  }

  function getMoneyMappingSignature(headers) {
    return getHeaderSignature(headers);
  }

  function loadMoneyMappingFromStorage(headers) {
    try {
      const signature = getMoneyMappingSignature(headers);
      if (!signature) {
        return null;
      }

      const raw = window.localStorage.getItem(MONEY_MAPPING_STORAGE_KEY);
      if (!raw) {
        return null;
      }

      const data = JSON.parse(raw);
      return normalizeMoneyMapping(data && data[signature], headers.length);
    } catch (error) {
      return null;
    }
  }

  function saveMoneyMappingToStorage(headers, mapping) {
    try {
      const signature = getMoneyMappingSignature(headers);
      if (!signature) {
        return;
      }

      const raw = window.localStorage.getItem(MONEY_MAPPING_STORAGE_KEY);
      const data = raw ? JSON.parse(raw) : {};
      data[signature] = {
        senderCurrency: Number.isInteger(mapping.senderCurrency) ? mapping.senderCurrency : null,
        senderAmount: Number.isInteger(mapping.senderAmount) ? mapping.senderAmount : null,
        receiverCurrency: Number.isInteger(mapping.receiverCurrency) ? mapping.receiverCurrency : null,
        receiverAmount: Number.isInteger(mapping.receiverAmount) ? mapping.receiverAmount : null,
      };
      window.localStorage.setItem(MONEY_MAPPING_STORAGE_KEY, JSON.stringify(data));
    } catch (error) {
      return;
    }
  }

  function removeMoneyMappingFromStorage(headers) {
    try {
      const signature = getMoneyMappingSignature(headers);
      if (!signature) {
        return;
      }

      const raw = window.localStorage.getItem(MONEY_MAPPING_STORAGE_KEY);
      if (!raw) {
        return;
      }

      const data = JSON.parse(raw);
      delete data[signature];
      window.localStorage.setItem(MONEY_MAPPING_STORAGE_KEY, JSON.stringify(data));
    } catch (error) {
      return;
    }
  }

  function normalizeMoneyMapping(mapping, totalColumns) {
    const safeTotal = Number.isInteger(totalColumns) && totalColumns > 0 ? totalColumns : 0;
    const normalized = {
      senderCurrency: normalizeMoneyMappingIndex(mapping && mapping.senderCurrency, safeTotal),
      senderAmount: normalizeMoneyMappingIndex(mapping && mapping.senderAmount, safeTotal),
      receiverCurrency: normalizeMoneyMappingIndex(mapping && mapping.receiverCurrency, safeTotal),
      receiverAmount: normalizeMoneyMappingIndex(mapping && mapping.receiverAmount, safeTotal),
      detectedAmountColumns: Array.isArray(mapping && mapping.detectedAmountColumns)
        ? mapping.detectedAmountColumns.filter(function (index) {
            return Number.isInteger(index) && index >= 0 && (safeTotal === 0 || index < safeTotal);
          })
        : [],
      detectedCurrencyColumns: Array.isArray(mapping && mapping.detectedCurrencyColumns)
        ? mapping.detectedCurrencyColumns.filter(function (index) {
            return Number.isInteger(index) && index >= 0 && (safeTotal === 0 || index < safeTotal);
          })
        : [],
      source: mapping && mapping.source === "manual" ? "manual" : "auto",
    };

    return normalized;
  }

  function normalizeMoneyMappingIndex(value, totalColumns) {
    if (!Number.isInteger(value) || value < 0) {
      return null;
    }

    if (Number.isInteger(totalColumns) && totalColumns > 0 && value >= totalColumns) {
      return null;
    }

    return value;
  }

  function resolveMoneyMappingForHeaders(headers, rows) {
    const activeHeaders = Array.isArray(headers) ? headers.slice() : [];
    const detected = detectMoneyMapping(activeHeaders, rows);
    const signature = getMoneyMappingSignature(activeHeaders);
    const stored = signature ? state.moneyMappings.get(signature) : null;
    const saved = stored || loadMoneyMappingFromStorage(activeHeaders);

    if (saved) {
      return {
        mapping: mergeMoneyMappings(detected, saved),
        source: "manual",
      };
    }

    return {
      mapping: detected,
      source: "auto",
    };
  }

  function mergeMoneyMappings(baseMapping, overrideMapping) {
    const base = normalizeMoneyMapping(baseMapping, 0);
    const override = normalizeMoneyMapping(overrideMapping, 0);

    return {
      senderCurrency: override.senderCurrency !== null ? override.senderCurrency : base.senderCurrency,
      senderAmount: override.senderAmount !== null ? override.senderAmount : base.senderAmount,
      receiverCurrency: override.receiverCurrency !== null ? override.receiverCurrency : base.receiverCurrency,
      receiverAmount: override.receiverAmount !== null ? override.receiverAmount : base.receiverAmount,
      detectedAmountColumns: base.detectedAmountColumns,
      detectedCurrencyColumns: base.detectedCurrencyColumns,
      source: override.source === "manual" ? "manual" : base.source || "auto",
    };
  }

  function detectMoneyMapping(headers, rows) {
    const activeHeaders = Array.isArray(headers) ? headers : [];
    const sampleRows = Array.isArray(rows) ? rows.slice(0, 40) : [];

    const stats = activeHeaders.map(function (header, index) {
      const values = sampleRows
        .map(function (row) {
          return row ? row[index] : null;
        })
        .filter(function (value) {
          return !isEmptyCell(value);
        });

      const nonEmptyCount = values.length;
      const numericLikeCount = values.filter(function (value) {
        return parseCurrencyAmount(value) !== null || isNumericLike(value);
      }).length;
      const currencyCodeCount = values.filter(function (value) {
        return Boolean(getSupportedCurrencyCode(value));
      }).length;
      const currencySymbolCount = values.filter(function (value) {
        return hasCurrencySymbol(value);
      }).length;
      const headerText = normalizeHeaderLabel(header);
      const side = getTransferSide(headerText);
      const accountLikeHeader = isAccountLikeHeader(headerText);

      return {
        index: index,
        amountScore:
          scoreMoneyAmountHeader(headerText) +
          (nonEmptyCount ? numericLikeCount / nonEmptyCount : 0) * 6 +
          currencySymbolCount * 2 -
          (accountLikeHeader ? 6 : 0),
        currencyScore:
          scoreMoneyCurrencyHeader(headerText) +
          (nonEmptyCount ? currencyCodeCount / nonEmptyCount : 0) * 8 +
          currencySymbolCount * 3 -
          (accountLikeHeader ? 8 : 0),
        side: side,
      };
    });

    const amountCandidates = stats
      .filter(function (item) {
        return item.amountScore >= 2.2;
      })
      .sort(function (left, right) {
        return right.amountScore - left.amountScore || left.index - right.index;
      });
    const currencyCandidates = stats
      .filter(function (item) {
        return item.currencyScore >= 2.2;
      })
      .sort(function (left, right) {
        return right.currencyScore - left.currencyScore || left.index - right.index;
      });

    const detectedAmountColumns = amountCandidates.map(function (item) {
      return item.index;
    });
    const detectedCurrencyColumns = currencyCandidates.map(function (item) {
      return item.index;
    });

    const senderAmount = pickBestMoneyCandidate(amountCandidates, "sender", "amount", []);
    const receiverAmount =
      amountCandidates.length > 1
        ? pickBestMoneyCandidate(
            amountCandidates,
            "receiver",
            "amount",
            senderAmount !== null ? [senderAmount] : []
          )
        : null;
    const senderCurrency = pickBestMoneyCandidate(currencyCandidates, "sender", "currency", []);
    const receiverCurrency = pickBestMoneyCandidate(
      currencyCandidates,
      "receiver",
      "currency",
      currencyCandidates.length > 1 && senderCurrency !== null ? [senderCurrency] : []
    );

    return normalizeMoneyMapping(
      {
        senderCurrency: senderCurrency,
        senderAmount: senderAmount,
        receiverCurrency: receiverCurrency,
        receiverAmount: receiverAmount,
        detectedAmountColumns: detectedAmountColumns,
        detectedCurrencyColumns: detectedCurrencyColumns,
        source: "auto",
      },
      activeHeaders.length
    );
  }

  function pickBestMoneyCandidate(candidates, side, type, excludedIndexes) {
    if (!Array.isArray(candidates) || !candidates.length) {
      return null;
    }

    const exclusions = new Set(Array.isArray(excludedIndexes) ? excludedIndexes : []);
    const scored = candidates
      .map(function (candidate) {
        const sideBonus = candidate.side === side ? 2.5 : candidate.side ? 0.5 : 0;
        return {
          index: candidate.index,
          score: (type === "currency" ? candidate.currencyScore : candidate.amountScore) + sideBonus,
        };
      })
      .filter(function (candidate) {
        return !exclusions.has(candidate.index);
      })
      .sort(function (left, right) {
        return right.score - left.score || left.index - right.index;
      });

    return scored.length ? scored[0].index : null;
  }

  function scoreMoneyAmountHeader(headerText) {
    let score = 0;

    if (/\b(amount|amt|value|total|net|gross|balance|debit|credit|fee|charge|payment|payout|settlement|principal|price|cost)\b/.test(headerText)) {
      score += 4;
    }

    if (/\b(sender|receiver|recipient|beneficiary|payer|payee|from|to)\b/.test(headerText)) {
      score += 1;
    }

    return score;
  }

  function scoreMoneyCurrencyHeader(headerText) {
    let score = 0;

    if (/\b(currency|curr|ccy|fx)\b/.test(headerText)) {
      score += 5;
    }

    if (/\b(usd|khr)\b/.test(headerText)) {
      score += 3;
    }

    return score;
  }

  function isAccountLikeHeader(headerText) {
    return /\b(account|acct|iban|wallet|number|no|accountno|accountnumber|beneficiaryaccount|receiveraccount|senderaccount)\b/.test(
      headerText
    );
  }

  function getTransferSide(headerText) {
    if (/\b(sender|from|source|debit|payer|origin)\b/.test(headerText)) {
      return "sender";
    }

    if (/\b(receiver|to|destination|credit|beneficiary|recipient|payee)\b/.test(headerText)) {
      return "receiver";
    }

    return "";
  }

  function hasCurrencySymbol(value) {
    const text = normalizeCell(value);
    return /[$៛]/.test(text);
  }

  function renderColumnPickerSurface(headers, selectedIndexes, options) {
    const settings = options || {};

    renderSimpleColumnPickerOptions(headers, selectedIndexes);

    const scrollTop = settings.preserveScroll ? state.columnPickerScrollTop : 0;
    const focusValue = settings.focusValue || "";

    window.requestAnimationFrame(function () {
      elements.columnsOptions.scrollTop =
        settings.resetScroll || !settings.preserveScroll
          ? 0
          : scrollTop;

      if (!focusValue) {
        return;
      }

      const checkbox = elements.columnsOptions.querySelector(
        'input[type="checkbox"][value="' + focusValue + '"]'
      );
      if (checkbox && typeof checkbox.focus === "function") {
        try {
          checkbox.focus({ preventScroll: true });
        } catch (error) {
          checkbox.focus();
        }
      }
    });
  }

  function updateColumnPickerFilterButtons() {
    elements.columnsFilterAllButton.classList.toggle("is-active", state.columnPickerFilter === "all");
    elements.columnsFilterSelectedButton.classList.toggle("is-active", state.columnPickerFilter === "selected");
    elements.columnsFilterHiddenButton.classList.toggle("is-active", state.columnPickerFilter === "hidden");
    elements.columnsFilterAllButton.setAttribute("aria-pressed", String(state.columnPickerFilter === "all"));
    elements.columnsFilterSelectedButton.setAttribute("aria-pressed", String(state.columnPickerFilter === "selected"));
    elements.columnsFilterHiddenButton.setAttribute("aria-pressed", String(state.columnPickerFilter === "hidden"));
  }

  function renderSelectedColumnsSummary(headers, selectedIndexes) {
    const fragment = document.createDocumentFragment();
    const label = document.createElement("span");
    label.className = "columns-selected-label";
    label.textContent = "Selected columns";
    fragment.append(label);

    if (!selectedIndexes.length) {
      const empty = document.createElement("span");
      empty.className = "columns-selected-empty";
      empty.textContent = "None selected";
      fragment.append(empty);
      elements.columnsSelectedSummary.innerHTML = "";
      elements.columnsSelectedSummary.append(fragment);
      return;
    }

    selectedIndexes.forEach(function (index) {
      const chip = document.createElement("span");
      chip.className = "column-chip";
      chip.title = headers[index] || "";

      const key = document.createElement("span");
      key.className = "column-chip-key";
      key.textContent = toSpreadsheetColumn(index);

      const name = document.createElement("span");
      name.className = "column-chip-name";
      name.textContent = headers[index] || ("Column " + (index + 1));

      chip.append(key, name);
      fragment.append(chip);
    });

    elements.columnsSelectedSummary.innerHTML = "";
    elements.columnsSelectedSummary.append(fragment);
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

  function renderFilteredColumnPickerOptions(headers, selectedIndexes) {
    const selectedSet = new Set(selectedIndexes);
    const filteredOptions = getFilteredColumnOptions(headers, selectedSet);
    const fragment = document.createDocumentFragment();

    if (!filteredOptions.length) {
      const empty = document.createElement("div");
      empty.className = "columns-options-empty";
      empty.textContent =
        state.columnPickerSearch || state.columnPickerFilter !== "all"
          ? "No columns match the current search or filter."
          : "No columns are available.";

      elements.columnsOptions.replaceChildren(empty);
      renderColumnPickerSearchMeta(headers.length, 0, 0);
      setColumnPickerView(state.columnPickerView);
      return;
    }

    filteredOptions.forEach(function (item) {
      const optionLabel = document.createElement("label");
      optionLabel.className = "columns-option";
      optionLabel.classList.toggle("is-selected", item.selected);

      const checkbox = document.createElement("input");
      checkbox.type = "checkbox";
      checkbox.value = String(item.index);
      checkbox.checked = item.selected;
      checkbox.disabled = !state.workbook;

      const key = document.createElement("span");
      key.className = "columns-option-key";
      key.textContent = item.key;

      const textWrap = document.createElement("span");
      textWrap.className = "columns-option-text";

      const name = document.createElement("span");
      name.className = "columns-option-name";
      name.textContent = item.label;

      const meta = document.createElement("span");
      meta.className = "columns-option-meta";
      meta.textContent = item.selected ? "Selected" : "Available";

      textWrap.append(name, meta);
      optionLabel.append(checkbox, key, textWrap);
      fragment.append(optionLabel);
    });

    elements.columnsOptions.replaceChildren(fragment);
    renderColumnPickerSearchMeta(
      headers.length,
      filteredOptions.length,
      filteredOptions.filter(function (item) {
        return item.selected;
      }).length
    );
    setColumnPickerView(state.columnPickerView);
  }

  function getFilteredColumnOptions(headers, selectedSet) {
    const query = state.columnPickerSearch;

    return headers.reduce(function (items, label, index) {
      const key = toSpreadsheetColumn(index);
      const isSelected = selectedSet.has(index);
      const matchesFilter =
        state.columnPickerFilter === "all" ||
        (state.columnPickerFilter === "selected" && isSelected) ||
        (state.columnPickerFilter === "hidden" && !isSelected);

      if (!matchesFilter) {
        return items;
      }

      const haystack = (key + " " + String(label || "")).toLowerCase();
      if (query && haystack.indexOf(query) === -1) {
        return items;
      }

      items.push({
        index: index,
        key: key,
        label: label || ("Column " + (index + 1)),
        selected: isSelected,
      });
      return items;
    }, []);
  }

  function renderColumnPickerSearchMeta(totalColumns, visibleColumns, visibleSelectedColumns) {
    let text = visibleColumns + " of " + totalColumns + " columns";

    if (visibleSelectedColumns) {
      text += " | " + visibleSelectedColumns + " selected";
    }

    elements.columnsSearchMeta.textContent = text;
  }

  function normalizeColumnPickerSearch(value) {
    return String(value || "")
      .trim()
      .toLowerCase()
      .replace(/\s+/g, " ");
  }

  function arraysEqual(left, right) {
    if (!Array.isArray(left) || !Array.isArray(right) || left.length !== right.length) {
      return false;
    }

    for (let index = 0; index < left.length; index += 1) {
      if (left[index] !== right[index]) {
        return false;
      }
    }

    return true;
  }

  function renderSimpleColumnPickerOptions(headers, selectedIndexes) {
    const selectedSet = new Set(selectedIndexes);
    const visibleOptions = getVisibleColumnOptions(headers, selectedSet);
    const fragment = document.createDocumentFragment();

    if (!visibleOptions.length) {
      const empty = document.createElement("div");
      empty.className = "columns-options-empty";
      empty.textContent = state.columnPickerSearch
        ? "No columns match your search."
        : "No columns are available.";

      elements.columnsOptions.replaceChildren(empty);
      renderColumnPickerSearchMeta(headers.length, 0, 0);
      return;
    }

    visibleOptions.forEach(function (item) {
      const optionLabel = document.createElement("label");
      optionLabel.className = "columns-option";
      optionLabel.classList.toggle("is-selected", item.selected);

      const checkbox = document.createElement("input");
      checkbox.type = "checkbox";
      checkbox.value = String(item.index);
      checkbox.checked = item.selected;
      checkbox.disabled = !state.workbook;

      const key = document.createElement("span");
      key.className = "columns-option-key";
      key.textContent = item.key;

      const name = document.createElement("span");
      name.className = "columns-option-name";
      name.textContent = item.label;

      optionLabel.append(checkbox, key, name);
      fragment.append(optionLabel);
    });

    elements.columnsOptions.replaceChildren(fragment);
    renderColumnPickerSearchMeta(
      headers.length,
      visibleOptions.length,
      visibleOptions.filter(function (item) {
        return item.selected;
      }).length
    );
  }

  function getVisibleColumnOptions(headers, selectedSet) {
    const query = state.columnPickerSearch;

    return headers.reduce(function (items, label, index) {
      const key = toSpreadsheetColumn(index);
      const haystack = (key + " " + String(label || "")).toLowerCase();

      if (query && haystack.indexOf(query) === -1) {
        return items;
      }

      items.push({
        index: index,
        key: key,
        label: label || ("Column " + (index + 1)),
        selected: selectedSet.has(index),
      });
      return items;
    }, []).sort(function (a, b) {
      if (a.selected !== b.selected) {
        return a.selected ? -1 : 1;
      }

      return a.index - b.index;
    });
  }

  function lockPageScroll() {
    lockedPageScrollTop =
      window.scrollY ||
      window.pageYOffset ||
      document.documentElement.scrollTop ||
      0;

    document.body.style.position = "fixed";
    document.body.style.top = "-" + lockedPageScrollTop + "px";
    document.body.style.left = "0";
    document.body.style.right = "0";
    document.body.style.width = "100%";
    document.body.style.overflow = "hidden";
  }

  function unlockPageScroll() {
    const top = document.body.style.top;

    document.body.style.position = "";
    document.body.style.top = "";
    document.body.style.left = "";
    document.body.style.right = "";
    document.body.style.width = "";
    document.body.style.overflow = "";

    const scrollTop = top ? Math.abs(Number.parseInt(top, 10)) || lockedPageScrollTop : lockedPageScrollTop;
    window.scrollTo(0, scrollTop);
    lockedPageScrollTop = 0;
  }

  function handleResetColumns() {
    if (!state.currentSheetName) {
      return;
    }

    if (!state.columnPickerHeaders.length) {
      return;
    }

    const defaults = resolveDefaultColumnSelection(
      state.columnPickerHeaders,
      state.currentSheetName
    );
    state.columnSelections.set(state.currentSheetName, defaults.indexes);
    saveColumnSelectionsToStorage(state.columnPickerHeaders, defaults.indexes);
    renderSheet(state.currentSheetName);
    showColumnsDialogNote(
      defaults.source === "preset"
        ? "Reset to configured default columns."
        : "Reset to first " + defaults.indexes.length + " columns.",
      "success"
    );
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
        const copyContext = getTableCopyContext(table);
        if (!copyContext || !Array.isArray(copyContext.records) || colIndex >= copyContext.headers.length) {
          return;
        }

        if (!copyContext.records.length) {
          showStatus("No rows to copy from this column.", "error");
          return;
        }

        const columnDataText = copyContext.records.map(function (row) {
          return getClipboardCellText(row[colIndex]);
        });
        const columnDataPlain = columnDataText.join("\n");
        const columnDataHtml = buildClipboardTableHtml(
          columnDataText.map(function (value) {
            return [value];
          })
        );

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
    let html = buildClipboardTableHtml([[text]]);

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

  function getTableCopyContext(table) {
    if (table === elements.previewTable) {
      return state.previewCopyContext;
    }

    if (table === elements.parsedTable) {
      return state.parsedCopyContext;
    }

    return null;
  }

  function getClipboardCellText(value) {
    if (value === null || value === undefined) {
      return "";
    }

    return String(value);
  }

  function buildClipboardTableHtml(rows) {
    return (
      "<table>" +
      rows
        .map(function (row) {
          return (
            "<tr>" +
            row
              .map(function (value) {
                return (
                  "<td style=\"mso-number-format:'\\@'\">" +
                  getClipboardCellHtml(value) +
                  "</td>"
                );
              })
              .join("") +
            "</tr>"
          );
        })
        .join("") +
      "</table>"
    );
  }

  function getClipboardCellHtml(value) {
    const text = getClipboardCellText(value);
    const hyperMatch = text
      .trim()
      .match(/^=HYPERLINK\(\s*"([^"]+)"\s*(?:,\s*"([^"]+)")?\s*\)$/i);

    if (hyperMatch) {
      const url = escapeHtml(hyperMatch[1]);
      const label = escapeHtml(hyperMatch[2] || hyperMatch[1]);
      return "<a href=\"" + url + "\">" + label + "</a>";
    }

    return escapeHtml(text);
  }

  function escapeHtml(text) {
    return String(text)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/\"/g, "&quot;")
      .replace(/'/g, "&#39;");
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

  function formatCellValue(value, options) {
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

    if (options) {
      const amountText = formatAmountCellValue(value, options);
      if (amountText !== null) {
        return amountText;
      }
    }

    return String(value);
  }

  function formatAmountCellValue(value, options) {
    if (!options || !Array.isArray(options.headers)) {
      return null;
    }

    const header = options.headers[options.activeHeaderIndex] || "";
    if (!isMoneyAmountColumn(options.activeHeaderIndex, header, options.moneyMapping)) {
      return null;
    }

    const currencyCode = resolveAmountCurrencyCode(value, options);
    if (!currencyCode) {
      return null;
    }

    const numericValue = parseCurrencyAmount(value);
    if (numericValue === null) {
      return null;
    }

    const fractionDigits = resolveCurrencyFractionDigits(value);
    return getCurrencyFormatter(currencyCode, fractionDigits).format(numericValue);
  }

  function resolveAmountCurrencyCode(value, options) {
    const pairedCurrencyCode = resolvePairedCurrencyCode(options);
    if (pairedCurrencyCode) {
      return pairedCurrencyCode;
    }

    const headerCurrencyCode = getSupportedCurrencyCode(options.headers[options.activeHeaderIndex]);
    if (headerCurrencyCode) {
      return headerCurrencyCode;
    }

    return getSupportedCurrencyCode(value);
  }

  function resolvePairedCurrencyCode(options) {
    if (
      !options ||
      !Array.isArray(options.amountCurrencyColumns) ||
      !Array.isArray(options.activeColumns) ||
      !Array.isArray(options.row)
    ) {
      return "";
    }

    const currencyHeaderIndex = options.amountCurrencyColumns[options.activeHeaderIndex];
    if (!Number.isInteger(currencyHeaderIndex) || currencyHeaderIndex < 0) {
      return "";
    }

    const sourceColumnIndex = options.activeColumns[currencyHeaderIndex];
    if (!Number.isInteger(sourceColumnIndex) || sourceColumnIndex < 0) {
      return "";
    }

    return getSupportedCurrencyCode(options.row[sourceColumnIndex]);
  }

  function getSupportedCurrencyCode(value) {
    const normalized = normalizeCell(value).toUpperCase();
    const matches = normalized.match(/\b(?:USD|KHR)\b/);
    if (matches && SUPPORTED_CURRENCY_CODES[matches[0]]) {
      return matches[0];
    }

    if (normalized.indexOf("$") !== -1) {
      return "USD";
    }

    if (normalized.indexOf("៛") !== -1) {
      return "KHR";
    }

    if (!matches || !SUPPORTED_CURRENCY_CODES[matches[0]]) {
      return "";
    }

    return matches[0];
  }

  function parseCurrencyAmount(value) {
    if (typeof value === "number" && Number.isFinite(value)) {
      return value;
    }

    const text = normalizeCell(value);
    if (!text) {
      return null;
    }

    const normalized = text
      .replace(/^\((.*)\)$/, "-$1")
      .replace(/\b(?:USD|KHR)\b/gi, "")
      .replace(/[$៛,\s]/g, "");

    if (!/^[-+]?(?:\d+\.?\d*|\.\d+)$/.test(normalized)) {
      return null;
    }

    const parsed = Number(normalized);
    return Number.isFinite(parsed) ? parsed : null;
  }

  function resolveCurrencyFractionDigits(value) {
    const text = normalizeCell(value).replace(/,/g, "");
    const fractionMatch = text.match(/\.(\d+)/);
    return fractionMatch ? Math.min(fractionMatch[1].length, 6) : 0;
  }

  function getCurrencyFormatter(currencyCode, fractionDigits) {
    const safeFractionDigits =
      Number.isInteger(fractionDigits) && fractionDigits >= 0 ? fractionDigits : 0;
    const cacheKey = currencyCode + ":" + safeFractionDigits;

    if (!currencyFormatterCache.has(cacheKey)) {
      currencyFormatterCache.set(
        cacheKey,
        new Intl.NumberFormat(undefined, {
          currency: currencyCode,
          currencyDisplay: "narrowSymbol",
          maximumFractionDigits: safeFractionDigits,
          minimumFractionDigits: safeFractionDigits,
          style: "currency",
        })
      );
    }

    return currencyFormatterCache.get(cacheKey);
  }

  function buildAmountCurrencyColumnLookup(headers, moneyMapping) {
    const safeHeaders = Array.isArray(headers) ? headers.slice() : [];
    const mapping = normalizeMoneyMapping(moneyMapping, safeHeaders.length);
    const currencyIndexes = safeHeaders.reduce(function (indexes, header, index) {
      if (isMoneyCurrencyColumn(index, header, mapping)) {
        indexes.push(index);
      }

      return indexes;
    }, []);
    const lookup = safeHeaders.map(function () {
      return -1;
    });
    const assignedAmountIndexes = new Set();

    function assignAmount(amountIndex, currencyIndex) {
      if (!Number.isInteger(amountIndex) || amountIndex < 0 || amountIndex >= lookup.length) {
        return;
      }

      lookup[amountIndex] = Number.isInteger(currencyIndex) ? currencyIndex : -1;
      assignedAmountIndexes.add(amountIndex);
    }

    function nearestCurrencyIndex(amountIndex) {
      if (!currencyIndexes.length) {
        return -1;
      }

      let bestIndex = currencyIndexes[0];
      let bestDistance = Number.POSITIVE_INFINITY;

      currencyIndexes.forEach(function (currencyIndex) {
        const distance = Math.abs(currencyIndex - amountIndex);
        if (distance < bestDistance) {
          bestDistance = distance;
          bestIndex = currencyIndex;
        }
      });

      return bestIndex;
    }

    [
      [mapping.senderAmount, mapping.senderCurrency],
      [mapping.receiverAmount, mapping.receiverCurrency],
    ].forEach(function (pair) {
      const amountIndex = pair[0];
      const currencyIndex = pair[1];
      if (Number.isInteger(amountIndex) && amountIndex >= 0) {
        assignAmount(amountIndex, Number.isInteger(currencyIndex) ? currencyIndex : nearestCurrencyIndex(amountIndex));
      }
    });

    safeHeaders.forEach(function (header, index) {
      if (assignedAmountIndexes.has(index) || !isMoneyAmountColumn(index, header, mapping)) {
        return;
      }

      assignAmount(index, nearestCurrencyIndex(index));
    });

    return lookup;
  }

  function isAmountColumnHeader(header) {
    return /\b(amount|amt|value|total|net|gross|balance|debit|credit|fee|charge|payment|payout|settlement|principal|price|cost)\b/.test(
      normalizeHeaderLabel(header)
    );
  }

  function isCurrencyColumnHeader(header) {
    return /\b(currency|ccy|curr|fx)\b/.test(normalizeHeaderLabel(header));
  }

  function isMoneyAmountColumn(index, header, moneyMapping) {
    const mapping = normalizeMoneyMapping(moneyMapping, 0);
    return (
      isAmountColumnHeader(header) ||
      index === mapping.senderAmount ||
      index === mapping.receiverAmount ||
      (Array.isArray(mapping.detectedAmountColumns) && mapping.detectedAmountColumns.indexOf(index) !== -1)
    );
  }

  function isMoneyCurrencyColumn(index, header, moneyMapping) {
    const mapping = normalizeMoneyMapping(moneyMapping, 0);
    return (
      isCurrencyColumnHeader(header) ||
      index === mapping.senderCurrency ||
      index === mapping.receiverCurrency ||
      (Array.isArray(mapping.detectedCurrencyColumns) && mapping.detectedCurrencyColumns.indexOf(index) !== -1)
    );
  }

  function getHeaderSemanticBase(header, type) {
    const pattern =
      type === "currency" ? /\b(currency|ccy|curr|fx)\b/g : /\b(amount|amt|value|total|net|gross|balance|debit|credit|fee|charge|payment|payout|settlement|principal|price|cost)\b/g;

    return normalizeHeaderLabel(header).replace(pattern, " ").replace(/\s+/g, " ").trim();
  }

  function normalizeHeaderLabel(header) {
    return normalizeCell(header)
      .toLowerCase()
      .replace(/[_-]+/g, " ")
      .replace(/[^a-z0-9\s]/g, " ")
      .replace(/\s+/g, " ")
      .trim();
  }

  function initializeAppMeta() {
    if (elements.appVersionBadge) {
      elements.appVersionBadge.textContent = APP_VERSION;
    }

    if (elements.footerVersion) {
      elements.footerVersion.textContent = APP_VERSION;
    }

    if (elements.floatingChangelogMeta) {
      elements.floatingChangelogMeta.textContent = "See what is new in " + APP_VERSION;
    }

    renderChangelog();
  }

  function renderChangelog() {
    if (!elements.changelogSummary || !elements.changelogList || !CHANGELOG_ENTRIES.length) {
      return;
    }

    const latest = CHANGELOG_ENTRIES[0];
    const summaryFragment = document.createDocumentFragment();
    summaryFragment.append(
      buildChangelogSummaryCard(
        latest.stage,
        "Version " + APP_VERSION,
        latest.headline,
        latest.summary,
        true
      ),
      buildChangelogSummaryCard(
        "Best for",
        "Who this update helps",
        latest.audience,
        "",
        false
      ),
      buildChangelogSummaryCard(
        "Release focus",
        "Why this update exists",
        latest.focus,
        "",
        false
      )
    );
    elements.changelogSummary.replaceChildren(summaryFragment);

    const listFragment = document.createDocumentFragment();
    CHANGELOG_ENTRIES.forEach(function (entry) {
      listFragment.append(buildChangelogEntry(entry));
    });
    elements.changelogList.replaceChildren(listFragment);
  }

  function buildChangelogSummaryCard(label, title, body, supportingText, isPrimary) {
    const card = document.createElement("article");
    card.className = "changelog-card" + (isPrimary ? " is-primary" : "");

    const labelElement = document.createElement("p");
    labelElement.className = "changelog-card-label";
    labelElement.textContent = label;

    const titleElement = document.createElement("h3");
    titleElement.className = "changelog-card-title";
    titleElement.textContent = title;

    const bodyElement = document.createElement("p");
    bodyElement.className = "changelog-card-body";
    bodyElement.textContent = body;

    card.append(labelElement, titleElement, bodyElement);

    if (supportingText) {
      const supportingElement = document.createElement("p");
      supportingElement.className = "changelog-card-supporting";
      supportingElement.textContent = supportingText;
      card.append(supportingElement);
    }

    return card;
  }

  function buildChangelogEntry(entry) {
    const article = document.createElement("article");
    article.className = "changelog-entry";

    const header = document.createElement("header");
    header.className = "changelog-entry-header";

    const headerCopy = document.createElement("div");
    headerCopy.className = "changelog-entry-copy";

    const versionLabel = document.createElement("p");
    versionLabel.className = "changelog-entry-version";
    versionLabel.textContent = "Version " + entry.version;

    const headline = document.createElement("h3");
    headline.className = "changelog-entry-title";
    headline.textContent = entry.headline;

    const date = document.createElement("p");
    date.className = "changelog-entry-date";
    date.textContent = entry.dateLabel;

    headerCopy.append(versionLabel, headline, date);

    const stage = document.createElement("span");
    stage.className = "changelog-entry-stage";
    stage.textContent = entry.stage;

    header.append(headerCopy, stage);

    const summary = document.createElement("p");
    summary.className = "changelog-entry-summary";
    summary.textContent = entry.summary;

    const sections = document.createElement("div");
    sections.className = "changelog-entry-sections";
    sections.append(
      buildChangelogListSection("What changed", entry.highlights),
      buildChangelogListSection("Why it matters", entry.impact),
      buildChangelogListSection("Good to know", entry.notes)
    );

    article.append(header, summary, sections);
    return article;
  }

  function buildChangelogListSection(title, items) {
    const section = document.createElement("section");
    section.className = "changelog-entry-section";

    const heading = document.createElement("h4");
    heading.className = "changelog-entry-section-title";
    heading.textContent = title;

    const list = document.createElement("ul");
    list.className = "changelog-points";

    items.forEach(function (item) {
      const listItem = document.createElement("li");
      listItem.textContent = item;
      list.append(listItem);
    });

    section.append(heading, list);
    return section;
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
    elements.autoDetectMoneyButton.disabled =
      isBusy || !state.workbook || !state.columnPickerHeaders.length;
    elements.clearMoneyMappingButton.disabled =
      isBusy || !state.workbook || !state.columnPickerHeaders.length;
    elements.senderCurrencySelect.disabled =
      isBusy || !state.workbook || !state.columnPickerHeaders.length;
    elements.senderAmountSelect.disabled =
      isBusy || !state.workbook || !state.columnPickerHeaders.length;
    elements.receiverCurrencySelect.disabled =
      isBusy || !state.workbook || !state.columnPickerHeaders.length;
    elements.receiverAmountSelect.disabled =
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
    state.moneyMappings.clear();
    stopAutoReload();
    state.isAutoReloading = false;
    state.autoReloadInterval = 1000;
    state.previewCopyContext = null;
    state.parsedCopyContext = null;
    state.columnPickerSelectedIndexes = [];
    state.columnPickerSearch = "";
    state.columnPickerScrollTop = 0;
    state.restoreColumnPickerScroll = false;
    state.columnPickerFocusValue = "";
    state.restoreColumnPickerFocus = false;
    state.columnPickerView = "list";

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
    elements.columnsSearchInput.value = "";
    elements.columnsSearchMeta.textContent = "Showing 0 columns";
    elements.columnCount.textContent = "0/" + MAX_VISIBLE_COLUMNS + " selected";
    elements.columnsDialogCount.textContent = "0/" + MAX_VISIBLE_COLUMNS + " selected";
    elements.resetColumnsButton.disabled = true;
    elements.openColumnsDialogButton.disabled = true;
    elements.autoDetectMoneyButton.disabled = true;
    elements.clearMoneyMappingButton.disabled = true;
    elements.senderCurrencySelect.disabled = true;
    elements.senderAmountSelect.disabled = true;
    elements.receiverCurrencySelect.disabled = true;
    elements.receiverAmountSelect.disabled = true;
    elements.senderCurrencySelect.innerHTML = "";
    elements.senderAmountSelect.innerHTML = "";
    elements.receiverCurrencySelect.innerHTML = "";
    elements.receiverAmountSelect.innerHTML = "";
    elements.moneyMappingNote.textContent = "Map sender and receiver amount columns to format money values.";
    showColumnsDialogNote("Search the list below and select up to " + MAX_VISIBLE_COLUMNS + " columns.", "");

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

  initializeAppMeta();
  initializeTheme();
  resetViewer();
})();

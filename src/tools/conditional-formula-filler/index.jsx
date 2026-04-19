import { useState, useCallback, useMemo, useEffect, useRef } from 'react';
import XLSX from 'xlsx-js-style';
import { patchXlsxCells, isPatchSupported } from './utils/xlsxPatch.js';
import './conditional-formula-filler.css';

// ── constants ────────────────────────────────────────────────
const STORAGE_KEY = 'conditional-formula-filler:v1';
const IDB_DB = 'cff-store';
const IDB_STORE = 'workbooks';
const STEP_LABELS = ['Upload', 'Lookup Values', 'Conditions', 'Targets', 'Preview', 'Export'];
const MAX_PREVIEW = 20;
const IDB_THRESHOLD = 4 * 1024 * 1024;

// ── utilities ────────────────────────────────────────────────

function uid() {
  return Math.random().toString(36).slice(2, 10);
}

function detectHeader(rows) {
  const limit = Math.min(20, rows.length);
  let best = 0, bestScore = -1;
  for (let i = 0; i < limit; i++) {
    const row = rows[i] || [];
    const uniq = new Set(
      row
        .filter(c => c != null && String(c).trim() !== '' && isNaN(Number(String(c).trim())))
        .map(c => String(c).trim())
    );
    if (uniq.size > bestScore) { bestScore = uniq.size; best = i; }
  }
  return best;
}

function sheetToRows(sheet) {
  if (!sheet || !sheet['!ref']) return [];
  const range = XLSX.utils.decode_range(sheet['!ref']);
  const rows = Array.from({ length: range.e.r + 1 }, () =>
    new Array(range.e.c + 1).fill('')
  );
  for (const [addr, cell] of Object.entries(sheet)) {
    if (addr.startsWith('!')) continue;
    const pos = XLSX.utils.decode_cell(addr);
    if (pos.r <= range.e.r && pos.c <= range.e.c && cell?.v != null) {
      rows[pos.r][pos.c] =
        cell.w !== undefined ? cell.w
        : cell.v instanceof Date ? cell.v.toLocaleDateString()
        : String(cell.v);
    }
  }
  return rows;
}

function arrToB64(buffer) {
  const bytes = new Uint8Array(buffer);
  let bin = '';
  for (let i = 0; i < bytes.length; i += 8192) {
    bin += String.fromCharCode(...bytes.subarray(i, i + 8192));
  }
  return btoa(bin);
}

function b64ToArr(b64) {
  const bin = atob(b64);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  return bytes.buffer;
}

function tsFilename(name) {
  const base = name.replace(/\.(xlsx?|xls|csv)$/i, '');
  const now = new Date();
  const p = n => String(n).padStart(2, '0');
  return `${base}_filled_${now.getFullYear()}${p(now.getMonth()+1)}${p(now.getDate())}_${p(now.getHours())}${p(now.getMinutes())}.xlsx`;
}

// ── IndexedDB helpers ────────────────────────────────────────

function openIdb() {
  return new Promise((res, rej) => {
    const req = indexedDB.open(IDB_DB, 1);
    req.onupgradeneeded = e => e.target.result.createObjectStore(IDB_STORE);
    req.onsuccess = e => res(e.target.result);
    req.onerror = rej;
  });
}
async function idbPut(key, val) {
  const db = await openIdb();
  return new Promise((res, rej) => {
    const tx = db.transaction(IDB_STORE, 'readwrite');
    tx.objectStore(IDB_STORE).put(val, key);
    tx.oncomplete = res; tx.onerror = rej;
  });
}
async function idbGet(key) {
  const db = await openIdb();
  return new Promise((res, rej) => {
    const tx = db.transaction(IDB_STORE, 'readonly');
    const req = tx.objectStore(IDB_STORE).get(key);
    req.onsuccess = e => res(e.target.result); req.onerror = rej;
  });
}
async function idbDel(key) {
  const db = await openIdb();
  return new Promise((res, rej) => {
    const tx = db.transaction(IDB_STORE, 'readwrite');
    tx.objectStore(IDB_STORE).delete(key);
    tx.oncomplete = res; tx.onerror = rej;
  });
}

// ── fallback export (xlsx-js-style round-trip) ───────────────

async function _exportFallback(workbook, selectedSheet, headers, fillMap, exportMode) {
  let outWb;
  if (exportMode === 'workbook') {
    outWb = { SheetNames: [...workbook.SheetNames], Sheets: {} };
    for (const sn of workbook.SheetNames) {
      outWb.Sheets[sn] = sn === selectedSheet
        ? applyFillsToSheet(workbook.Sheets[sn], headers, fillMap)
        : workbook.Sheets[sn];
    }
  } else {
    outWb = { SheetNames: [selectedSheet], Sheets: {} };
    outWb.Sheets[selectedSheet] = applyFillsToSheet(workbook.Sheets[selectedSheet], headers, fillMap);
  }
  const buf = XLSX.write(outWb, { bookType: 'xlsx', type: 'array', cellStyles: true });
  return buf instanceof ArrayBuffer ? buf : new Uint8Array(buf).buffer;
}

// ── sheet fill ───────────────────────────────────────────────

function applyFillsToSheet(sheet, headers, fillMap) {
  const out = {};
  for (const [k, v] of Object.entries(sheet)) {
    if (k.startsWith('!')) {
      out[k] = Array.isArray(v)
        ? v.map(x => (typeof x === 'object' && x ? { ...x } : x))
        : (typeof v === 'object' && v ? { ...v } : v);
    } else {
      out[k] = typeof v === 'object' && v !== null ? { ...v } : v;
    }
  }
  for (const [rIdxStr, fills] of Object.entries(fillMap)) {
    const r = Number(rIdxStr);
    for (const [colName, value] of Object.entries(fills)) {
      const c = headers.indexOf(colName);
      if (c === -1) continue;
      const addr = XLSX.utils.encode_cell({ r, c });
      const existing = out[addr];
      const numVal = Number(value);
      if (existing?.t === 'n' && !isNaN(numVal) && value !== '') {
        out[addr] = { ...existing, v: numVal, w: String(value) };
      } else {
        out[addr] = { t: 's', v: String(value), ...(existing?.s ? { s: existing.s } : {}) };
      }
    }
  }
  return out;
}

// ── main component ───────────────────────────────────────────

export default function ConditionalFormulaFiller() {
  const [step, setStep] = useState(0);
  const [maxStep, setMaxStep] = useState(0);

  // file state
  const [file, setFile] = useState(null);
  const [fileBuffer, setFileBuffer] = useState(null);
  const [workbook, setWorkbook] = useState(null);
  const [selectedSheet, setSelectedSheet] = useState('');
  const [rows, setRows] = useState([]);
  const [detectedHeader, setDetectedHeader] = useState(0);
  const [headerRowIndex, setHeaderRowIndex] = useState(0);
  const [headerInput, setHeaderInput] = useState('1');

  // step 1
  const [lookupEntries, setLookupEntries] = useState([]);
  // [{id, label, value}]

  // step 2
  const [conditionColumns, setConditionColumns] = useState([]);
  const [ruleConditions, setRuleConditions] = useState({});
  // {entryId: {colName: value}}

  // step 3
  const [targetColumns, setTargetColumns] = useState([]);
  const [targetMode, setTargetMode] = useState('same');
  const [perTargetValues, setPerTargetValues] = useState({});
  // {targetCol: {entryId: value}}

  // step 5
  const [exportFilename, setExportFilename] = useState('');
  const [exportMode, setExportMode] = useState('workbook');
  const [isExporting, setIsExporting] = useState(false);
  const [exportDone, setExportDone] = useState(false);
  const [exportedFile, setExportedFile] = useState('');

  // ui
  const [isBusy, setIsBusy] = useState(false);
  const [busyMsg, setBusyMsg] = useState('');
  const [error, setError] = useState('');
  const [warnMsg, setWarnMsg] = useState('');

  const restoringRef = useRef(false);
  const saveTimer = useRef(null);

  // ── derived ──────────────────────────────────────────────

  const headers = useMemo(() => {
    if (!rows.length) return [];
    return (rows[headerRowIndex] || []).map(c => (c != null ? String(c).trim() : '')).filter(Boolean);
  }, [rows, headerRowIndex]);

  const headerPreview = useMemo(() => {
    const idx = parseInt(headerInput, 10) - 1;
    if (isNaN(idx) || idx < 0 || idx >= rows.length) return [];
    return (rows[idx] || []).map(c => (c != null ? String(c).trim() : '')).filter(Boolean);
  }, [rows, headerInput]);

  const colUniqueValues = useMemo(() => {
    if (!headers.length || !rows.length || !conditionColumns.length) return {};
    const result = {};
    for (const col of conditionColumns) {
      const ci = headers.indexOf(col);
      if (ci === -1) continue;
      const vals = new Set();
      for (let r = headerRowIndex + 1; r < rows.length; r++) {
        const v = rows[r]?.[ci];
        if (v != null && String(v).trim()) vals.add(String(v).trim());
      }
      result[col] = [...vals].sort();
    }
    return result;
  }, [headers, rows, headerRowIndex, conditionColumns]);

  const fillResults = useMemo(() => {
    if (!rows.length || !lookupEntries.length || !conditionColumns.length || !targetColumns.length) return [];
    const colIndices = Object.fromEntries(headers.map((h, i) => [h, i]));
    const results = [];
    for (let r = headerRowIndex + 1; r < rows.length; r++) {
      const row = rows[r];
      if (!row) continue;
      let matched = null;
      for (const entry of lookupEntries) {
        const conds = ruleConditions[entry.id] || {};
        if (!conditionColumns.every(col => conds[col])) continue;
        const ok = conditionColumns.every(col => {
          const ci = colIndices[col];
          return ci !== undefined && String(row[ci] ?? '').trim() === conds[col];
        });
        if (ok) { matched = entry; break; }
      }
      if (!matched) continue;
      const fills = {};
      for (const tCol of targetColumns) {
        const val = targetMode === 'same'
          ? matched.value
          : (perTargetValues?.[tCol]?.[matched.id] ?? '');
        if (val !== '' && val != null) fills[tCol] = val;
      }
      if (Object.keys(fills).length > 0) results.push({ rowIndex: r, fills });
    }
    return results;
  }, [rows, headerRowIndex, headers, lookupEntries, conditionColumns, ruleConditions, targetColumns, targetMode, perTargetValues]);

  const fillStats = useMemo(() => {
    const total = Math.max(0, rows.length - headerRowIndex - 1);
    const filled = new Set(fillResults.map(r => r.rowIndex)).size;
    return { total, filled, empty: total - filled };
  }, [rows, headerRowIndex, fillResults]);

  const previewRows = useMemo(() => {
    if (!rows.length || !headers.length) return [];
    const fillMap = Object.fromEntries(fillResults.map(r => [r.rowIndex, r.fills]));
    const out = [];
    for (let r = headerRowIndex + 1; r < rows.length && out.length < MAX_PREVIEW; r++) {
      out.push({ rowIndex: r, cells: rows[r] || [], fills: fillMap[r] || {} });
    }
    return out;
  }, [rows, headerRowIndex, headers, fillResults]);

  const conflictingRules = useMemo(() => {
    if (!conditionColumns.length || lookupEntries.length < 2) return [];
    const seen = new Map();
    const conflicts = [];
    for (const entry of lookupEntries) {
      const conds = ruleConditions[entry.id] || {};
      const key = conditionColumns.map(col => `${col}=${conds[col] ?? ''}`).join('|');
      if (seen.has(key) && conditionColumns.every(col => conds[col])) {
        conflicts.push(`"${entry.label}" conflicts with "${seen.get(key)}"`);
      } else {
        seen.set(key, entry.label);
      }
    }
    return conflicts;
  }, [conditionColumns, lookupEntries, ruleConditions]);

  // validation
  const dupLabels = useMemo(() => {
    const labels = lookupEntries.map(e => e.label.trim()).filter(Boolean);
    return labels.filter((l, i) => labels.indexOf(l) !== i);
  }, [lookupEntries]);

  const step1Valid = lookupEntries.length > 0
    && lookupEntries.every(e => e.label.trim() && e.value.trim())
    && dupLabels.length === 0;

  const step2Valid = conditionColumns.length > 0
    && lookupEntries.every(e => conditionColumns.every(col => (ruleConditions[e.id] || {})[col]))
    && conflictingRules.length === 0;

  const step3Valid = targetColumns.length > 0 && (
    targetMode === 'same'
    || lookupEntries.every(e => targetColumns.every(t => (perTargetValues[t] || {})[e.id]?.trim()))
  );

  // ── persistence ──────────────────────────────────────────

  const doSave = useCallback(async () => {
    if (restoringRef.current) return;
    const cfg = {
      step, maxStep, selectedSheet, headerRowIndex,
      lookupEntries, conditionColumns, ruleConditions,
      targetColumns, targetMode, perTargetValues,
      exportMode, exportFilename, fileName: file?.name ?? '',
    };
    try {
      if (fileBuffer) {
        if (fileBuffer.byteLength > IDB_THRESHOLD) {
          cfg.wbStore = 'idb';
          await idbPut(STORAGE_KEY, fileBuffer);
          localStorage.setItem(STORAGE_KEY, JSON.stringify(cfg));
        } else {
          cfg.wbStore = 'ls';
          cfg.wbB64 = arrToB64(fileBuffer);
          localStorage.setItem(STORAGE_KEY, JSON.stringify(cfg));
        }
      } else {
        localStorage.setItem(STORAGE_KEY, JSON.stringify(cfg));
      }
    } catch { /* quota / IDB unavailable — ignore */ }
  }, [step, maxStep, selectedSheet, headerRowIndex, lookupEntries, conditionColumns, ruleConditions,
      targetColumns, targetMode, perTargetValues, exportMode, exportFilename, file, fileBuffer]);

  useEffect(() => {
    if (restoringRef.current) return;
    clearTimeout(saveTimer.current);
    saveTimer.current = setTimeout(doSave, 600);
    return () => clearTimeout(saveTimer.current);
  }, [doSave]);

  // restore on mount
  useEffect(() => {
    (async () => {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (!raw) return;
      try {
        restoringRef.current = true;
        const cfg = JSON.parse(raw);
        let buf = null;
        if (cfg.wbStore === 'idb') buf = await idbGet(STORAGE_KEY);
        else if (cfg.wbStore === 'ls' && cfg.wbB64) buf = b64ToArr(cfg.wbB64);

        if (buf) {
          const wb = XLSX.read(buf, { type: 'array', cellDates: true, cellNF: true, cellStyles: true });
          setFileBuffer(buf);
          setWorkbook(wb);
          setFile({ name: cfg.fileName ?? 'restored.xlsx' });
          if (cfg.selectedSheet && wb.Sheets[cfg.selectedSheet]) {
            setSelectedSheet(cfg.selectedSheet);
            const r = sheetToRows(wb.Sheets[cfg.selectedSheet]);
            setRows(r);
            setDetectedHeader(detectHeader(r));
          }
        }
        if (cfg.selectedSheet) setSelectedSheet(cfg.selectedSheet);
        if (cfg.headerRowIndex != null) {
          setHeaderRowIndex(cfg.headerRowIndex);
          setHeaderInput(String(cfg.headerRowIndex + 1));
        }
        if (cfg.lookupEntries) setLookupEntries(cfg.lookupEntries);
        if (cfg.conditionColumns) setConditionColumns(cfg.conditionColumns);
        if (cfg.ruleConditions) setRuleConditions(cfg.ruleConditions);
        if (cfg.targetColumns) setTargetColumns(cfg.targetColumns);
        if (cfg.targetMode) setTargetMode(cfg.targetMode);
        if (cfg.perTargetValues) setPerTargetValues(cfg.perTargetValues);
        if (cfg.exportMode) setExportMode(cfg.exportMode);
        if (cfg.exportFilename) setExportFilename(cfg.exportFilename);
        if (cfg.step != null) setStep(cfg.step);
        if (cfg.maxStep != null) setMaxStep(cfg.maxStep);
      } catch { /* ignore corrupt state */ }
      finally { restoringRef.current = false; }
    })();
  }, []);

  // ── clear ────────────────────────────────────────────────

  const handleClear = useCallback(async () => {
    localStorage.removeItem(STORAGE_KEY);
    try { await idbDel(STORAGE_KEY); } catch {}
    restoringRef.current = true;
    setStep(0); setMaxStep(0); setFile(null); setFileBuffer(null); setWorkbook(null);
    setSelectedSheet(''); setRows([]); setDetectedHeader(0);
    setHeaderRowIndex(0); setHeaderInput('1');
    setLookupEntries([]); setConditionColumns([]); setRuleConditions({});
    setTargetColumns([]); setTargetMode('same'); setPerTargetValues({});
    setExportMode('workbook'); setExportFilename('');
    setIsExporting(false); setExportDone(false); setExportedFile('');
    setError(''); setWarnMsg('');
    setTimeout(() => { restoringRef.current = false; }, 100);
  }, []);

  // ── step 0 handlers ──────────────────────────────────────

  const handleFileChange = useCallback(async (e) => {
    const f = e.target.files?.[0];
    if (!f) return;
    setError(''); setIsBusy(true); setBusyMsg('Reading file…');
    try {
      const buf = await f.arrayBuffer();
      const wb = XLSX.read(buf, { type: 'array', cellDates: true, cellNF: true, cellStyles: true });
      if (!wb.SheetNames.length) throw new Error('Workbook has no sheets.');
      const firstSheet = wb.SheetNames[0];
      const r = sheetToRows(wb.Sheets[firstSheet]);
      const det = detectHeader(r);
      setFile(f); setFileBuffer(buf); setWorkbook(wb);
      setSelectedSheet(firstSheet); setRows(r);
      setDetectedHeader(det); setHeaderRowIndex(det); setHeaderInput(String(det + 1));
    } catch (err) {
      setError('Could not read the file: ' + (err.message ?? err));
    } finally { setIsBusy(false); setBusyMsg(''); }
  }, []);

  const handleSheetChange = useCallback((sheetName) => {
    setSelectedSheet(sheetName);
    if (!workbook?.Sheets[sheetName]) return;
    const r = sheetToRows(workbook.Sheets[sheetName]);
    const det = detectHeader(r);
    setRows(r); setDetectedHeader(det);
    setHeaderRowIndex(det); setHeaderInput(String(det + 1));
    if (lookupEntries.length > 0 || conditionColumns.length > 0) {
      setLookupEntries([]); setConditionColumns([]); setRuleConditions({});
      setTargetColumns([]); setPerTargetValues({});
      setWarnMsg('Sheet changed — downstream configuration cleared.');
    }
  }, [workbook, lookupEntries.length, conditionColumns.length]);

  const handleApplyHeader = useCallback(() => {
    const idx = parseInt(headerInput, 10) - 1;
    if (isNaN(idx) || idx < 0 || idx >= rows.length) return;
    setHeaderRowIndex(idx);
    if (conditionColumns.length > 0 || targetColumns.length > 0) {
      setConditionColumns([]); setRuleConditions({});
      setTargetColumns([]); setPerTargetValues({});
      setWarnMsg('Header row changed — condition and target columns cleared.');
    }
  }, [headerInput, rows.length, conditionColumns.length, targetColumns.length]);

  // ── step 1 handlers ──────────────────────────────────────

  const addEntry = useCallback(() => {
    setLookupEntries(p => [...p, { id: uid(), label: '', value: '' }]);
  }, []);

  const updateEntry = useCallback((id, field, value) => {
    setLookupEntries(p => p.map(e => e.id === id ? { ...e, [field]: value } : e));
  }, []);

  const removeEntry = useCallback((id) => {
    setLookupEntries(p => p.filter(e => e.id !== id));
    setRuleConditions(p => { const n = { ...p }; delete n[id]; return n; });
    setPerTargetValues(p => {
      const n = {};
      for (const [t, vals] of Object.entries(p)) {
        const nv = { ...vals }; delete nv[id]; n[t] = nv;
      }
      return n;
    });
  }, []);

  // ── step 2 handlers ──────────────────────────────────────

  const toggleCondCol = useCallback((col) => {
    setConditionColumns(p => {
      if (p.includes(col)) {
        setRuleConditions(r => {
          const n = {};
          for (const [k, v] of Object.entries(r)) {
            const nv = { ...v }; delete nv[col]; n[k] = nv;
          }
          return n;
        });
        return p.filter(c => c !== col);
      }
      return [...p, col];
    });
  }, []);

  const setRuleCond = useCallback((entryId, col, val) => {
    setRuleConditions(p => ({
      ...p,
      [entryId]: { ...(p[entryId] ?? {}), [col]: val },
    }));
  }, []);

  // ── step 3 handlers ──────────────────────────────────────

  const toggleTargetCol = useCallback((col) => {
    setTargetColumns(p => p.includes(col) ? p.filter(c => c !== col) : [...p, col]);
  }, []);

  const setPerTargetVal = useCallback((targetCol, entryId, value) => {
    setPerTargetValues(p => ({
      ...p,
      [targetCol]: { ...(p[targetCol] ?? {}), [entryId]: value },
    }));
  }, []);

  // ── export ───────────────────────────────────────────────

  const handleExport = useCallback(async () => {
    if (!workbook || !selectedSheet) return;
    setIsExporting(true); setError('');
    const raw   = exportFilename.trim() || tsFilename(file?.name ?? 'output.xlsx');
    const fname = raw.endsWith('.xlsx') ? raw : raw + '.xlsx';

    try {
      const fillMap = Object.fromEntries(fillResults.map(r => [r.rowIndex, r.fills]));
      let outBuf;

      // ── preferred: direct ZIP/XML patch (preserves all original styles) ──
      if (fileBuffer && isPatchSupported() && exportMode === 'workbook') {
        outBuf = await patchXlsxCells(fileBuffer, selectedSheet, fillMap, headers);
      } else if (fileBuffer && isPatchSupported() && exportMode === 'sheet') {
        // Patch whole workbook then re-extract just the target sheet via XLSX
        // (simplest to just use the fallback for single-sheet exports)
        outBuf = await _exportFallback(workbook, selectedSheet, headers, fillMap, exportMode);
      } else {
        // ── fallback: xlsx-js-style round-trip (may lose some styles) ──
        outBuf = await _exportFallback(workbook, selectedSheet, headers, fillMap, exportMode);
      }

      const blob = new Blob([outBuf], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url; a.download = fname;
      document.body.appendChild(a); a.click(); document.body.removeChild(a);
      setTimeout(() => URL.revokeObjectURL(url), 60_000);
      setExportDone(true); setExportedFile(fname);
    } catch (err) {
      setError('Export failed: ' + (err.message ?? err));
    } finally { setIsExporting(false); }
  }, [workbook, selectedSheet, exportMode, exportFilename, file, fileBuffer, fillResults, headers]);

  // ── navigation ───────────────────────────────────────────

  const goTo = useCallback((s) => {
    setError(''); setWarnMsg('');
    setStep(s);
    setMaxStep(prev => Math.max(prev, s));
  }, []);

  // ── render ───────────────────────────────────────────────

  return (
    <div className="cff-tool">

      {/* Hero */}
      <div className="cff-hero">
        <h1 className="cff-hero-title">Conditional Formula Filler</h1>
        <p className="cff-hero-sub">
          Fill cells based on multi-column conditions — no formulas needed.
          Equivalent to nested IF/AND logic, built visually. Runs entirely in your browser.
        </p>
      </div>

      {/* Clear saved data */}
      <div className="cff-topbar">
        <button className="cff-clear-btn" onClick={handleClear}>
          ✕ Clear saved data
        </button>
      </div>

      {/* Stepper */}
      <nav className="cff-stepper" aria-label="Wizard steps">
        {STEP_LABELS.map((label, i) => (
          <button
            key={i}
            className={
              'cff-step'
              + (i === step ? ' is-active' : '')
              + (i < step ? ' is-done' : (i <= maxStep ? ' is-done' : ''))
              + (i > maxStep ? ' is-locked' : '')
            }
            onClick={() => i !== step && i <= maxStep && goTo(i)}
            disabled={i > maxStep}
            aria-current={i === step ? 'step' : undefined}
          >
            <span className="cff-step-dot">{i < step ? '✓' : (i < maxStep ? '✓' : i + 1)}</span>
            <span className="cff-step-lbl">{label}</span>
          </button>
        ))}
      </nav>

      {/* Global messages */}
      {error && (
        <div className="cff-error" role="alert">
          {error}
          <button className="cff-dismiss" onClick={() => setError('')}>✕</button>
        </div>
      )}
      {warnMsg && (
        <div className="cff-warn" role="status">
          {warnMsg}
          <button className="cff-dismiss" onClick={() => setWarnMsg('')}>✕</button>
        </div>
      )}
      {isBusy && (
        <div className="cff-progress" role="status">
          <span className="cff-spinner" aria-hidden="true" />
          {busyMsg}
        </div>
      )}

      <div className="cff-body">

        {/* ── Step 0: Upload ─────────────────────────────────── */}
        {step === 0 && (
          <div className="cff-card">
            <h2 className="cff-card-h">Upload workbook</h2>
            <p className="cff-hint">
              Accepts .xlsx, .xls, and .csv files. Nothing leaves your browser.
            </p>

            <label className="cff-dropzone">
              <input
                type="file"
                accept=".xlsx,.xls,.csv"
                onChange={handleFileChange}
                className="cff-file-input"
              />
              {file
                ? <span className="cff-file-chosen">📄 {file.name}</span>
                : <span className="cff-file-prompt">Click to choose a file</span>
              }
            </label>

            {workbook && (
              <>
                <div className="cff-field-group">
                  <label className="cff-field-label" htmlFor="cff-sheet">Sheet</label>
                  <select
                    id="cff-sheet"
                    className="cff-select"
                    value={selectedSheet}
                    onChange={e => handleSheetChange(e.target.value)}
                  >
                    {workbook.SheetNames.map(n => <option key={n} value={n}>{n}</option>)}
                  </select>
                </div>

                {rows.length > 0 && (
                  <>
                    <div className="cff-field-group">
                      <label className="cff-field-label" htmlFor="cff-hrow">
                        Header row
                        <span className="cff-field-hint"> (auto-detected: row {detectedHeader + 1})</span>
                      </label>
                      <div className="cff-row-ctrl">
                        <input
                          id="cff-hrow"
                          className="cff-input cff-input-narrow"
                          type="number"
                          min="1"
                          max={rows.length}
                          value={headerInput}
                          onChange={e => setHeaderInput(e.target.value)}
                        />
                        <button className="cff-ghost-btn" onClick={handleApplyHeader}>
                          Apply
                        </button>
                      </div>
                    </div>

                    {/* Header preview chips */}
                    {headerPreview.length > 0 && (
                      <div className="cff-chips">
                        {headerPreview.slice(0, 12).map((h, i) => (
                          <span key={i} className="cff-chip">{h}</span>
                        ))}
                        {headerPreview.length > 12 && (
                          <span className="cff-chip cff-chip-more">
                            +{headerPreview.length - 12} more
                          </span>
                        )}
                      </div>
                    )}

                    {/* Clickable raw preview */}
                    <div className="cff-raw-preview">
                      <p className="cff-section-label">Click a row to use it as the header</p>
                      <div className="cff-raw-scroll">
                        <table className="cff-raw-table">
                          <tbody>
                            {rows.slice(0, 10).map((row, i) => (
                              <tr
                                key={i}
                                className={
                                  'cff-raw-row'
                                  + (i === headerRowIndex ? ' is-header' : '')
                                }
                                onClick={() => {
                                  setHeaderInput(String(i + 1));
                                  setHeaderRowIndex(i);
                                  if (conditionColumns.length > 0 || targetColumns.length > 0) {
                                    setConditionColumns([]); setRuleConditions({});
                                    setTargetColumns([]); setPerTargetValues({});
                                    setWarnMsg('Header row changed — downstream config cleared.');
                                  }
                                }}
                              >
                                <td className="cff-raw-rownum">{i + 1}</td>
                                {row.slice(0, 8).map((cell, ci) => (
                                  <td key={ci} className="cff-raw-cell">
                                    {cell != null && String(cell).trim() ? String(cell) : ''}
                                  </td>
                                ))}
                                {row.length > 8 && <td className="cff-raw-more">…</td>}
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    </div>
                  </>
                )}

                <div className="cff-actions">
                  <button
                    className="cff-btn-primary"
                    onClick={() => goTo(1)}
                    disabled={!rows.length || !headers.length}
                  >
                    Continue →
                  </button>
                </div>
              </>
            )}
          </div>
        )}

        {/* ── Step 1: Lookup values ───────────────────────────── */}
        {step === 1 && (
          <div className="cff-card">
            <h2 className="cff-card-h">Define lookup values</h2>
            <p className="cff-hint">
              Each entry needs a <strong>label</strong> (e.g. <em>PSP-KHR</em>) and a
              <strong> fill value</strong> (e.g. <em>009899838</em>). Labels must be unique.
              In "per target" mode (Step 3) you can override the value per column.
            </p>

            <div className="cff-lookup-table">
              <div className="cff-lookup-header">
                <span>Label</span>
                <span>Fill value</span>
                <span />
              </div>
              {lookupEntries.map((entry, idx) => (
                <div key={entry.id} className="cff-lookup-row">
                  <input
                    className={'cff-input' + (!entry.label.trim() ? ' cff-input-invalid' : '')}
                    placeholder={`e.g. Case ${idx + 1}`}
                    value={entry.label}
                    onChange={e => updateEntry(entry.id, 'label', e.target.value)}
                    aria-label="Label"
                  />
                  <input
                    className={'cff-input' + (!entry.value.trim() ? ' cff-input-invalid' : '')}
                    placeholder="Fill value…"
                    value={entry.value}
                    onChange={e => updateEntry(entry.id, 'value', e.target.value)}
                    aria-label="Fill value"
                  />
                  <button
                    className="cff-remove-btn"
                    onClick={() => removeEntry(entry.id)}
                    title="Remove entry"
                    aria-label="Remove"
                  >✕</button>
                </div>
              ))}
              {lookupEntries.length === 0 && (
                <p className="cff-empty-hint">No entries yet — add one below.</p>
              )}
            </div>

            <button className="cff-add-btn" onClick={addEntry}>+ Add entry</button>

            {dupLabels.length > 0 && (
              <p className="cff-inline-warn">
                Duplicate labels: {[...new Set(dupLabels)].join(', ')}
              </p>
            )}

            <div className="cff-actions">
              <button className="cff-ghost-btn" onClick={() => goTo(0)}>← Back</button>
              <button className="cff-btn-primary" onClick={() => goTo(2)} disabled={!step1Valid}>
                Continue →
              </button>
            </div>
          </div>
        )}

        {/* ── Step 2: Conditions ─────────────────────────────── */}
        {step === 2 && (
          <div className="cff-card">
            <h2 className="cff-card-h">Define conditions</h2>
            <p className="cff-hint">
              Pick which columns to check, then set the required value per column for each
              lookup entry. All conditions in a row must match (AND logic).
            </p>

            {/* Condition column picker */}
            <div className="cff-section">
              <p className="cff-section-label">Condition columns</p>
              <div className="cff-col-picker">
                {headers.map(h => (
                  <label
                    key={h}
                    className={'cff-col-option' + (conditionColumns.includes(h) ? ' is-selected' : '')}
                  >
                    <input
                      type="checkbox"
                      checked={conditionColumns.includes(h)}
                      onChange={() => toggleCondCol(h)}
                    />
                    <span>{h}</span>
                  </label>
                ))}
              </div>
            </div>

            {/* Rules grid */}
            {conditionColumns.length > 0 && lookupEntries.length > 0 && (
              <div className="cff-section">
                <p className="cff-section-label">
                  Rules — one row per lookup entry, one dropdown per condition column
                </p>
                <div className="cff-table-scroll">
                  <table className="cff-rules-table">
                    <thead>
                      <tr>
                        <th>Lookup entry</th>
                        {conditionColumns.map(col => <th key={col}>{col}</th>)}
                      </tr>
                    </thead>
                    <tbody>
                      {lookupEntries.map(entry => (
                        <tr key={entry.id}>
                          <td className="cff-rules-label">
                            {entry.label || <em className="cff-unnamed">unnamed</em>}
                            {entry.value && (
                              <span className="cff-rules-val">→ {entry.value}</span>
                            )}
                          </td>
                          {conditionColumns.map(col => {
                            const val = (ruleConditions[entry.id] || {})[col] ?? '';
                            return (
                              <td key={col}>
                                <select
                                  className={'cff-rules-select' + (!val ? ' cff-select-empty' : '')}
                                  value={val}
                                  onChange={e => setRuleCond(entry.id, col, e.target.value)}
                                >
                                  <option value="">— pick —</option>
                                  {(colUniqueValues[col] || []).map(v => (
                                    <option key={v} value={v}>{v}</option>
                                  ))}
                                </select>
                              </td>
                            );
                          })}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            )}

            {conflictingRules.length > 0 && (
              <div className="cff-inline-warn">
                ⚠ Conflicting rules: {conflictingRules.join('; ')}
              </div>
            )}

            {conditionColumns.length === 0 && (
              <p className="cff-hint cff-hint-sm">Select at least one condition column above.</p>
            )}

            <div className="cff-actions">
              <button className="cff-ghost-btn" onClick={() => goTo(1)}>← Back</button>
              <button className="cff-btn-primary" onClick={() => goTo(3)} disabled={!step2Valid}>
                Continue →
              </button>
            </div>
          </div>
        )}

        {/* ── Step 3: Target columns ──────────────────────────── */}
        {step === 3 && (
          <div className="cff-card">
            <h2 className="cff-card-h">Choose target columns</h2>
            <p className="cff-hint">
              Select which columns to fill. "Same rules" fills all targets with the same
              matched value. "Per target" lets you specify a different value per column.
            </p>

            <div className="cff-section">
              <p className="cff-section-label">Target columns</p>
              <div className="cff-col-picker">
                {headers.map(h => (
                  <label
                    key={h}
                    className={'cff-col-option' + (targetColumns.includes(h) ? ' is-selected' : '')}
                  >
                    <input
                      type="checkbox"
                      checked={targetColumns.includes(h)}
                      onChange={() => toggleTargetCol(h)}
                    />
                    <span>{h}</span>
                  </label>
                ))}
              </div>
            </div>

            {targetColumns.length > 0 && (
              <div className="cff-section">
                <p className="cff-section-label">Fill mode</p>
                <div className="cff-mode-toggle">
                  <label className={'cff-mode-opt' + (targetMode === 'same' ? ' is-active' : '')}>
                    <input
                      type="radio"
                      name="targetMode"
                      value="same"
                      checked={targetMode === 'same'}
                      onChange={() => setTargetMode('same')}
                    />
                    <div className="cff-mode-text">
                      <strong>Same rules for all targets</strong>
                      <span>Every target column receives the same matched fill value</span>
                    </div>
                  </label>
                  <label className={'cff-mode-opt' + (targetMode === 'per-target' ? ' is-active' : '')}>
                    <input
                      type="radio"
                      name="targetMode"
                      value="per-target"
                      checked={targetMode === 'per-target'}
                      onChange={() => setTargetMode('per-target')}
                    />
                    <div className="cff-mode-text">
                      <strong>Different values per target column</strong>
                      <span>Configure a separate fill value for each target column</span>
                    </div>
                  </label>
                </div>
              </div>
            )}

            {/* Per-target value grid */}
            {targetMode === 'per-target' && targetColumns.length > 0 && lookupEntries.length > 0 && (
              <div className="cff-section">
                <p className="cff-section-label">Fill values per target column</p>
                <div className="cff-table-scroll">
                  <table className="cff-rules-table">
                    <thead>
                      <tr>
                        <th>Lookup entry</th>
                        {targetColumns.map(col => <th key={col}>{col}</th>)}
                      </tr>
                    </thead>
                    <tbody>
                      {lookupEntries.map(entry => (
                        <tr key={entry.id}>
                          <td className="cff-rules-label">
                            {entry.label || <em className="cff-unnamed">unnamed</em>}
                          </td>
                          {targetColumns.map(tCol => (
                            <td key={tCol}>
                              <input
                                className="cff-rules-input"
                                placeholder="Fill value…"
                                value={(perTargetValues[tCol] || {})[entry.id] ?? ''}
                                onChange={e => setPerTargetVal(tCol, entry.id, e.target.value)}
                              />
                            </td>
                          ))}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            )}

            <div className="cff-actions">
              <button className="cff-ghost-btn" onClick={() => goTo(2)}>← Back</button>
              <button className="cff-btn-primary" onClick={() => goTo(4)} disabled={!step3Valid}>
                Preview →
              </button>
            </div>
          </div>
        )}

        {/* ── Step 4: Preview ────────────────────────────────── */}
        {step === 4 && (
          <div className="cff-card cff-card-wide">
            <h2 className="cff-card-h">Preview</h2>

            {/* Stats */}
            <div className="cff-stats">
              <div className="cff-stat">
                <span className="cff-stat-val">{fillStats.total.toLocaleString()}</span>
                <span className="cff-stat-lbl">Total rows</span>
              </div>
              <div className="cff-stat cff-stat-success">
                <span className="cff-stat-val">{fillStats.filled.toLocaleString()}</span>
                <span className="cff-stat-lbl">Will be filled</span>
              </div>
              <div className="cff-stat cff-stat-muted">
                <span className="cff-stat-val">{fillStats.empty.toLocaleString()}</span>
                <span className="cff-stat-lbl">No matching rule</span>
              </div>
            </div>

            {previewRows.length > 0 ? (
              <div className="cff-preview-wrap">
                <p className="cff-hint cff-hint-sm">
                  Showing first {previewRows.length} data rows. Highlighted cells show
                  before → after. Target columns marked with ★.
                </p>
                <div className="cff-table-scroll">
                  <table className="cff-preview-table">
                    <thead>
                      <tr>
                        <th className="cff-row-num-h">#</th>
                        {headers.map(h => (
                          <th
                            key={h}
                            className={targetColumns.includes(h) ? 'cff-target-th' : ''}
                          >
                            {targetColumns.includes(h) ? `★ ${h}` : h}
                          </th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {previewRows.map(({ rowIndex, cells, fills }) => (
                        <tr key={rowIndex} className={Object.keys(fills).length > 0 ? 'cff-row-changed' : ''}>
                          <td className="cff-row-num">{rowIndex + 1}</td>
                          {headers.map((h, ci) => {
                            const before = cells[ci] != null ? String(cells[ci]) : '';
                            const after = fills[h];
                            const willChange = after != null && after !== '';
                            return (
                              <td
                                key={h}
                                className={willChange ? 'cff-cell-changed' : (targetColumns.includes(h) ? 'cff-cell-target' : '')}
                                title={willChange ? `"${before}" → "${after}"` : undefined}
                              >
                                {willChange ? (
                                  <span className="cff-cell-diff">
                                    {before && <span className="cff-before">{before}</span>}
                                    <span className="cff-arrow">→</span>
                                    <span className="cff-after">{after}</span>
                                  </span>
                                ) : before}
                              </td>
                            );
                          })}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            ) : (
              <p className="cff-empty-hint">No data rows found or no rules matched.</p>
            )}

            <div className="cff-actions">
              <button className="cff-ghost-btn" onClick={() => goTo(3)}>← Back</button>
              <button
                className="cff-btn-primary"
                onClick={() => {
                  if (!exportFilename) setExportFilename(tsFilename(file?.name ?? 'output.xlsx'));
                  setExportDone(false);
                  goTo(5);
                }}
              >
                Export →
              </button>
            </div>
          </div>
        )}

        {/* ── Step 5: Export ─────────────────────────────────── */}
        {step === 5 && (
          <div className="cff-card">
            <h2 className="cff-card-h">Export</h2>
            <p className="cff-hint">
              A new file is always generated — your original is never modified.
            </p>

            <div className="cff-export-summary">
              <div className="cff-export-row">
                <span>Rows to fill</span>
                <strong>{fillStats.filled.toLocaleString()} of {fillStats.total.toLocaleString()}</strong>
              </div>
              <div className="cff-export-row">
                <span>Target columns</span>
                <strong>{targetColumns.join(', ')}</strong>
              </div>
              <div className="cff-export-row">
                <span>Sheet</span>
                <strong>{selectedSheet}</strong>
              </div>
            </div>

            <div className="cff-field-group">
              <label className="cff-field-label" htmlFor="cff-export-name">Filename</label>
              <input
                id="cff-export-name"
                className="cff-input cff-mono"
                value={exportFilename}
                onChange={e => setExportFilename(e.target.value)}
                placeholder="output_filled_YYYYMMDD_HHmm.xlsx"
              />
            </div>

            <div className="cff-section">
              <p className="cff-section-label">Export scope</p>
              <div className="cff-mode-toggle">
                <label className={'cff-mode-opt' + (exportMode === 'workbook' ? ' is-active' : '')}>
                  <input type="radio" name="exportMode" value="workbook"
                    checked={exportMode === 'workbook'} onChange={() => setExportMode('workbook')} />
                  <div className="cff-mode-text">
                    <strong>Entire workbook</strong>
                    <span>All sheets preserved — only the selected sheet is updated</span>
                  </div>
                </label>
                <label className={'cff-mode-opt' + (exportMode === 'sheet' ? ' is-active' : '')}>
                  <input type="radio" name="exportMode" value="sheet"
                    checked={exportMode === 'sheet'} onChange={() => setExportMode('sheet')} />
                  <div className="cff-mode-text">
                    <strong>Selected sheet only</strong>
                    <span>Export only "{selectedSheet}" as a new single-sheet workbook</span>
                  </div>
                </label>
              </div>
            </div>

            {exportDone && (
              <div className="cff-export-success">
                ✓ Downloaded <strong>{exportedFile}</strong>
              </div>
            )}

            <div className="cff-actions">
              <button className="cff-ghost-btn" onClick={() => goTo(4)}>← Back</button>
              <button
                className="cff-btn-primary"
                onClick={handleExport}
                disabled={isExporting || fillStats.filled === 0}
              >
                {isExporting ? 'Exporting…' : '⬇ Download'}
              </button>
              <button className="cff-ghost-btn cff-btn-danger" onClick={handleClear}>
                Start over
              </button>
            </div>
          </div>
        )}

        {/* ── Sidebar ────────────────────────────────────────── */}
        {step > 0 && (
          <aside className="cff-aside">
            <p className="cff-aside-title">Configuration</p>

            {file && (
              <div className="cff-pill">
                <span className="cff-pill-label">File</span>
                <span className="cff-pill-value" title={file.name}>
                  {file.name.length > 26 ? file.name.slice(0, 24) + '…' : file.name}
                </span>
                <button className="cff-pill-edit" onClick={() => goTo(0)} title="Edit">✎</button>
              </div>
            )}
            {selectedSheet && (
              <div className="cff-pill">
                <span className="cff-pill-label">Sheet</span>
                <span className="cff-pill-value">{selectedSheet}</span>
                <button className="cff-pill-edit" onClick={() => goTo(0)} title="Edit">✎</button>
              </div>
            )}
            {step > 1 && lookupEntries.length > 0 && (
              <div className="cff-pill">
                <span className="cff-pill-label">Lookups</span>
                <span className="cff-pill-value">{lookupEntries.length} entries</span>
                <button className="cff-pill-edit" onClick={() => goTo(1)} title="Edit">✎</button>
              </div>
            )}
            {step > 2 && conditionColumns.length > 0 && (
              <div className="cff-pill">
                <span className="cff-pill-label">Conditions</span>
                <span className="cff-pill-value" title={conditionColumns.join(', ')}>
                  {conditionColumns.length} column{conditionColumns.length !== 1 ? 's' : ''}
                </span>
                <button className="cff-pill-edit" onClick={() => goTo(2)} title="Edit">✎</button>
              </div>
            )}
            {step > 3 && targetColumns.length > 0 && (
              <>
                <div className="cff-pill">
                  <span className="cff-pill-label">Targets</span>
                  <span className="cff-pill-value" title={targetColumns.join(', ')}>
                    {targetColumns.length} column{targetColumns.length !== 1 ? 's' : ''}
                  </span>
                  <button className="cff-pill-edit" onClick={() => goTo(3)} title="Edit">✎</button>
                </div>
                <div className="cff-pill">
                  <span className="cff-pill-label">Mode</span>
                  <span className="cff-pill-value">
                    {targetMode === 'same' ? 'Same rules' : 'Per target'}
                  </span>
                  <button className="cff-pill-edit" onClick={() => goTo(3)} title="Edit">✎</button>
                </div>
              </>
            )}
            {step > 4 && (
              <div className="cff-pill cff-pill-stat">
                <span className="cff-pill-label">Will fill</span>
                <span className="cff-pill-value cff-pill-success">
                  {fillStats.filled.toLocaleString()} rows
                </span>
              </div>
            )}
          </aside>
        )}
      </div>
    </div>
  );
}

import { useState, useCallback, useMemo } from 'react';
import XLSX from 'xlsx-js-style';
import './hyperlink-generator.css';

/* ── utilities ───────────────────────────────────────────── */

function detectHeaderRow(rows) {
  const limit = Math.min(20, rows.length - 1);
  for (let i = 0; i < limit; i++) {
    const row = rows[i] || [];
    const next = rows[i + 1] || [];
    const nonEmpty = row.filter(c => c != null && String(c).trim() !== '');
    const nextNonEmpty = next.filter(c => c != null && String(c).trim() !== '');
    const strCells = nonEmpty.filter(c => typeof c === 'string');
    if (
      nonEmpty.length >= 2 &&
      nextNonEmpty.length >= 1 &&
      strCells.length >= Math.ceil(nonEmpty.length / 2)
    ) {
      return i;
    }
  }
  return 0;
}

function buildFormula(basePath, srcVal, friendlyBase, dynamic) {
  const path = `${basePath}${srcVal}`.replace(/"/g, '""');
  const rawLabel = dynamic ? `${friendlyBase} ${srcVal}`.trim() : friendlyBase;
  const label = (rawLabel || path).replace(/"/g, '""');
  return `HYPERLINK("${path}","${label}")`;
}

const FORMULA_STYLE = {
  alignment: { vertical: 'center', horizontal: 'center' },
  border: {
    top:    { style: 'thin', color: { rgb: '000000' } },
    bottom: { style: 'thin', color: { rgb: '000000' } },
    left:   { style: 'thin', color: { rgb: '000000' } },
    right:  { style: 'thin', color: { rgb: '000000' } },
  },
};

const STEP_LABELS = ['Upload', 'Headers', 'Target', 'Source', 'Path', 'Label', 'Generate'];

/* ── sub-components ──────────────────────────────────────── */

function SummaryRow({ label, value, mono }) {
  return (
    <div className="hl-summary-row">
      <span className="hl-summary-label">{label}</span>
      <span className={'hl-summary-value' + (mono ? ' hl-mono' : '')}>{value}</span>
    </div>
  );
}

function SummaryPill({ label, value, onEdit, mono }) {
  const display = value.length > 22 ? value.slice(0, 20) + '…' : value;
  return (
    <div className="hl-pill">
      <span className="hl-pill-label">{label}</span>
      <span className={'hl-pill-value' + (mono ? ' hl-mono' : '')} title={value}>{display}</span>
      <button className="hl-pill-edit" onClick={onEdit} title="Edit this step">✎</button>
    </div>
  );
}

/* ── main component ──────────────────────────────────────── */

export default function HyperlinkGeneratorTool() {
  const [step, setStep] = useState(0);

  // Step 0
  const [file, setFile] = useState(null);
  const [workbook, setWorkbook] = useState(null);
  const [selectedSheet, setSelectedSheet] = useState('');

  // Step 1
  const [rows, setRows] = useState([]);
  const [detectedHeaderRow, setDetectedHeaderRow] = useState(0);
  const [manualHeaderInput, setManualHeaderInput] = useState('1');
  const [confirmedHeaderRow, setConfirmedHeaderRow] = useState(null);

  // Step 2
  const [targetColumn, setTargetColumn] = useState('');
  const [targetCustom, setTargetCustom] = useState('');

  // Step 3
  const [sourceColumn, setSourceColumn] = useState('');

  // Step 4
  const [basePath, setBasePath] = useState('');

  // Step 5
  const [friendlyName, setFriendlyName] = useState('Open folder');
  const [friendlyDynamic, setFriendlyDynamic] = useState(false);

  // Step 6
  const [isGenerating, setIsGenerating] = useState(false);

  /* ── derived ── */

  const effectiveTarget = targetColumn === '__custom__' ? targetCustom : targetColumn;

  const headers = useMemo(() => {
    if (!rows.length || confirmedHeaderRow == null) return [];
    return (rows[confirmedHeaderRow] || [])
      .map(c => (c != null ? String(c).trim() : ''))
      .filter(Boolean);
  }, [rows, confirmedHeaderRow]);

  const manualRowIdx = useMemo(() => {
    const n = parseInt(manualHeaderInput, 10);
    return isNaN(n) ? 0 : Math.max(0, n - 1);
  }, [manualHeaderInput]);

  const manualHeaderPreview = useMemo(() => {
    return (rows[manualRowIdx] || [])
      .map(c => (c != null ? String(c).trim() : ''))
      .filter(Boolean);
  }, [rows, manualRowIdx]);

  const previewFormulas = useMemo(() => {
    if (!rows.length || confirmedHeaderRow == null || !sourceColumn || !basePath) return [];
    const hRow = rows[confirmedHeaderRow] || [];
    const srcIdx = hRow.findIndex(h => String(h ?? '').trim() === sourceColumn);
    if (srcIdx === -1) return [];
    return rows
      .slice(confirmedHeaderRow + 1, confirmedHeaderRow + 4)
      .map(row => (row && row[srcIdx] != null ? String(row[srcIdx]).trim() : ''))
      .filter(v => v !== '')
      .map(v => `=${buildFormula(basePath, v, friendlyName, friendlyDynamic)}`);
  }, [rows, confirmedHeaderRow, sourceColumn, basePath, friendlyName, friendlyDynamic]);

  const canGenerate = !!(
    file && workbook && confirmedHeaderRow != null && effectiveTarget && sourceColumn && basePath
  );

  /* ── navigation ── */

  const goToStep = useCallback((s) => {
    if (s === 0) {
      setConfirmedHeaderRow(null);
      setTargetColumn(''); setTargetCustom('');
      setSourceColumn(''); setBasePath('');
    } else if (s === 1) {
      setTargetColumn(''); setTargetCustom('');
      setSourceColumn('');
    }
    setStep(s);
  }, []);

  /* ── handlers ── */

  const handleFileChange = useCallback(async (e) => {
    const f = e.target.files?.[0];
    if (!f) return;
    try {
      const buf = await f.arrayBuffer();
      const wb = XLSX.read(buf, { type: 'array' });
      setFile(f);
      setWorkbook(wb);
      setSelectedSheet(wb.SheetNames[0]);
      setRows([]); setDetectedHeaderRow(0); setConfirmedHeaderRow(null);
      setTargetColumn(''); setTargetCustom(''); setSourceColumn(''); setBasePath('');
      setStep(0);
    } catch (err) {
      console.error('Failed to read workbook:', err);
    }
  }, []);

  const handleSheetConfirm = useCallback(() => {
    if (!workbook || !selectedSheet) return;
    const sheet = workbook.Sheets[selectedSheet];
    if (!sheet['!ref']) return;

    const range = XLSX.utils.decode_range(sheet['!ref']);

    // Build rows indexed by 0-based Excel row number so rows[3] = Excel row 4.
    // sheet_to_json starts at !ref's first row and loses the row offset —
    // this approach preserves it so detection and display stay in sync.
    const rawRows = Array.from(
      { length: range.e.r + 1 },
      () => new Array(range.e.c + 1).fill(''),
    );
    for (const [addr, cell] of Object.entries(sheet)) {
      if (addr.startsWith('!')) continue;
      const pos = XLSX.utils.decode_cell(addr);
      if (cell && cell.v != null) {
        rawRows[pos.r][pos.c] = cell.w !== undefined ? cell.w : String(cell.v);
      }
    }

    const detected = detectHeaderRow(rawRows);
    setRows(rawRows);
    setDetectedHeaderRow(detected);
    setManualHeaderInput(String(detected + 1));
    setConfirmedHeaderRow(null);
    setTargetColumn(''); setTargetCustom(''); setSourceColumn('');
    setStep(1);
  }, [workbook, selectedSheet]);

  const handleHeaderConfirm = useCallback(() => {
    setConfirmedHeaderRow(manualRowIdx);
    setStep(2);
  }, [manualRowIdx]);

  const handleGenerate = useCallback(() => {
    if (!canGenerate) return;
    setIsGenerating(true);
    try {
      const newWb = XLSX.utils.book_new();
      const origSheet = workbook.Sheets[selectedSheet];

      // Shallow-copy all cell objects so we don't mutate the original
      const newSheet = {};
      for (const [k, v] of Object.entries(origSheet)) {
        if (k === '!ref') { newSheet[k] = v; continue; }
        if (k.startsWith('!')) {
          newSheet[k] = Array.isArray(v)
            ? v.map(x => (typeof x === 'object' && x ? { ...x } : x))
            : (typeof v === 'object' && v ? { ...v } : v);
          continue;
        }
        newSheet[k] = typeof v === 'object' && v !== null ? { ...v } : v;
      }

      const hRow = rows[confirmedHeaderRow] || [];
      let targetColIdx = hRow.findIndex(h => String(h ?? '').trim() === effectiveTarget);
      const range = XLSX.utils.decode_range(newSheet['!ref']);

      if (targetColIdx === -1) {
        // Append new column
        targetColIdx = range.e.c + 1;
        const hdrAddr = XLSX.utils.encode_cell({ r: confirmedHeaderRow, c: targetColIdx });
        newSheet[hdrAddr] = { t: 's', v: effectiveTarget };
      }

      const srcIdx = hRow.findIndex(h => String(h ?? '').trim() === sourceColumn);
      if (srcIdx === -1) { setIsGenerating(false); return; }

      let maxC = range.e.c;
      for (let r = confirmedHeaderRow + 1; r <= range.e.r; r++) {
        // Use the string-formatted rows array for clean source values
        const dataRow = rows[r] || [];
        const srcVal = dataRow[srcIdx] != null ? String(dataRow[srcIdx]).trim() : '';
        if (!srcVal) continue;

        const tgtAddr = XLSX.utils.encode_cell({ r, c: targetColIdx });
        newSheet[tgtAddr] = { f: buildFormula(basePath, srcVal, friendlyName, friendlyDynamic), s: FORMULA_STYLE };
        if (targetColIdx > maxC) maxC = targetColIdx;
      }

      if (maxC > range.e.c) {
        range.e.c = maxC;
        newSheet['!ref'] = XLSX.utils.encode_range(range);
      }

      XLSX.utils.book_append_sheet(newWb, newSheet, selectedSheet);

      const buf = XLSX.write(newWb, { bookType: 'xlsx', type: 'array', cellStyles: true });
      const blob = new Blob([buf], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `${file.name.replace(/\.(xlsx?|xls)$/i, '')}_hyperlinks.xlsx`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      setTimeout(() => URL.revokeObjectURL(url), 60_000);
    } catch (err) {
      console.error('Generate failed:', err);
    } finally {
      setIsGenerating(false);
    }
  }, [
    canGenerate, workbook, selectedSheet, rows,
    confirmedHeaderRow, effectiveTarget, sourceColumn,
    basePath, friendlyName, friendlyDynamic, file,
  ]);

  /* ── render ── */

  return (
    <div className="hl-tool">
      <div className="hl-hero">
        <h1 className="hl-hero-title">Hyperlink Generator</h1>
        <p className="hl-hero-sub">
          Bulk-generate Excel <code>=HYPERLINK(…)</code> formulas from a column of values.
          Everything runs in your browser — nothing is uploaded.
        </p>
      </div>

      {/* Stepper */}
      <nav className="hl-stepper" aria-label="Steps">
        {STEP_LABELS.map((label, i) => (
          <button
            key={i}
            className={
              'hl-step' +
              (i === step ? ' is-active' : '') +
              (i < step ? ' is-done' : '') +
              (i > step ? ' is-locked' : '')
            }
            onClick={() => i < step && goToStep(i)}
            disabled={i > step}
            aria-current={i === step ? 'step' : undefined}
          >
            <span className="hl-step-dot">{i < step ? '✓' : i + 1}</span>
            <span className="hl-step-lbl">{label}</span>
          </button>
        ))}
      </nav>

      <div className="hl-body">
        {/* ── Step 0: Upload ── */}
        {step === 0 && (
          <div className="hl-card">
            <h2 className="hl-card-h">Upload workbook</h2>
            <label className="hl-dropzone">
              <input
                type="file"
                accept=".xlsx,.xls"
                onChange={handleFileChange}
                className="hl-file-input"
              />
              {file
                ? <span className="hl-file-chosen">📄 {file.name}</span>
                : <span className="hl-file-prompt">Click to choose an .xlsx or .xls file</span>
              }
            </label>

            {workbook && (
              <>
                <label className="hl-field-label" htmlFor="hl-sheet-select">Sheet</label>
                <select
                  id="hl-sheet-select"
                  className="hl-select"
                  value={selectedSheet}
                  onChange={e => setSelectedSheet(e.target.value)}
                >
                  {workbook.SheetNames.map(n => <option key={n} value={n}>{n}</option>)}
                </select>
                <div className="hl-actions">
                  <button className="hl-btn-primary" onClick={handleSheetConfirm}>
                    Continue →
                  </button>
                </div>
              </>
            )}
          </div>
        )}

        {/* ── Step 1: Header row ── */}
        {step === 1 && (
          <div className="hl-card">
            <h2 className="hl-card-h">Confirm header row</h2>
            <p className="hl-hint">
              Auto-detected header at <strong>row {detectedHeaderRow + 1}</strong>. Adjust if
              the first non-empty rows are metadata, not column names.
            </p>

            <label className="hl-field-label" htmlFor="hl-header-row">Header row (1-based)</label>
            <input
              id="hl-header-row"
              className="hl-input hl-input-narrow"
              type="number"
              min="1"
              max={rows.length || 100}
              value={manualHeaderInput}
              onChange={e => setManualHeaderInput(e.target.value)}
            />

            <div className="hl-header-preview-box">
              <p className="hl-preview-title">Headers from row {manualHeaderInput}:</p>
              <div className="hl-chips">
                {manualHeaderPreview.length > 0
                  ? <>
                      {manualHeaderPreview.slice(0, 8).map((h, i) => (
                        <span key={i} className="hl-chip">{h}</span>
                      ))}
                      {manualHeaderPreview.length > 8 && (
                        <span className="hl-chip hl-chip-more">
                          +{manualHeaderPreview.length - 8} more
                        </span>
                      )}
                    </>
                  : <span className="hl-no-headers">No headers found at this row</span>
                }
              </div>
            </div>

            <div className="hl-actions">
              <button
                className="hl-btn-primary"
                onClick={handleHeaderConfirm}
                disabled={manualHeaderPreview.length === 0}
              >
                Confirm row {manualHeaderInput} →
              </button>
            </div>
          </div>
        )}

        {/* ── Step 2: Target column ── */}
        {step === 2 && (
          <div className="hl-card">
            <h2 className="hl-card-h">Target column</h2>
            <p className="hl-hint">
              Where should the hyperlink formulas be written? Pick an existing column or create
              a new one — it will be appended to the right.
            </p>

            <label className="hl-field-label" htmlFor="hl-target-select">Column</label>
            <select
              id="hl-target-select"
              className="hl-select"
              value={targetColumn}
              onChange={e => {
                setTargetColumn(e.target.value);
                if (e.target.value !== '__custom__') setTargetCustom('');
              }}
            >
              <option value="">— select a column —</option>
              {headers.map(h => <option key={h} value={h}>{h}</option>)}
              <option value="__custom__">＋ Create new column…</option>
            </select>

            {targetColumn === '__custom__' && (
              <>
                <label className="hl-field-label" htmlFor="hl-target-custom">New column name</label>
                <input
                  id="hl-target-custom"
                  className="hl-input"
                  placeholder="e.g. HyperLink"
                  value={targetCustom}
                  onChange={e => setTargetCustom(e.target.value)}
                />
              </>
            )}

            <div className="hl-actions">
              <button
                className="hl-btn-primary"
                disabled={!effectiveTarget}
                onClick={() => setStep(3)}
              >
                Continue →
              </button>
            </div>
          </div>
        )}

        {/* ── Step 3: Source column ── */}
        {step === 3 && (
          <div className="hl-card">
            <h2 className="hl-card-h">Source column</h2>
            <p className="hl-hint">
              Values from this column will be appended to the base path to form each hyperlink
              target (e.g. <code>TC-001</code>, <code>TC-002</code>).
            </p>

            <label className="hl-field-label" htmlFor="hl-source-select">Column</label>
            <select
              id="hl-source-select"
              className="hl-select"
              value={sourceColumn}
              onChange={e => setSourceColumn(e.target.value)}
            >
              <option value="">— select a column —</option>
              {headers.filter(h => h !== effectiveTarget).map(h => (
                <option key={h} value={h}>{h}</option>
              ))}
            </select>

            <div className="hl-actions">
              <button
                className="hl-btn-primary"
                disabled={!sourceColumn}
                onClick={() => setStep(4)}
              >
                Continue →
              </button>
            </div>
          </div>
        )}

        {/* ── Step 4: Base path ── */}
        {step === 4 && (
          <div className="hl-card">
            <h2 className="hl-card-h">Base path</h2>
            <p className="hl-hint">
              The source value is appended directly — include a trailing separator if needed.
            </p>
            <div className="hl-path-examples-hint">
              <span>Windows: <code>.\Cline\Rules\</code></span>
              <span>Mac / Linux: <code>./Cline/Rules/</code></span>
            </div>

            <label className="hl-field-label" htmlFor="hl-base-path">Base path</label>
            <input
              id="hl-base-path"
              className="hl-input hl-mono"
              placeholder={'e.g.  .\\Cline\\Rules\\  or  ./Cline/Rules/'}
              value={basePath}
              onChange={e => setBasePath(e.target.value)}
            />

            {basePath && !/[/\\]$/.test(basePath) && (
              <div className="hl-path-warn">
                <span>⚠ Path has no trailing separator — values will be joined directly (e.g. <code>{basePath}TC-001</code>).</span>
                <button
                  className="hl-fix-btn"
                  onClick={() => setBasePath(basePath + (basePath.includes('\\') ? '\\' : '/'))}
                >
                  Add {basePath.includes('\\') ? '\\' : '/'}
                </button>
              </div>
            )}

            {basePath && (
              <p className="hl-path-example">
                Example: <code className="hl-mono">{basePath}<em>TC-001</em></code>
              </p>
            )}

            <div className="hl-actions">
              <button className="hl-btn-primary" disabled={!basePath} onClick={() => setStep(5)}>
                Continue →
              </button>
            </div>
          </div>
        )}

        {/* ── Step 5: Label ── */}
        {step === 5 && (
          <div className="hl-card">
            <h2 className="hl-card-h">Friendly name</h2>
            <p className="hl-hint">The display text shown for each hyperlink cell in Excel.</p>

            <label className="hl-field-label" htmlFor="hl-label-input">Base label</label>
            <input
              id="hl-label-input"
              className="hl-input"
              placeholder="Open folder"
              value={friendlyName}
              onChange={e => setFriendlyName(e.target.value)}
            />
            <p className="hl-field-hint">Leave blank to use the full path as the label.</p>

            <label className="hl-toggle-label">
              <input
                type="checkbox"
                checked={friendlyDynamic}
                onChange={e => setFriendlyDynamic(e.target.checked)}
                className="hl-checkbox"
              />
              <span>Append source value (dynamic)</span>
            </label>

            <div className="hl-label-preview">
              <span className="hl-preview-title">Preview:</span>
              <code className="hl-mono">
                {friendlyDynamic
                  ? (`${friendlyName} TC-001`.trim() || 'TC-001')
                  : (friendlyName || `${basePath}TC-001`)}
              </code>
            </div>

            <div className="hl-actions">
              <button className="hl-btn-primary" onClick={() => setStep(6)}>Continue →</button>
            </div>
          </div>
        )}

        {/* ── Step 6: Generate ── */}
        {step === 6 && (
          <div className="hl-card">
            <h2 className="hl-card-h">Preview & generate</h2>

            <div className="hl-summary-table">
              <SummaryRow label="File" value={file?.name ?? ''} />
              <SummaryRow label="Sheet" value={selectedSheet} />
              <SummaryRow label="Header row" value={`Row ${(confirmedHeaderRow ?? 0) + 1}`} />
              <SummaryRow label="Target column" value={effectiveTarget} />
              <SummaryRow label="Source column" value={sourceColumn} />
              <SummaryRow label="Base path" value={basePath} mono />
              <SummaryRow
                label="Label"
                value={
                  friendlyDynamic
                    ? (`${friendlyName} <value>`.trim())
                    : (friendlyName || '(uses full path)')
                }
              />
            </div>

            {previewFormulas.length > 0 && (
              <div className="hl-formula-preview">
                <p className="hl-preview-title">
                  First {previewFormulas.length} formula{previewFormulas.length !== 1 ? 's' : ''}:
                </p>
                {previewFormulas.map((f, i) => (
                  <code key={i} className="hl-formula-line">{f}</code>
                ))}
              </div>
            )}

            {previewFormulas.length === 0 && confirmedHeaderRow != null && sourceColumn && basePath && (
              <p className="hl-warn">
                No data rows found with non-empty source values. Check your header row and source
                column.
              </p>
            )}

            <div className="hl-actions">
              <button
                className="hl-btn-primary hl-btn-generate"
                disabled={!canGenerate || isGenerating}
                onClick={handleGenerate}
              >
                {isGenerating ? 'Generating…' : '⬇ Download _hyperlinks.xlsx'}
              </button>
            </div>
          </div>
        )}

        {/* ── Sidebar summary ── */}
        {step > 0 && (
          <aside className="hl-aside">
            <p className="hl-aside-title">Configuration</p>
            {file && (
              <SummaryPill label="File" value={file.name} onEdit={() => goToStep(0)} />
            )}
            {selectedSheet && (
              <SummaryPill label="Sheet" value={selectedSheet} onEdit={() => goToStep(0)} />
            )}
            {confirmedHeaderRow != null && (
              <SummaryPill
                label="Header row"
                value={`Row ${confirmedHeaderRow + 1}`}
                onEdit={() => goToStep(1)}
              />
            )}
            {effectiveTarget && step > 2 && (
              <SummaryPill label="Target col" value={effectiveTarget} onEdit={() => goToStep(2)} />
            )}
            {sourceColumn && step > 3 && (
              <SummaryPill label="Source col" value={sourceColumn} onEdit={() => goToStep(3)} />
            )}
            {basePath && step > 4 && (
              <SummaryPill label="Base path" value={basePath} onEdit={() => goToStep(4)} mono />
            )}
          </aside>
        )}
      </div>
    </div>
  );
}

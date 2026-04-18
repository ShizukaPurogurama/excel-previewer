export default function StatsGrid({ fileName, sheetName, headerRowLabel, recordCount }) {
  return (
    <section className="stats-grid" aria-label="Workbook summary">
      <article className="stat-card">
        <span className="stat-label">Workbook</span>
        <strong id="stat-workbook">{fileName || 'None loaded'}</strong>
      </article>
      <article className="stat-card">
        <span className="stat-label">Selected sheet</span>
        <strong id="stat-sheet">{sheetName || '-'}</strong>
      </article>
      <article className="stat-card">
        <span className="stat-label">Header row</span>
        <strong id="stat-header-row">{headerRowLabel || '-'}</strong>
      </article>
      <article className="stat-card">
        <span className="stat-label">Parsed records</span>
        <strong id="stat-records">{recordCount}</strong>
      </article>
    </section>
  );
}

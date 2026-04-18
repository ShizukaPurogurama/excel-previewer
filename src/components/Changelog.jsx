import { APP_VERSION, CHANGELOG_ENTRIES } from '../constants.js';

function SummaryCard({ label, title, body, supportingText, isPrimary }) {
  return (
    <article className={'changelog-card' + (isPrimary ? ' is-primary' : '')}>
      <p className="changelog-card-label">{label}</p>
      <h3 className="changelog-card-title">{title}</h3>
      <p className="changelog-card-body">{body}</p>
      {supportingText ? <p className="changelog-card-supporting">{supportingText}</p> : null}
    </article>
  );
}

function ListSection({ title, items }) {
  return (
    <section className="changelog-entry-section">
      <h4 className="changelog-entry-section-title">{title}</h4>
      <ul className="changelog-points">
        {items.map((item, index) => (
          <li key={index}>{item}</li>
        ))}
      </ul>
    </section>
  );
}

function ChangelogEntry({ entry }) {
  return (
    <article className="changelog-entry">
      <header className="changelog-entry-header">
        <div className="changelog-entry-copy">
          <p className="changelog-entry-version">Version {entry.version}</p>
          <h3 className="changelog-entry-title">{entry.headline}</h3>
          <p className="changelog-entry-date">{entry.dateLabel}</p>
        </div>
        <span className="changelog-entry-stage">{entry.stage}</span>
      </header>
      <p className="changelog-entry-summary">{entry.summary}</p>
      <div className="changelog-entry-sections">
        <ListSection title="What changed" items={entry.highlights} />
        <ListSection title="Why it matters" items={entry.impact} />
        <ListSection title="Good to know" items={entry.notes} />
      </div>
    </article>
  );
}

export default function Changelog({ isVisible }) {
  if (!isVisible) return null;

  const latest = CHANGELOG_ENTRIES[0];

  return (
    <section id="changelog-section" className="panel data-panel">
      <div className="panel-heading">
        <div>
          <p className="section-label">Changelog</p>
          <h2>What changed in this version</h2>
        </div>
        <p id="changelog-meta" className="panel-meta">
          Plain-language release notes so anyone can understand what improved and why it matters.
        </p>
      </div>

      <div id="changelog-summary" className="changelog-summary">
        {latest ? (
          <>
            <SummaryCard
              label={latest.stage}
              title={'Version ' + APP_VERSION}
              body={latest.headline}
              supportingText={latest.summary}
              isPrimary
            />
            <SummaryCard
              label="Best for"
              title="Who this update helps"
              body={latest.audience}
            />
            <SummaryCard
              label="Release focus"
              title="Why this update exists"
              body={latest.focus}
            />
          </>
        ) : null}
      </div>

      <div id="changelog-list" className="changelog-list">
        {CHANGELOG_ENTRIES.map((entry) => (
          <ChangelogEntry key={entry.version} entry={entry} />
        ))}
      </div>
    </section>
  );
}

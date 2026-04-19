import { APP_VERSION, CHANGELOG_ENTRIES } from '../../../constants.js';

function Section({ title, items }) {
  return (
    <div className="ep-entry-section">
      <h4 className="ep-entry-sec-title">{title}</h4>
      <ul className="ep-entry-points">
        {items.map((item, i) => (
          <li key={i}>{item}</li>
        ))}
      </ul>
    </div>
  );
}

function Entry({ entry }) {
  return (
    <article className="ep-entry">
      <header className="ep-entry-hdr">
        <div>
          <p className="ep-entry-ver">Version {entry.version}</p>
          <h3 className="ep-entry-headline">{entry.headline}</h3>
          <p className="ep-entry-date">{entry.dateLabel}</p>
        </div>
        <span className="ep-entry-stage">{entry.stage}</span>
      </header>
      <p className="ep-entry-summary">{entry.summary}</p>
      <div className="ep-entry-sections">
        <Section title="What changed" items={entry.highlights} />
        <Section title="Why it matters" items={entry.impact} />
        <Section title="Good to know" items={entry.notes} />
      </div>
    </article>
  );
}

export default function ChangelogPanel({ isVisible }) {
  if (!isVisible) return null;

  return (
    <div className="ep-changelog">
      <h2 className="ep-cl-title">What&rsquo;s new in {APP_VERSION}</h2>
      <div className="ep-cl-entries">
        {CHANGELOG_ENTRIES.map((entry) => (
          <Entry key={entry.version} entry={entry} />
        ))}
      </div>
    </div>
  );
}

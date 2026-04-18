import { APP_VERSION } from '../../../constants.js';

export default function Topbar({ theme, onToggleTheme }) {
  const isDark = theme === 'dark';
  const label = isDark ? 'Light mode' : 'Dark mode';

  return (
    <header className="topbar">
      <div className="brand-block">
        <div className="brand-logo" aria-hidden="true">
          <svg
            width="22"
            height="22"
            viewBox="0 0 24 24"
            fill="none"
            xmlns="http://www.w3.org/2000/svg"
          >
            <rect
              x="2"
              y="3"
              width="20"
              height="18"
              rx="2.5"
              stroke="currentColor"
              strokeWidth="1.8"
              fill="none"
            />
            <path
              d="M8 8h8M8 12h8M8 16h5"
              stroke="currentColor"
              strokeWidth="1.8"
              strokeLinecap="round"
            />
            <path d="M2 7h20" stroke="currentColor" strokeWidth="1.8" />
          </svg>
        </div>
        <div className="brand-text">
          <div className="brand-meta">
            <p className="eyebrow">Excel Previewer</p>
            <span id="app-version-badge" className="version-badge">
              {APP_VERSION}
            </span>
          </div>
          <h1 className="brand-title">Workbook inspection in the browser</h1>
          <p className="brand-copy">
            Review spreadsheets locally, choose the right header row, and work with a cleaner
            parsed table without sending files to a server.
          </p>
        </div>
      </div>
      <button
        id="theme-toggle"
        className="theme-toggle"
        type="button"
        aria-pressed={isDark ? 'true' : 'false'}
        onClick={onToggleTheme}
      >
        <span className="theme-toggle-indicator" aria-hidden="true"></span>
        <span id="theme-toggle-label">{label}</span>
      </button>
    </header>
  );
}

import { Link, useLocation } from 'react-router-dom';
import { TOOLS } from '../tools/registry.js';

export default function Header({ onMenuOpen, onToggleTheme, theme }) {
  const location = useLocation();
  const isDark = theme === 'dark';

  const currentTool = TOOLS.find((t) => location.hash.replace('#', '') === t.path);

  return (
    <header className="site-header">
      <button
        className="site-header-burger"
        type="button"
        aria-label="Open menu"
        onClick={onMenuOpen}
      >
        <svg width="18" height="18" viewBox="0 0 18 18" fill="none">
          <path
            d="M2 4h14M2 9h14M2 14h14"
            stroke="currentColor"
            strokeWidth="1.6"
            strokeLinecap="round"
          />
        </svg>
      </button>

      <div className="site-header-breadcrumb">
        <Link to="/">Tools</Link>
        {currentTool && (
          <>
            <span className="site-header-breadcrumb-sep">/</span>
            <span className="site-header-current">{currentTool.name}</span>
          </>
        )}
      </div>

      <button
        className="theme-toggle"
        type="button"
        aria-pressed={isDark ? 'true' : 'false'}
        onClick={onToggleTheme}
      >
        <span className="theme-toggle-indicator" aria-hidden="true"></span>
        <span>{isDark ? 'Light mode' : 'Dark mode'}</span>
      </button>
    </header>
  );
}

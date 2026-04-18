import { NavLink, Link } from 'react-router-dom';
import { TOOLS } from '../tools/registry.js';

function Logo() {
  return (
    <div className="sidebar-logo">
      <div className="sidebar-logo-icon">
        <svg width="16" height="16" viewBox="0 0 24 24" fill="none">
          <rect x="2" y="3" width="20" height="18" rx="2.5" stroke="currentColor" strokeWidth="2" fill="none" />
          <path d="M8 8h8M8 12h8M8 16h5" stroke="currentColor" strokeWidth="2" strokeLinecap="round" />
          <path d="M2 7h20" stroke="currentColor" strokeWidth="2" />
        </svg>
      </div>
      <div>
        <div className="sidebar-logo-name">Tool Suite</div>
        <div className="sidebar-logo-sub">Browser tools</div>
      </div>
    </div>
  );
}

export default function Sidebar({ isOpen, onClose, isCollapsed, onCollapse }) {
  return (
    <>
      <div
        className={'sidebar-overlay' + (isOpen ? ' is-open' : '')}
        onClick={onClose}
        aria-hidden="true"
      />
      <nav
        className={'sidebar' + (isOpen ? ' is-open' : '') + (isCollapsed ? ' is-collapsed' : '')}
        aria-label="Tool navigation"
      >
        <div className="sidebar-header">
          <Link to="/" onClick={onClose}>
            <Logo />
          </Link>
          <button
            className="sidebar-collapse-btn"
            onClick={onCollapse}
            aria-label="Collapse sidebar"
            title="Collapse sidebar"
          >
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none">
              <path d="M15 18l-6-6 6-6" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
            </svg>
          </button>
        </div>

        <div className="sidebar-nav">
          <div className="sidebar-nav-label">Tools</div>

          {TOOLS.map((tool) => (
            <NavLink
              key={tool.id}
              to={tool.path}
              className={({ isActive }) =>
                'sidebar-nav-item' +
                (isActive ? ' is-active' : '') +
                (tool.comingSoon ? ' is-disabled' : '')
              }
              onClick={onClose}
              aria-disabled={tool.comingSoon ? 'true' : undefined}
              tabIndex={tool.comingSoon ? -1 : undefined}
            >
              <span className="sidebar-nav-icon">{tool.icon}</span>
              {tool.name}
            </NavLink>
          ))}
        </div>
      </nav>
    </>
  );
}

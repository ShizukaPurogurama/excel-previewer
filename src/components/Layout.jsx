import { useCallback, useEffect, useState } from 'react';
import { useLocation } from 'react-router-dom';
import Sidebar from './Sidebar.jsx';
import Header from './Header.jsx';

const SIDEBAR_KEY = 'toolsuite-sidebar-collapsed';

function readCollapsed() {
  try { return localStorage.getItem(SIDEBAR_KEY) === 'true'; } catch { return false; }
}
function saveCollapsed(val) {
  try { localStorage.setItem(SIDEBAR_KEY, String(val)); } catch {}
}

export default function Layout({ children, theme, onToggleTheme }) {
  const [sidebarOpen, setSidebarOpen] = useState(false);       // mobile drawer
  const [sidebarCollapsed, setSidebarCollapsed] = useState(readCollapsed); // desktop
  const location = useLocation();

  useEffect(() => { setSidebarOpen(false); }, [location]);

  const handleMenuToggle = useCallback(() => {
    if (window.innerWidth <= 768) {
      setSidebarOpen(true);
    } else {
      setSidebarCollapsed((prev) => {
        saveCollapsed(!prev);
        return !prev;
      });
    }
  }, []);

  const handleCollapse = useCallback(() => {
    saveCollapsed(true);
    setSidebarCollapsed(true);
  }, []);

  return (
    <div className={'app-frame' + (sidebarCollapsed ? ' sidebar-is-collapsed' : '')}>
      <Sidebar
        isOpen={sidebarOpen}
        onClose={() => setSidebarOpen(false)}
        isCollapsed={sidebarCollapsed}
        onCollapse={handleCollapse}
      />
      <div className="main-frame">
        <Header
          theme={theme}
          onToggleTheme={onToggleTheme}
          onMenuOpen={handleMenuToggle}
          sidebarCollapsed={sidebarCollapsed}
        />
        <main className="tool-content">{children}</main>
      </div>
    </div>
  );
}

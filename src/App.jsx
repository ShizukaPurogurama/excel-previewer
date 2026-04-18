import { Suspense } from 'react';
import { Routes, Route, Navigate } from 'react-router-dom';
import { TOOLS } from './tools/registry.js';
import { useTheme } from './hooks/useTheme.js';
import Layout from './components/Layout.jsx';
import HomePage from './pages/HomePage.jsx';

export default function App() {
  const { theme, toggleTheme } = useTheme();

  return (
    <Layout theme={theme} onToggleTheme={toggleTheme}>
      <Routes>
        <Route path="/" element={<HomePage />} />
        {TOOLS.map((tool) => (
          <Route
            key={tool.id}
            path={tool.path}
            element={
              <Suspense fallback={<div className="tool-loading">Loading…</div>}>
                <tool.component />
              </Suspense>
            }
          />
        ))}
        <Route path="*" element={<Navigate to="/" replace />} />
      </Routes>
    </Layout>
  );
}

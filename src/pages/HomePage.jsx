import { TOOLS } from '../tools/registry.js';
import ToolCard from '../components/ToolCard.jsx';

export default function HomePage() {
  return (
    <div className="home-page">
      <div className="home-page-hero">
        <p className="home-page-eyebrow">Tool Suite</p>
        <h1 className="home-page-title">Browser-based tools</h1>
        <p className="home-page-subtitle">
          Lightweight utilities that run entirely in your browser. No account, no uploads, no
          backend — your data stays on your device.
        </p>
      </div>

      <p className="home-section-title">Available tools</p>
      <div className="tools-grid">
        {TOOLS.map((tool) => (
          <ToolCard key={tool.id} tool={tool} />
        ))}
      </div>
    </div>
  );
}

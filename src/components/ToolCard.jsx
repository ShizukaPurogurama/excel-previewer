import { Link } from 'react-router-dom';

export default function ToolCard({ tool }) {
  if (tool.comingSoon) {
    return (
      <div className="tool-card is-disabled" aria-label={tool.name + ' — coming soon'}>
        <div className="tool-card-top">
          <span className="tool-card-icon">{tool.icon}</span>
          <span className="tool-card-badge">Coming soon</span>
        </div>
        <p className="tool-card-name">{tool.name}</p>
        <p className="tool-card-desc">{tool.description}</p>
      </div>
    );
  }

  return (
    <Link to={tool.path} className="tool-card">
      <div className="tool-card-top">
        <span className="tool-card-icon">{tool.icon}</span>
      </div>
      <p className="tool-card-name">{tool.name}</p>
      <p className="tool-card-desc">{tool.description}</p>
    </Link>
  );
}

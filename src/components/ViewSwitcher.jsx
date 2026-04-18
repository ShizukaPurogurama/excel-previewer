const VIEWS = [
  { id: 'preview', tabId: 'view-preview-tab', label: 'Preview rows' },
  { id: 'parsed', tabId: 'view-parsed-tab', label: 'Parsed table' },
  { id: 'changelog', tabId: 'view-changelog-tab', label: "What's new" },
];

export default function ViewSwitcher({ activeView, onChange }) {
  return (
    <section className="view-switcher" aria-label="Data view switcher">
      {VIEWS.map((view) => {
        const isActive = activeView === view.id;
        return (
          <button
            key={view.id}
            id={view.tabId}
            className={'view-tab' + (isActive ? ' is-active' : '')}
            type="button"
            data-view={view.id}
            aria-pressed={isActive ? 'true' : 'false'}
            onClick={() => onChange(view.id)}
          >
            {view.label}
          </button>
        );
      })}
    </section>
  );
}

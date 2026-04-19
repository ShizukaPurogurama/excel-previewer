export default function StatusStrip({ message, tone }) {
  const cls =
    'ep-status' +
    (tone === 'error' ? ' ep-err' : tone === 'success' ? ' ep-ok' : '');

  return (
    <div className={cls} role="status" aria-live="polite">
      <span className="ep-status-dot" aria-hidden="true" />
      {message}
    </div>
  );
}

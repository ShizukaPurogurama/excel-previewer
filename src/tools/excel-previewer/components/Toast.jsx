import { useEffect, useState } from 'react';

function ToastItem({ message, tone }) {
  const [visible, setVisible] = useState(false);

  useEffect(() => {
    const id = requestAnimationFrame(() => setVisible(true));
    return () => cancelAnimationFrame(id);
  }, []);

  const cls = [
    'toast',
    tone === 'error' && 'is-error',
    tone === 'success' && 'is-success',
    visible && 'is-visible',
  ]
    .filter(Boolean)
    .join(' ');

  return (
    <div className={cls} role="alert" aria-live="assertive">
      {message}
    </div>
  );
}

export default function Toast({ toasts }) {
  return (
    <div className="toast-container" aria-live="polite" aria-atomic="true">
      {toasts.map((t) => (
        <ToastItem key={t.id} message={t.message} tone={t.tone} />
      ))}
    </div>
  );
}

import { useEffect, useState } from 'react';

function Toast({ message, tone }) {
  const [visible, setVisible] = useState(false);

  useEffect(() => {
    const id = requestAnimationFrame(() => setVisible(true));
    return () => cancelAnimationFrame(id);
  }, []);

  const classes = ['toast'];
  if (tone === 'error') classes.push('is-error');
  else if (tone === 'success') classes.push('is-success');
  if (visible) classes.push('is-visible');

  return <div className={classes.join(' ')}>{message}</div>;
}

export default function ToastContainer({ toasts }) {
  return (
    <div
      id="toast-container"
      className="toast-container"
      aria-live="polite"
      aria-atomic="true"
    >
      {toasts.map((t) => (
        <Toast key={t.id} message={t.message} tone={t.tone} />
      ))}
    </div>
  );
}

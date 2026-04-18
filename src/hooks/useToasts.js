import { useCallback, useRef, useState } from 'react';

export function useToasts() {
  const [toasts, setToasts] = useState([]);
  const seedRef = useRef(0);

  const dismiss = useCallback((id) => {
    setToasts((prev) => prev.filter((t) => t.id !== id));
  }, []);

  const showToast = useCallback(
    (message, tone) => {
      if (!message) return;
      seedRef.current += 1;
      const id = seedRef.current;
      const timeoutMs = tone === 'error' ? 3200 : 2200;
      setToasts((prev) => [...prev, { id, message, tone: tone || '', createdAt: Date.now() }]);
      setTimeout(() => dismiss(id), timeoutMs);
    },
    [dismiss]
  );

  return { toasts, showToast, dismiss };
}

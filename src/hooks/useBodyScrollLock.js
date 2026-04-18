import { useEffect, useRef } from 'react';

export function useBodyScrollLock(isLocked) {
  const lockedScrollRef = useRef(0);

  useEffect(() => {
    if (!isLocked) return undefined;

    lockedScrollRef.current =
      window.scrollY ||
      window.pageYOffset ||
      document.documentElement.scrollTop ||
      0;

    document.body.style.position = 'fixed';
    document.body.style.top = '-' + lockedScrollRef.current + 'px';
    document.body.style.left = '0';
    document.body.style.right = '0';
    document.body.style.width = '100%';
    document.body.style.overflow = 'hidden';

    return () => {
      const top = document.body.style.top;

      document.body.style.position = '';
      document.body.style.top = '';
      document.body.style.left = '';
      document.body.style.right = '';
      document.body.style.width = '';
      document.body.style.overflow = '';

      const scrollTop = top
        ? Math.abs(Number.parseInt(top, 10)) || lockedScrollRef.current
        : lockedScrollRef.current;
      window.scrollTo(0, scrollTop);
      lockedScrollRef.current = 0;
    };
  }, [isLocked]);
}

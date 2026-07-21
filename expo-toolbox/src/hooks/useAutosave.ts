import { useEffect, useRef } from 'react';

type AutosaveOptions<T> = {
  value: T;
  enabled?: boolean;
  delayMs?: number;
  save: (value: T) => Promise<void> | void;
};

/** Debounced autosave for form-like offline edits. */
export function useAutosave<T>({ value, enabled = true, delayMs = 350, save }: AutosaveOptions<T>): void {
  const saveRef = useRef(save);
  saveRef.current = save;
  const first = useRef(true);

  useEffect(() => {
    if (!enabled) return;
    if (first.current) {
      first.current = false;
      return;
    }
    const timer = setTimeout(() => {
      void saveRef.current(value);
    }, delayMs);
    return () => clearTimeout(timer);
  }, [value, enabled, delayMs]);
}

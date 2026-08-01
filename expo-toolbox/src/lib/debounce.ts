/** Returns a debounced function with cancel and flush helpers. */
export function debounce<T extends (...args: never[]) => void>(
  fn: T,
  delayMs: number
): ((...args: Parameters<T>) => void) & { cancel: () => void; flush: () => void } {
  let timer: ReturnType<typeof setTimeout> | null = null;
  let lastArgs: Parameters<T> | null = null;

  const debounced = (...args: Parameters<T>) => {
    lastArgs = args;
    if (timer) clearTimeout(timer);
    timer = setTimeout(() => {
      timer = null;
      lastArgs = null;
      fn(...args);
    }, delayMs);
  };

  debounced.cancel = () => {
    if (timer) {
      clearTimeout(timer);
      timer = null;
    }
    lastArgs = null;
  };

  debounced.flush = () => {
    if (!timer || !lastArgs) return;
    clearTimeout(timer);
    timer = null;
    const args = lastArgs;
    lastArgs = null;
    fn(...args);
  };

  return debounced;
}

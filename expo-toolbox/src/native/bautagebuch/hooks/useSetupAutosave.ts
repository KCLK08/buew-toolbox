import { useCallback, useEffect, useRef } from 'react';

import { saveSetupModel } from '../db/database';

const AUTOSAVE_DELAY_MS = 420;

export function useSetupAutosave(templateId: string) {
  const timerRef = useRef<ReturnType<typeof setTimeout> | null>(null);
  const savingRef = useRef(false);
  const pendingRef = useRef<Record<string, unknown> | null>(null);
  const saveWaitersRef = useRef<Array<() => void>>([]);

  const notifySaveWaiters = useCallback(() => {
    const waiters = saveWaitersRef.current.splice(0);
    for (const resolve of waiters) {
      resolve();
    }
  }, []);

  const waitForSaveIdle = useCallback(async () => {
    if (!savingRef.current) return;
    await new Promise<void>((resolve) => {
      saveWaitersRef.current.push(resolve);
    });
  }, []);

  const flush = useCallback(async () => {
    if (!templateId) return;

    if (timerRef.current) {
      clearTimeout(timerRef.current);
      timerRef.current = null;
    }

    await waitForSaveIdle();

    while (pendingRef.current) {
      const snapshot = pendingRef.current;
      pendingRef.current = null;
      savingRef.current = true;
      try {
        const status =
          snapshot.status === 'ready'
            ? 'ready'
            : snapshot.status === 'archived'
              ? 'archived'
              : 'in_progress';
        await saveSetupModel(templateId, snapshot, status);
      } finally {
        savingRef.current = false;
        notifySaveWaiters();
      }
    }
  }, [templateId, notifySaveWaiters, waitForSaveIdle]);

  const schedule = useCallback(
    (model: Record<string, unknown>) => {
      pendingRef.current = model;
      if (timerRef.current) {
        clearTimeout(timerRef.current);
      }
      timerRef.current = setTimeout(() => {
        timerRef.current = null;
        void flush();
      }, AUTOSAVE_DELAY_MS);
    },
    [flush]
  );

  useEffect(() => {
    return () => {
      if (timerRef.current) {
        clearTimeout(timerRef.current);
        timerRef.current = null;
      }
      void flush();
    };
  }, [flush]);

  const isPending = useCallback(() => {
    return pendingRef.current !== null || timerRef.current !== null || savingRef.current;
  }, []);

  return { schedule, flush, isPending };
}

function cloneModel<T>(value: T): T {
  return JSON.parse(JSON.stringify(value)) as T;
}

export function mutateSetupModel(
  current: Record<string, unknown>,
  mutator: (model: Record<string, unknown>) => void
): Record<string, unknown> {
  const next = cloneModel(current);
  mutator(next);
  next.updatedAt = new Date().toISOString();
  return next;
}

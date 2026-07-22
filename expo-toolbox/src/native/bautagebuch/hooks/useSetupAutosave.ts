import { useCallback, useEffect, useRef } from 'react';

import { saveSetupModel } from '../db/database';

const AUTOSAVE_DELAY_MS = 420;

export function useSetupAutosave(templateId: string) {
  const timerRef = useRef<ReturnType<typeof setTimeout> | null>(null);
  const savingRef = useRef(false);
  const pendingRef = useRef<Record<string, unknown> | null>(null);

  const flush = useCallback(async () => {
    if (!templateId || !pendingRef.current) return;
    if (savingRef.current) return;
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
      if (pendingRef.current) {
        await flush();
      }
    }
  }, [templateId]);

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

  return { schedule, flush };
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

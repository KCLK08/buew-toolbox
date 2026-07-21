import { useCallback, useEffect, useRef } from 'react';

import { updateRun } from '../db/database';
import type { BautagebuchRun } from '../types';

const AUTOSAVE_DELAY_MS = 450;

type RunPatch = Partial<Pick<BautagebuchRun, 'title' | 'values' | 'sectionIndex' | 'status' | 'photoDoc' | 'completedAt'>>;

export function useRunAutosave(runId: string | undefined) {
  const timerRef = useRef<ReturnType<typeof setTimeout> | null>(null);
  const pendingRef = useRef<RunPatch | null>(null);
  const savingRef = useRef(false);

  const flush = useCallback(async () => {
    if (!runId || !pendingRef.current) return null;
    if (savingRef.current) return null;
    const patch = pendingRef.current;
    pendingRef.current = null;
    savingRef.current = true;
    try {
      return await updateRun(runId, patch);
    } finally {
      savingRef.current = false;
      if (pendingRef.current) {
        await flush();
      }
    }
  }, [runId]);

  const schedule = useCallback(
    (patch: RunPatch) => {
      if (!runId) return;
      pendingRef.current = { ...(pendingRef.current || {}), ...patch };
      if (timerRef.current) clearTimeout(timerRef.current);
      timerRef.current = setTimeout(() => {
        timerRef.current = null;
        void flush();
      }, AUTOSAVE_DELAY_MS);
    },
    [flush, runId]
  );

  useEffect(() => {
    return () => {
      if (timerRef.current) clearTimeout(timerRef.current);
      void flush();
    };
  }, [flush]);

  return { schedule, flush };
}

import { useCallback, useEffect, useState } from 'react';
import { AppState, type AppStateStatus } from 'react-native';

import {
  confirmPendingRestore,
  declinePendingRestore,
  runStartupIntegrityCheck
} from '../services/integrityService';
import { requestDatabaseBackup } from '../storage/backupService';
import type { IntegrityReport } from '../types/offline';

const STARTUP_CHECK_TIMEOUT_MS = 12_000;
const BACKUP_WARMUP_MS = 30_000;

function withTimeout<T>(promise: Promise<T>, timeoutMs: number): Promise<T> {
  return new Promise<T>((resolve, reject) => {
    const timer = setTimeout(() => {
      reject(new Error('Offline-Start Zeitüberschreitung'));
    }, timeoutMs);
    promise
      .then((value) => {
        clearTimeout(timer);
        resolve(value);
      })
      .catch((error) => {
        clearTimeout(timer);
        reject(error);
      });
  });
}

export function useOfflineBootstrap() {
  const [ready, setReady] = useState(false);
  const [report, setReport] = useState<IntegrityReport | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [restoreBusy, setRestoreBusy] = useState(false);

  useEffect(() => {
    let active = true;
    withTimeout(runStartupIntegrityCheck(), STARTUP_CHECK_TIMEOUT_MS)
      .then((result) => {
        if (!active) return;
        setReport(result);
        setReady(true);
        if (!result.ok && !result.pendingRestore) {
          setError(
            result.issues.find((issue) => issue.severity === 'error')?.message ??
              'Offline-Datenprüfung fehlgeschlagen.'
          );
        }
      })
      .catch((err: unknown) => {
        if (!active) return;
        setError(err instanceof Error ? err.message : 'Offline-Start fehlgeschlagen.');
        setReady(true);
      });
    return () => {
      active = false;
    };
  }, []);

  useEffect(() => {
    if (!ready) return;

    const startedAt = Date.now();
    const onChange = (state: AppStateStatus) => {
      if (state !== 'background') return;
      if (Date.now() - startedAt < BACKUP_WARMUP_MS) return;
      void requestDatabaseBackup('app_background');
    };
    const sub = AppState.addEventListener('change', onChange);
    return () => sub.remove();
  }, [ready]);

  const acceptRestore = useCallback(async () => {
    if (!report?.pendingRestore) return;
    setRestoreBusy(true);
    try {
      const next = await confirmPendingRestore(report.pendingRestore.backupUri);
      setReport(next);
      setError(next.ok ? null : next.issues.find((i) => i.severity === 'error')?.message ?? null);
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Wiederherstellung fehlgeschlagen.');
    } finally {
      setRestoreBusy(false);
    }
  }, [report]);

  const rejectRestore = useCallback(() => {
    declinePendingRestore();
    setReport((prev) => (prev ? { ...prev, pendingRestore: null } : prev));
  }, []);

  return {
    ready,
    report,
    error,
    restoreBusy,
    acceptRestore,
    rejectRestore
  };
}

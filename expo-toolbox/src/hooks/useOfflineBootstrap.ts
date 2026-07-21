import { useCallback, useEffect, useState } from 'react';
import { AppState, type AppStateStatus } from 'react-native';

import {
  confirmPendingRestore,
  declinePendingRestore,
  runStartupIntegrityCheck
} from '../services/integrityService';
import { requestDatabaseBackup } from '../storage/backupService';
import type { IntegrityReport } from '../types/offline';

export function useOfflineBootstrap() {
  const [ready, setReady] = useState(false);
  const [report, setReport] = useState<IntegrityReport | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [restoreBusy, setRestoreBusy] = useState(false);

  useEffect(() => {
    let active = true;
    runStartupIntegrityCheck()
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
    const onChange = (state: AppStateStatus) => {
      if (state === 'background' || state === 'inactive') {
        void requestDatabaseBackup('app_background');
      }
    };
    const sub = AppState.addEventListener('change', onChange);
    return () => sub.remove();
  }, []);

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

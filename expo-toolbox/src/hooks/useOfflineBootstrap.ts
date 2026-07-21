import { useEffect, useState } from 'react';

import { runStartupIntegrityCheck } from '../services/integrityService';
import type { IntegrityReport } from '../types/offline';

export function useOfflineBootstrap() {
  const [ready, setReady] = useState(false);
  const [report, setReport] = useState<IntegrityReport | null>(null);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    let active = true;
    runStartupIntegrityCheck()
      .then((result) => {
        if (!active) return;
        setReport(result);
        setReady(true);
        if (!result.ok) {
          setError(result.issues.find((issue) => issue.severity === 'error')?.message ?? 'Offline-Datenprüfung fehlgeschlagen.');
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

  return { ready, report, error };
}

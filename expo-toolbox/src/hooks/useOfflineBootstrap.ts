import { useEffect } from 'react';

import { runStartupIntegrityCheck } from '../services/integrityService';

/** Runs DB migrations and storage checks silently in the background. */
export function useOfflineBootstrap() {
  useEffect(() => {
    void runStartupIntegrityCheck().catch(() => undefined);
  }, []);
}

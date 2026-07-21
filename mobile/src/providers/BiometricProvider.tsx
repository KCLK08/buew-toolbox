import {
  createContext,
  useCallback,
  useContext,
  useEffect,
  useMemo,
  useState,
  type PropsWithChildren
} from 'react';
import { Alert } from 'react-native';
import { useAuth } from '@buew/shared';

import {
  authenticateWithBiometrics,
  getBiometricAvailability,
  isBiometricEnabled,
  setBiometricEnabled
} from '../services/biometric';

type BiometricContextValue = {
  ready: boolean;
  enabled: boolean;
  available: boolean;
  unlocked: boolean;
  needsUnlock: boolean;
  markUnlocked: () => void;
  enableBiometric: () => Promise<{ error: string | null }>;
  disableBiometric: () => Promise<{ error: string | null }>;
  unlockWithBiometric: () => Promise<boolean>;
  skipBiometricUnlock: () => void;
  promptEnableAfterLogin: () => void;
};

const BiometricContext = createContext<BiometricContextValue | null>(null);

export function BiometricProvider({ children }: PropsWithChildren) {
  const { isAuthenticated, signOut } = useAuth();
  const [ready, setReady] = useState(false);
  const [enabled, setEnabled] = useState(false);
  const [available, setAvailable] = useState(false);
  const [unlocked, setUnlocked] = useState(false);

  useEffect(() => {
    let active = true;
    Promise.all([getBiometricAvailability(), isBiometricEnabled()]).then(
      ([availability, enabledFlag]) => {
        if (!active) return;
        setAvailable(availability.available);
        setEnabled(enabledFlag);
        setReady(true);
      }
    );
    return () => {
      active = false;
    };
  }, []);

  useEffect(() => {
    if (!isAuthenticated) {
      setUnlocked(false);
    }
  }, [isAuthenticated]);

  const needsUnlock = isAuthenticated && enabled && available && !unlocked;

  const markUnlocked = useCallback(() => {
    setUnlocked(true);
  }, []);

  const unlockWithBiometric = useCallback(async () => {
    const ok = await authenticateWithBiometrics('BÜW-Toolbox entsperren');
    if (ok) setUnlocked(true);
    return ok;
  }, []);

  const skipBiometricUnlock = useCallback(() => {
    setUnlocked(false);
    void signOut();
  }, [signOut]);

  const enableBiometric = useCallback(async () => {
    const availability = await getBiometricAvailability();
    if (!availability.available) {
      return { error: 'Biometrie ist auf diesem Gerät nicht verfügbar.' };
    }
    const ok = await authenticateWithBiometrics('Biometrische Anmeldung aktivieren');
    if (!ok) {
      return { error: 'Biometrische Bestätigung abgebrochen.' };
    }
    await setBiometricEnabled(true);
    setEnabled(true);
    setAvailable(true);
    setUnlocked(true);
    return { error: null };
  }, []);

  const disableBiometric = useCallback(async () => {
    const ok = await authenticateWithBiometrics('Biometrische Anmeldung deaktivieren');
    if (!ok) {
      return { error: 'Biometrische Bestätigung abgebrochen.' };
    }
    await setBiometricEnabled(false);
    setEnabled(false);
    return { error: null };
  }, []);

  const promptEnableAfterLogin = useCallback(() => {
    void (async () => {
      const availability = await getBiometricAvailability();
      if (!availability.available) return;
      const already = await isBiometricEnabled();
      if (already) {
        setEnabled(true);
        setUnlocked(true);
        return;
      }
      Alert.alert(
        'Biometrische Anmeldung',
        'Möchtest du die biometrische Anmeldung aktivieren?',
        [
          {
            text: 'Später',
            style: 'cancel',
            onPress: () => setUnlocked(true)
          },
          {
            text: 'Aktivieren',
            onPress: () => {
              void enableBiometric().then((result) => {
                if (result.error) {
                  Alert.alert('Hinweis', result.error);
                  setUnlocked(true);
                }
              });
            }
          }
        ]
      );
    })();
  }, [enableBiometric]);

  const value = useMemo<BiometricContextValue>(
    () => ({
      ready,
      enabled,
      available,
      unlocked: !enabled || !available ? true : unlocked,
      needsUnlock,
      markUnlocked,
      enableBiometric,
      disableBiometric,
      unlockWithBiometric,
      skipBiometricUnlock,
      promptEnableAfterLogin
    }),
    [
      ready,
      enabled,
      available,
      unlocked,
      needsUnlock,
      markUnlocked,
      enableBiometric,
      disableBiometric,
      unlockWithBiometric,
      skipBiometricUnlock,
      promptEnableAfterLogin
    ]
  );

  return <BiometricContext.Provider value={value}>{children}</BiometricContext.Provider>;
}

export function useBiometric(): BiometricContextValue {
  const ctx = useContext(BiometricContext);
  if (!ctx) {
    throw new Error('useBiometric must be used within BiometricProvider');
  }
  return ctx;
}

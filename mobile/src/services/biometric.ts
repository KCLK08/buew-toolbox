import * as LocalAuthentication from 'expo-local-authentication';
import * as SecureStore from 'expo-secure-store';

const BIOMETRIC_ENABLED_KEY = 'buew.biometric.enabled';

export type BiometricAvailability = {
  hardware: boolean;
  enrolled: boolean;
  available: boolean;
};

export async function getBiometricAvailability(): Promise<BiometricAvailability> {
  const hardware = await LocalAuthentication.hasHardwareAsync();
  const enrolled = hardware ? await LocalAuthentication.isEnrolledAsync() : false;
  return {
    hardware,
    enrolled,
    available: hardware && enrolled
  };
}

export async function isBiometricEnabled(): Promise<boolean> {
  const value = await SecureStore.getItemAsync(BIOMETRIC_ENABLED_KEY);
  return value === '1';
}

export async function setBiometricEnabled(enabled: boolean): Promise<void> {
  if (enabled) {
    await SecureStore.setItemAsync(BIOMETRIC_ENABLED_KEY, '1');
    return;
  }
  await SecureStore.deleteItemAsync(BIOMETRIC_ENABLED_KEY);
}

export async function authenticateWithBiometrics(promptMessage: string): Promise<boolean> {
  const result = await LocalAuthentication.authenticateAsync({
    promptMessage,
    cancelLabel: 'Abbrechen',
    disableDeviceFallback: false,
    biometricsSecurityLevel: 'strong'
  });
  return result.success;
}

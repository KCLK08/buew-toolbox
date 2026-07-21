import { useEffect } from 'react';
import { useRouter } from 'expo-router';
import { AUTH_COPY, colors } from '@buew/shared';
import { StyleSheet, Text, View } from 'react-native';

import { AuthScreen } from '../src/components/AuthScreen';
import { PrimaryButton } from '../src/components/PrimaryButton';
import { useBiometric } from '../src/providers/BiometricProvider';

export default function BiometricUnlockScreen() {
  const router = useRouter();
  const biometric = useBiometric();

  useEffect(() => {
    if (!biometric.needsUnlock) {
      router.replace('/dashboard');
    }
  }, [biometric.needsUnlock, router]);

  const onUnlock = async () => {
    const ok = await biometric.unlockWithBiometric();
    if (ok) {
      router.replace('/dashboard');
    }
  };

  return (
    <AuthScreen title="Entsperren" subtitle="Melde dich biometrisch an, um fortzufahren.">
      <PrimaryButton label="Biometrisch entsperren" onPress={() => void onUnlock()} />
      <View style={styles.spacer} />
      <Text
        style={styles.link}
        onPress={() => {
          biometric.skipBiometricUnlock();
          router.replace('/login');
        }}
      >
        Stattdessen mit E-Mail und Passwort anmelden
      </Text>
      <Text style={styles.hint}>{AUTH_COPY.biometricPrompt}</Text>
    </AuthScreen>
  );
}

const styles = StyleSheet.create({
  spacer: {
    height: 4
  },
  link: {
    color: colors.accent,
    fontFamily: 'SpaceGrotesk_600SemiBold',
    fontSize: 14,
    textAlign: 'center'
  },
  hint: {
    color: colors.muted,
    fontFamily: 'SpaceGrotesk_400Regular',
    fontSize: 13,
    textAlign: 'center'
  }
});

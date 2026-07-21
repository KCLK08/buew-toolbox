import { Redirect } from 'expo-router';
import { useAuth, colors } from '@buew/shared';
import { ActivityIndicator, View } from 'react-native';

import { useBiometric } from '../src/providers/BiometricProvider';

export default function IndexGate() {
  const { ready, isAuthenticated } = useAuth();
  const biometric = useBiometric();

  if (!ready || !biometric.ready) {
    return (
      <View style={{ flex: 1, alignItems: 'center', justifyContent: 'center', backgroundColor: colors.bg }}>
        <ActivityIndicator color={colors.accent} size="large" />
      </View>
    );
  }

  if (!isAuthenticated) {
    return <Redirect href="/login" />;
  }

  if (biometric.needsUnlock) {
    return <Redirect href="/biometric-unlock" />;
  }

  return <Redirect href="/dashboard" />;
}

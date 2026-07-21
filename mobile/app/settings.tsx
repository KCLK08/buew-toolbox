import { useState } from 'react';
import { Redirect, useRouter } from 'expo-router';
import { useAuth, colors } from '@buew/shared';
import { Alert, Pressable, StyleSheet, Switch, Text, View } from 'react-native';
import { useSafeAreaInsets } from 'react-native-safe-area-context';

import { ToolboxBackground } from '../src/components/ToolboxBackground';
import { useBiometric } from '../src/providers/BiometricProvider';

export default function SettingsScreen() {
  const { isAuthenticated, user } = useAuth();
  const biometric = useBiometric();
  const router = useRouter();
  const insets = useSafeAreaInsets();
  const [busy, setBusy] = useState(false);

  if (!isAuthenticated) return <Redirect href="/login" />;
  if (biometric.needsUnlock) return <Redirect href="/biometric-unlock" />;

  const onToggle = async (next: boolean) => {
    setBusy(true);
    const result = next ? await biometric.enableBiometric() : await biometric.disableBiometric();
    setBusy(false);
    if (result.error) {
      Alert.alert('Biometrie', result.error);
    }
  };

  return (
    <ToolboxBackground>
      <View style={[styles.page, { paddingTop: insets.top + 20, paddingBottom: insets.bottom + 24 }]}>
        <Pressable onPress={() => router.back()} style={styles.back}>
          <Text style={styles.backText}>← Toolbox</Text>
        </Pressable>
        <View style={styles.card}>
          <Text style={styles.title}>Einstellungen</Text>
          <Text style={styles.meta}>{user?.displayName ?? user?.email}</Text>
          <Text style={styles.meta}>Rolle: {user?.role ?? 'benutzer'}</Text>

          <View style={styles.row}>
            <View style={{ flex: 1 }}>
              <Text style={styles.rowTitle}>Biometrische Anmeldung</Text>
              <Text style={styles.rowHint}>
                {biometric.available
                  ? 'Face ID / Fingerabdruck für App-Entsperrung'
                  : 'Auf diesem Gerät nicht verfügbar'}
              </Text>
            </View>
            <Switch
              value={biometric.enabled}
              disabled={!biometric.available || busy}
              onValueChange={(value) => void onToggle(value)}
              trackColor={{ true: colors.accent, false: colors.border }}
            />
          </View>
        </View>
      </View>
    </ToolboxBackground>
  );
}

const styles = StyleSheet.create({
  page: {
    flex: 1,
    paddingHorizontal: 20,
    gap: 16
  },
  back: {
    alignSelf: 'flex-start',
    paddingVertical: 8
  },
  backText: {
    color: colors.accent,
    fontFamily: 'SpaceGrotesk_600SemiBold'
  },
  card: {
    backgroundColor: colors.panel,
    borderColor: colors.border,
    borderWidth: 1,
    borderRadius: 18,
    padding: 20,
    gap: 10
  },
  title: {
    fontFamily: 'SpaceGrotesk_700Bold',
    fontSize: 28,
    color: colors.ink
  },
  meta: {
    fontFamily: 'SpaceGrotesk_400Regular',
    color: colors.muted
  },
  row: {
    marginTop: 12,
    flexDirection: 'row',
    alignItems: 'center',
    gap: 12
  },
  rowTitle: {
    fontFamily: 'SpaceGrotesk_600SemiBold',
    color: colors.ink,
    fontSize: 16
  },
  rowHint: {
    fontFamily: 'SpaceGrotesk_400Regular',
    color: colors.muted,
    fontSize: 13,
    marginTop: 4
  }
});

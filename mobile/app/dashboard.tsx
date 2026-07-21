import { useEffect } from 'react';
import { Redirect, useRouter } from 'expo-router';
import { AUTH_COPY, useAuth, colors } from '@buew/shared';
import { Pressable, ScrollView, StyleSheet, Text, View } from 'react-native';
import { useSafeAreaInsets } from 'react-native-safe-area-context';

import { ToolCard } from '../src/components/ToolCard';
import { ToolboxBackground } from '../src/components/ToolboxBackground';
import { TOOLBOX_TOOLS } from '../src/constants/tools';
import { useBiometric } from '../src/providers/BiometricProvider';

export default function DashboardScreen() {
  const { ready, isAuthenticated, user, signOut } = useAuth();
  const biometric = useBiometric();
  const router = useRouter();
  const insets = useSafeAreaInsets();

  useEffect(() => {
    if (ready && isAuthenticated && !biometric.needsUnlock) {
      // ensure unlock state after classic login
    }
  }, [ready, isAuthenticated, biometric.needsUnlock]);

  if (!ready) return null;
  if (!isAuthenticated) return <Redirect href="/login" />;
  if (biometric.needsUnlock) return <Redirect href="/biometric-unlock" />;

  return (
    <ToolboxBackground>
      <ScrollView
        contentContainerStyle={[
          styles.page,
          { paddingTop: insets.top + 24, paddingBottom: insets.bottom + 48 }
        ]}
      >
        <View style={styles.topbar}>
          <View style={{ flex: 1 }}>
            <Text style={styles.brand}>BÜW-Toolbox</Text>
            <Text style={styles.meta}>
              {user?.displayName ?? user?.email} · {user?.role ?? 'benutzer'}
            </Text>
          </View>
          <Pressable onPress={() => router.push('/settings')} style={styles.ghost}>
            <Text style={styles.ghostText}>Einstellungen</Text>
          </Pressable>
          <Pressable
            onPress={() => {
              void signOut().then(() => router.replace('/login'));
            }}
            style={styles.ghost}
          >
            <Text style={styles.ghostText}>{AUTH_COPY.logout}</Text>
          </Pressable>
        </View>

        <Text style={styles.title}>BÜW-Toolbox</Text>
        <Text style={styles.sub}>
          Zentrale Übersicht für digitale Baustellen‑Workflows und Dokumentation.
        </Text>

        <View style={styles.grid}>
          {TOOLBOX_TOOLS.map((tool) => (
            <ToolCard
              key={tool.id}
              tool={tool}
              onPress={() => router.push(tool.href)}
            />
          ))}
        </View>
      </ScrollView>
    </ToolboxBackground>
  );
}

const styles = StyleSheet.create({
  page: {
    paddingHorizontal: 20,
    gap: 12
  },
  topbar: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: 8,
    marginBottom: 16
  },
  brand: {
    fontFamily: 'SpaceGrotesk_700Bold',
    color: colors.ink,
    fontSize: 16
  },
  meta: {
    fontFamily: 'SpaceGrotesk_400Regular',
    color: colors.muted,
    fontSize: 13
  },
  ghost: {
    borderWidth: 1,
    borderColor: colors.border,
    backgroundColor: colors.white,
    borderRadius: 12,
    paddingHorizontal: 10,
    paddingVertical: 8
  },
  ghostText: {
    fontFamily: 'SpaceGrotesk_600SemiBold',
    color: colors.accent2,
    fontSize: 13
  },
  title: {
    fontFamily: 'SpaceGrotesk_700Bold',
    color: colors.ink,
    fontSize: 40,
    letterSpacing: -0.6
  },
  sub: {
    fontFamily: 'SpaceGrotesk_400Regular',
    color: colors.muted,
    fontSize: 16,
    lineHeight: 24,
    marginBottom: 12
  },
  grid: {
    gap: 18
  }
});

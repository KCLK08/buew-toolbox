import { useMemo, useState } from 'react';
import { ActivityIndicator, Alert, StyleSheet, Text, View } from 'react-native';
import { useRouter } from 'expo-router';

import { Card, PrimaryButton, Screen, TextField } from '../../../src/components/mobile';
import { colors, spacing, typography } from '../../../src/constants/theme';
import { buildBtbTitle } from '../../../src/native/bautagebuch/lib/btb-naming';
import { useBautagebuchWorkspace } from '../../../src/native/bautagebuch/hooks/useBautagebuchWorkspace';

function formatGreeting(): string {
  const hour = new Date().getHours();
  if (hour < 11) return 'Guten Morgen';
  if (hour < 18) return 'Guten Tag';
  return 'Guten Abend';
}

function formatTodayLabel(): string {
  return new Date().toLocaleDateString('de-DE', {
    weekday: 'long',
    day: 'numeric',
    month: 'long',
    year: 'numeric'
  });
}

export default function BautagebuchHomeTabScreen() {
  const router = useRouter();
  const { initialLoading, refreshing, error, templateName, templateReady, templateId, load, createNewRun } =
    useBautagebuchWorkspace();
  const [newName, setNewName] = useState('');
  const [creating, setCreating] = useState(false);

  const btbPreviewTitle = useMemo(() => (newName.trim() ? buildBtbTitle(newName) : ''), [newName]);

  const startRun = async () => {
    if (!templateId || !templateReady) {
      Alert.alert(
        'Vorlage nicht bereit',
        'Die aktive Vorlage ist noch nicht startbereit. Bitte Setup abschließen.'
      );
      return;
    }
    setCreating(true);
    try {
      const run = await createNewRun(newName);
      setNewName('');
      router.push(`/bautagebuch/run/${run.runId}`);
    } catch (err) {
      Alert.alert('Fehler', err instanceof Error ? err.message : 'BTB konnte nicht erstellt werden.');
    } finally {
      setCreating(false);
    }
  };

  return (
    <Screen
      title={formatGreeting()}
      subtitle={formatTodayLabel()}
      toolboxBack
      reserveTabBarSpace
      scroll
      refreshing={refreshing}
      onRefresh={() => void load('refresh')}
    >
      {initialLoading ? (
        <View style={styles.center}>
          <ActivityIndicator color={colors.accent} />
          <Text style={styles.muted}>Vorlage wird geladen…</Text>
        </View>
      ) : null}

      {error ? (
        <Card style={styles.errorCard}>
          <Text style={styles.error}>{error}</Text>
          <PrimaryButton label="Erneut versuchen" variant="secondary" onPress={() => void load('initial')} />
        </Card>
      ) : null}

      {!initialLoading && !error ? (
        <Card style={styles.startCard}>
          <Text style={styles.startCardTitle}>Neues Bautagebuch</Text>
          <Text style={styles.startCardHint}>
            Kurzen Namen für die Baustelle eingeben — Datum und Vorlage werden automatisch gesetzt.
          </Text>
          {templateName ? (
            <Text style={styles.activeTemplate}>
              Vorlage: {templateName}
              {!templateReady ? ' · Setup offen' : ''}
            </Text>
          ) : null}
          <TextField
            label="Baustelle / Strecke"
            hint="z. B. Strecke Nord, Tunnel Süd"
            value={newName}
            onChangeText={setNewName}
            autoCapitalize="sentences"
          />
          {btbPreviewTitle ? (
            <View style={styles.previewBox}>
              <Text style={styles.previewLabel}>Name des Bautagebuchs</Text>
              <Text style={styles.previewTitle}>{btbPreviewTitle}</Text>
            </View>
          ) : null}
          <PrimaryButton
            label={creating ? 'Wird erstellt…' : 'BTB starten'}
            disabled={creating || !templateId || !templateReady || !newName.trim()}
            onPress={() => void startRun()}
          />
        </Card>
      ) : null}
    </Screen>
  );
}

const styles = StyleSheet.create({
  center: { alignItems: 'center', gap: spacing.sm, paddingVertical: spacing.xl },
  muted: { ...typography.caption, color: colors.muted },
  errorCard: { gap: spacing.sm, borderColor: colors.danger },
  error: { ...typography.body, color: colors.danger },
  startCard: { gap: spacing.sm },
  startCardTitle: { ...typography.subtitle, color: colors.ink },
  startCardHint: { ...typography.caption, color: colors.muted },
  activeTemplate: { ...typography.caption, color: colors.accent2 },
  previewBox: {
    backgroundColor: colors.bg,
    borderRadius: spacing.inputRadius,
    borderWidth: 1,
    borderColor: colors.border,
    paddingHorizontal: spacing.sm,
    paddingVertical: spacing.sm,
    gap: spacing.xxs
  },
  previewLabel: { ...typography.label, color: colors.muted },
  previewTitle: { ...typography.bodyStrong, color: colors.accent }
});

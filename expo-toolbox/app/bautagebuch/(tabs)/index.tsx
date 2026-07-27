import { useMemo, useState } from 'react';
import { ActivityIndicator, Alert, Modal, StyleSheet, Text, View } from 'react-native';
import { useRouter } from 'expo-router';

import { PrimaryButton, Screen, TextField } from '../../../src/components/mobile';
import { colors, spacing, typography } from '../../../src/constants/theme';
import { BTBOverviewCard } from '../../../src/native/bautagebuch/components/home/BTBOverviewCard';
import { HomeHeader } from '../../../src/native/bautagebuch/components/home/HomeHeader';
import { NewBTBCard } from '../../../src/native/bautagebuch/components/home/NewBTBCard';
import { RecentActivityCard } from '../../../src/native/bautagebuch/components/home/RecentActivityCard';
import { useBautagebuchWorkspace } from '../../../src/native/bautagebuch/hooks/useBautagebuchWorkspace';
import { buildBtbTitle } from '../../../src/native/bautagebuch/lib/btb-naming';
import { computeBtbHomeStats, getRecentRuns } from '../../../src/native/bautagebuch/lib/home-utils';

export default function BautagebuchHomeTabScreen() {
  const router = useRouter();
  const { initialLoading, refreshing, error, templateReady, templateId, runs, load, createNewRun } =
    useBautagebuchWorkspace();
  const [creating, setCreating] = useState(false);
  const [startModalOpen, setStartModalOpen] = useState(false);
  const [newName, setNewName] = useState('');

  const stats = useMemo(() => computeBtbHomeStats(runs), [runs]);
  const recentRuns = useMemo(() => getRecentRuns(runs, 5), [runs]);
  const btbPreviewTitle = useMemo(() => (newName.trim() ? buildBtbTitle(newName) : ''), [newName]);

  const openStartModal = () => {
    if (!templateId || !templateReady) {
      Alert.alert(
        'Vorlage nicht bereit',
        'Die aktive Vorlage ist noch nicht startbereit. Bitte im Setup-Tab abschließen.'
      );
      return;
    }
    setStartModalOpen(true);
  };

  const closeStartModal = () => {
    setStartModalOpen(false);
    setNewName('');
  };

  const startRun = async () => {
    if (!newName.trim()) {
      Alert.alert('Baustelle fehlt', 'Bitte einen Namen für die Baustelle oder Strecke eingeben.');
      return;
    }
    setCreating(true);
    try {
      const run = await createNewRun(newName);
      closeStartModal();
      router.push(`/bautagebuch/run/${run.runId}`);
    } catch (err) {
      Alert.alert('Fehler', err instanceof Error ? err.message : 'BTB konnte nicht erstellt werden.');
    } finally {
      setCreating(false);
    }
  };

  return (
    <Screen
      title="Home"
      toolboxBack
      reserveTabBarSpace
      scroll
      refreshing={refreshing}
      onRefresh={() => void load('refresh')}
      contentStyle={styles.screenContent}
    >
      {initialLoading ? (
        <View style={styles.center}>
          <ActivityIndicator color={colors.accent} />
          <Text style={styles.muted}>Dashboard wird geladen…</Text>
        </View>
      ) : null}

      {error ? (
        <View style={styles.errorBox}>
          <Text style={styles.error}>{error}</Text>
          <PrimaryButton label="Erneut versuchen" variant="secondary" onPress={() => void load('initial')} />
        </View>
      ) : null}

      {!initialLoading && !error ? (
        <View style={styles.dashboard}>
          <HomeHeader />
          <NewBTBCard
            creating={creating}
            disabled={!templateId || !templateReady}
            onStart={openStartModal}
          />
          <BTBOverviewCard stats={stats} onPress={() => router.push('/bautagebuch/btbs')} />
          <RecentActivityCard
            runs={recentRuns}
            onOpenRun={(runId) => router.push(`/bautagebuch/run/${runId}`)}
          />
        </View>
      ) : null}

      <Modal visible={startModalOpen} transparent animationType="fade" onRequestClose={closeStartModal}>
        <View style={styles.modalBackdrop}>
          <View style={styles.modalCard}>
            <Text style={styles.modalTitle}>Neues Bautagebuch</Text>
            <Text style={styles.modalHint}>Baustelle oder Strecke benennen — Datum wird automatisch gesetzt.</Text>
            <TextField
              label="Baustelle / Strecke"
              hint="z. B. Bahnhof Frankfurt-Griesheim"
              value={newName}
              onChangeText={setNewName}
              autoCapitalize="sentences"
            />
            {btbPreviewTitle ? (
              <View style={styles.previewBox}>
                <Text style={styles.previewLabel}>Vorschau</Text>
                <Text style={styles.previewTitle}>{btbPreviewTitle}</Text>
              </View>
            ) : null}
            <View style={styles.modalActions}>
              <PrimaryButton label="Abbrechen" variant="ghost" onPress={closeStartModal} />
              <PrimaryButton
                label={creating ? 'Wird erstellt…' : '+ Starten'}
                disabled={creating || !newName.trim()}
                onPress={() => void startRun()}
              />
            </View>
          </View>
        </View>
      </Modal>
    </Screen>
  );
}

const styles = StyleSheet.create({
  screenContent: {
    gap: spacing.md
  },
  dashboard: {
    gap: spacing.md
  },
  center: {
    alignItems: 'center',
    gap: spacing.sm,
    paddingVertical: spacing.xl
  },
  muted: {
    ...typography.caption,
    color: colors.muted
  },
  errorBox: {
    gap: spacing.sm,
    padding: spacing.sm,
    borderRadius: 12,
    borderWidth: 1,
    borderColor: colors.danger,
    backgroundColor: colors.panel
  },
  error: {
    ...typography.body,
    color: colors.danger
  },
  modalBackdrop: {
    flex: 1,
    backgroundColor: colors.overlay,
    justifyContent: 'center',
    padding: spacing.xl
  },
  modalCard: {
    backgroundColor: colors.panelElevated,
    borderRadius: spacing.cardRadius,
    padding: spacing.md,
    gap: spacing.sm,
    borderWidth: 1,
    borderColor: colors.border
  },
  modalTitle: {
    ...typography.subtitle,
    color: colors.ink
  },
  modalHint: {
    ...typography.caption,
    color: colors.muted
  },
  previewBox: {
    backgroundColor: colors.bg,
    borderRadius: spacing.inputRadius,
    borderWidth: 1,
    borderColor: colors.border,
    paddingHorizontal: spacing.sm,
    paddingVertical: spacing.sm,
    gap: spacing.xxs
  },
  previewLabel: {
    ...typography.label,
    color: colors.muted
  },
  previewTitle: {
    ...typography.bodyStrong,
    color: colors.accent
  },
  modalActions: {
    flexDirection: 'row',
    justifyContent: 'flex-end',
    gap: spacing.xs,
    marginTop: spacing.xs
  }
});

import { useEffect, useRef, useState } from 'react';
import { ActivityIndicator, Alert, Modal, StyleSheet, Text, View } from 'react-native';
import { useRouter } from 'expo-router';

import {
  EmptyState,
  PrimaryButton,
  Screen,
  Section,
  TextField
} from '../../../src/components/mobile';
import { colors, spacing, typography } from '../../../src/constants/theme';
import { useToast } from '../../../src/contexts/ToastContext';
import { BautagebuchRunList } from '../../../src/native/bautagebuch/components/BautagebuchRunList';
import { useBautagebuchWorkspace } from '../../../src/native/bautagebuch/hooks/useBautagebuchWorkspace';
import type { BautagebuchRun } from '../../../src/native/bautagebuch/types';

export default function BautagebuchBtbsScreen() {
  const router = useRouter();
  const { showToast } = useToast();
  const {
    initialLoading,
    refreshing,
    error,
    runs,
    runTree,
    load,
    deleteRunById,
    renameRunById
  } = useBautagebuchWorkspace();

  const [selectionMode, setSelectionMode] = useState(false);
  const [selectedRunIds, setSelectedRunIds] = useState<string[]>([]);
  const [expandedYears, setExpandedYears] = useState<Set<number>>(new Set());
  const [expandedWeeks, setExpandedWeeks] = useState<Set<string>>(new Set());
  const [expandedProjects, setExpandedProjects] = useState<Set<string>>(new Set());
  const [renameRunId, setRenameRunId] = useState<string | null>(null);
  const [renameTitle, setRenameTitle] = useState('');
  const expansionInitialized = useRef(false);

  useEffect(() => {
    if (runs.length === 0) {
      expansionInitialized.current = false;
      setExpandedYears(new Set());
      setExpandedWeeks(new Set());
      setExpandedProjects(new Set());
      return;
    }
    if (expansionInitialized.current) return;
    setExpandedYears(new Set());
    setExpandedWeeks(new Set());
    setExpandedProjects(new Set());
    expansionInitialized.current = true;
  }, [runTree, runs.length]);

  const toggleYear = (year: number) => {
    setExpandedYears((current) => {
      const next = new Set(current);
      if (next.has(year)) next.delete(year);
      else next.add(year);
      return next;
    });
  };

  const toggleWeek = (weekKey: string) => {
    setExpandedWeeks((current) => {
      const next = new Set(current);
      if (next.has(weekKey)) next.delete(weekKey);
      else next.add(weekKey);
      return next;
    });
  };

  const toggleProject = (groupKey: string) => {
    setExpandedProjects((current) => {
      const next = new Set(current);
      if (next.has(groupKey)) next.delete(groupKey);
      else next.add(groupKey);
      return next;
    });
  };

  const toggleSelection = (runId: string) => {
    setSelectedRunIds((current) =>
      current.includes(runId) ? current.filter((id) => id !== runId) : [...current, runId]
    );
  };

  const deleteRun = (runId: string) => {
    Alert.alert('BTB löschen', 'Dieses Bautagebuch wirklich löschen?', [
      { text: 'Abbrechen', style: 'cancel' },
      {
        text: 'Löschen',
        style: 'destructive',
        onPress: () => {
          void deleteRunById(runId).then(() => {
            setSelectedRunIds((current) => current.filter((id) => id !== runId));
            showToast('BTB gelöscht');
          });
        }
      }
    ]);
  };

  const deleteSelected = () => {
    if (selectedRunIds.length === 0) return;
    Alert.alert(
      'Auswahl löschen',
      `${selectedRunIds.length} Bautagebuch${selectedRunIds.length === 1 ? '' : 'er'} wirklich löschen?`,
      [
        { text: 'Abbrechen', style: 'cancel' },
        {
          text: 'Löschen',
          style: 'destructive',
          onPress: () => {
            void Promise.all(selectedRunIds.map((runId) => deleteRunById(runId))).then(() => {
              setSelectedRunIds([]);
              setSelectionMode(false);
              showToast('Auswahl gelöscht');
            });
          }
        }
      ]
    );
  };

  const openRename = (run: BautagebuchRun) => {
    setRenameRunId(run.runId);
    setRenameTitle(run.title);
  };

  const submitRename = async () => {
    if (!renameRunId) return;
    const updated = await renameRunById(renameRunId, renameTitle);
    if (!updated) {
      Alert.alert('Umbenennen', 'Bitte einen gültigen Titel eingeben.');
      return;
    }
    setRenameRunId(null);
    setRenameTitle('');
    showToast('BTB umbenannt');
  };

  return (
    <Screen
      title="Bautagebücher"
      subtitle={`${runs.length} BTB${runs.length === 1 ? '' : 's'}`}
      toolboxBack
      reserveTabBarSpace
      scroll
      refreshing={refreshing}
      onRefresh={() => void load('refresh')}
    >
      {initialLoading ? (
        <View style={styles.center}>
          <ActivityIndicator color={colors.accent} />
          <Text style={styles.muted}>Bautagebücher werden geladen…</Text>
        </View>
      ) : null}

      {error ? (
        <View style={styles.errorBox}>
          <Text style={styles.error}>{error}</Text>
          <PrimaryButton label="Erneut versuchen" variant="secondary" onPress={() => void load('initial')} />
        </View>
      ) : null}

      {!initialLoading && !error ? (
        <Section
          title="Deine Bautagebücher"
          action={
            selectionMode ? (
              <View style={styles.listActions}>
                <PrimaryButton
                  label="Abbrechen"
                  variant="ghost"
                  compact
                  onPress={() => {
                    setSelectionMode(false);
                    setSelectedRunIds([]);
                  }}
                />
                <PrimaryButton
                  label={`Löschen (${selectedRunIds.length})`}
                  variant="secondary"
                  compact
                  disabled={selectedRunIds.length === 0}
                  onPress={deleteSelected}
                />
              </View>
            ) : (
              <PrimaryButton
                label="Auswahl"
                variant="ghost"
                compact
                disabled={runs.length === 0}
                onPress={() => setSelectionMode(true)}
              />
            )
          }
        >
          {runs.length === 0 ? (
            <EmptyState
              title="Noch keine BTB-Läufe"
              description="Wechsle zum Home-Tab, um dein erstes Bautagebuch zu starten."
            />
          ) : (
            <BautagebuchRunList
              tree={runTree}
              expandedYears={expandedYears}
              expandedWeeks={expandedWeeks}
              expandedProjects={expandedProjects}
              selectionMode={selectionMode}
              selectedRunIds={selectedRunIds}
              onToggleYear={toggleYear}
              onToggleWeek={toggleWeek}
              onToggleProject={toggleProject}
              onOpenRun={(runId) => router.push(`/bautagebuch/run/${runId}`)}
              onToggleSelect={toggleSelection}
              onRename={openRename}
              onDelete={deleteRun}
            />
          )}
        </Section>
      ) : null}

      <Modal visible={Boolean(renameRunId)} transparent animationType="fade">
        <View style={styles.modalBackdrop}>
          <View style={styles.modalCard}>
            <Text style={styles.modalTitle}>BTB umbenennen</Text>
            <TextField label="Titel" value={renameTitle} onChangeText={setRenameTitle} />
            <View style={styles.modalActions}>
              <PrimaryButton label="Abbrechen" variant="ghost" onPress={() => setRenameRunId(null)} />
              <PrimaryButton label="Speichern" onPress={() => void submitRename()} />
            </View>
          </View>
        </View>
      </Modal>
    </Screen>
  );
}

const styles = StyleSheet.create({
  center: { alignItems: 'center', gap: spacing.sm, paddingVertical: spacing.xl },
  muted: { ...typography.caption, color: colors.muted },
  errorBox: { gap: spacing.sm },
  error: { ...typography.body, color: colors.danger },
  listActions: { flexDirection: 'row', gap: spacing.xs },
  modalBackdrop: {
    flex: 1,
    backgroundColor: colors.overlay,
    justifyContent: 'center',
    padding: spacing.xl
  },
  modalCard: {
    backgroundColor: colors.panel,
    borderRadius: spacing.cardRadius,
    padding: spacing.md,
    gap: spacing.sm,
    borderWidth: 1,
    borderColor: colors.border
  },
  modalTitle: { ...typography.subtitle, color: colors.ink },
  modalActions: { flexDirection: 'row', justifyContent: 'flex-end', gap: spacing.xs }
});

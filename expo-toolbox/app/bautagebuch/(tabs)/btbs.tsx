import { useEffect, useMemo, useRef, useState } from 'react';
import { ActivityIndicator, Alert, Modal, StyleSheet, Text, View } from 'react-native';
import { useRouter } from 'expo-router';

import { EmptyState, PrimaryButton, Screen, TextField } from '../../../src/components/mobile';
import { colors, spacing, typography } from '../../../src/constants/theme';
import { useToast } from '../../../src/contexts/ToastContext';
import { BTBFilterBar } from '../../../src/native/bautagebuch/components/btb-list/BTBFilterBar';
import { BTBGroupedList } from '../../../src/native/bautagebuch/components/btb-list/BTBGroupedList';
import { ProjectFilterList } from '../../../src/native/bautagebuch/components/btb-list/ProjectFilterList';
import { useBautagebuchWorkspace } from '../../../src/native/bautagebuch/hooks/useBautagebuchWorkspace';
import {
  buildCalendarTree,
  buildProjectFirstTree,
  DEFAULT_BTB_FILTERS,
  listProjectsFromRuns,
  type BtbListFilters
} from '../../../src/native/bautagebuch/lib/btb-filter';
import { filterRunsBySearchQuery } from '../../../src/native/bautagebuch/lib/btb-search';
import type { BautagebuchRun } from '../../../src/native/bautagebuch/types';

function buildSubtitle(totalCount: number, visibleCount: number, searchQuery: string): string {
  const totalLabel = `${totalCount} BTB${totalCount === 1 ? '' : 's'}`;
  if (!searchQuery.trim()) return totalLabel;
  return `${visibleCount} von ${totalLabel}`;
}

export default function BautagebuchBtbsScreen() {
  const router = useRouter();
  const { showToast } = useToast();
  const { initialLoading, refreshing, error, runs, setupModel, load, deleteRunById, renameRunById } =
    useBautagebuchWorkspace();

  const [filters, setFilters] = useState<BtbListFilters>(DEFAULT_BTB_FILTERS);
  const [searchQuery, setSearchQuery] = useState('');
  const [expandedYears, setExpandedYears] = useState<Set<number>>(new Set());
  const [expandedWeeks, setExpandedWeeks] = useState<Set<string>>(new Set());
  const [expandedProjects, setExpandedProjects] = useState<Set<string>>(new Set());
  const [renameRunId, setRenameRunId] = useState<string | null>(null);
  const [renameTitle, setRenameTitle] = useState('');
  const expansionInitialized = useRef(false);

  const visibleRuns = useMemo(
    () => filterRunsBySearchQuery(runs, searchQuery),
    [runs, searchQuery]
  );
  const projectOptions = useMemo(
    () => listProjectsFromRuns(visibleRuns, setupModel),
    [visibleRuns, setupModel]
  );
  const calendarTree = useMemo(
    () => buildCalendarTree(visibleRuns, setupModel, filters),
    [visibleRuns, setupModel, filters]
  );
  const projectTree = useMemo(() => {
    if (filters.groupMode !== 'project' || !filters.projectKey) return null;
    return buildProjectFirstTree(visibleRuns, setupModel, filters.projectKey, filters);
  }, [visibleRuns, setupModel, filters]);

  useEffect(() => {
    expansionInitialized.current = false;
    setExpandedYears(new Set());
    setExpandedWeeks(new Set());
    setExpandedProjects(new Set());
  }, [filters.groupMode, filters.projectKey, searchQuery]);

  useEffect(() => {
    if (visibleRuns.length === 0 || expansionInitialized.current) return;
    expansionInitialized.current = true;
  }, [visibleRuns.length, calendarTree, projectTree]);

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

  const deleteRun = (runId: string) => {
    Alert.alert('BTB löschen', 'Dieses Bautagebuch wirklich löschen?', [
      { text: 'Abbrechen', style: 'cancel' },
      {
        text: 'Löschen',
        style: 'destructive',
        onPress: () => {
          void deleteRunById(runId).then(() => showToast('BTB gelöscht'));
        }
      }
    ]);
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

  const showProjectList = filters.groupMode === 'project' && !filters.projectKey;
  const hasSearchQuery = searchQuery.trim().length > 0;
  const subtitle = buildSubtitle(runs.length, visibleRuns.length, searchQuery);

  return (
    <Screen
      title="Bautagebücher"
      subtitle={subtitle}
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
        <View style={styles.content}>
          {runs.length > 0 ? (
            <>
              <TextField
                label="Suche"
                value={searchQuery}
                onChangeText={setSearchQuery}
                placeholder="Begriff im BTB suchen, z. B. Zaun"
                autoCapitalize="none"
                autoCorrect={false}
                returnKeyType="search"
                clearButtonMode="while-editing"
              />
              <BTBFilterBar filters={filters} onChange={setFilters} />
            </>
          ) : null}

          {filters.groupMode === 'project' && filters.projectKey ? (
            <PrimaryButton
              label="← Alle Projekte"
              variant="ghost"
              onPress={() => setFilters((current) => ({ ...current, projectKey: null }))}
            />
          ) : null}

          {runs.length === 0 ? (
            <EmptyState
              title="Noch keine BTB-Läufe"
              description="Wechsle zum Home-Tab, um dein erstes Bautagebuch zu starten."
            />
          ) : visibleRuns.length === 0 && hasSearchQuery ? (
            <EmptyState
              title="Keine Treffer"
              description={`Kein Bautagebuch enthält „${searchQuery.trim()}“.`}
            />
          ) : showProjectList ? (
            <ProjectFilterList
              projects={projectOptions}
              onSelectProject={(projectKey) =>
                setFilters((current) => ({ ...current, projectKey }))
              }
            />
          ) : filters.groupMode === 'project' && projectTree ? (
            <BTBGroupedList
              mode="project"
              tree={projectTree}
              expandedYears={expandedYears}
              expandedWeeks={expandedWeeks}
              onToggleYear={toggleYear}
              onToggleWeek={toggleWeek}
              onOpenRun={(runId) => router.push(`/bautagebuch/run/${runId}`)}
              onRename={openRename}
              onDelete={deleteRun}
            />
          ) : (
            <BTBGroupedList
              mode="calendar"
              tree={calendarTree}
              expandedYears={expandedYears}
              expandedWeeks={expandedWeeks}
              expandedProjects={expandedProjects}
              onToggleYear={toggleYear}
              onToggleWeek={toggleWeek}
              onToggleProject={toggleProject}
              onOpenRun={(runId) => router.push(`/bautagebuch/run/${runId}`)}
              onRename={openRename}
              onDelete={deleteRun}
            />
          )}
        </View>
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
  content: { gap: spacing.md },
  center: { alignItems: 'center', gap: spacing.sm, paddingVertical: spacing.xl },
  muted: { ...typography.caption, color: colors.muted },
  errorBox: { gap: spacing.sm },
  error: { ...typography.body, color: colors.danger },
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

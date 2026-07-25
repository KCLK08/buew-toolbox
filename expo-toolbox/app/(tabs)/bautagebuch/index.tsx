import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import {
  ActivityIndicator,
  Alert,
  Modal,
  StyleSheet,
  Text,
  View
} from 'react-native';
import { useRouter, useFocusEffect } from 'expo-router';

import {
  Card,
  EmptyState,
  Fab,
  PrimaryButton,
  Screen,
  Section,
  TextField
} from '../../../src/components/mobile';
import { colors, spacing, typography } from '../../../src/constants/theme';
import { useToast } from '../../../src/contexts/ToastContext';
import { BautagebuchRunList } from '../../../src/native/bautagebuch/components/BautagebuchRunList';
import { groupRunsByCalendar, projectGroupKey } from '../../../src/native/bautagebuch/lib/group-runs-by-calendar';
import {
  createRun,
  deleteRunCascade,
  listRuns,
  renameRun,
  updateRun
} from '../../../src/native/bautagebuch/db/database';
import { applyRunDefaultsFromModel } from '../../../src/native/bautagebuch/lib/run-defaults';
import { getActiveTemplateBundle } from '../../../src/native/bautagebuch/services/templateService';
import type { BautagebuchRun } from '../../../src/native/bautagebuch/types';

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

function buildBtbTitle(siteName: string): string {
  const today = new Date().toISOString().slice(0, 10);
  return `BTB ${today} - ${siteName.trim() || 'Baustelle'}`;
}

export default function BautagebuchHomeScreen() {
  const router = useRouter();
  const { showToast } = useToast();
  const [initialLoading, setInitialLoading] = useState(true);
  const [refreshing, setRefreshing] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [runs, setRuns] = useState<BautagebuchRun[]>([]);
  const [newName, setNewName] = useState('');
  const [templateId, setTemplateId] = useState('');
  const [templateName, setTemplateName] = useState('');
  const [templateReady, setTemplateReady] = useState(false);
  const [setupModel, setSetupModel] = useState<Record<string, unknown> | null>(null);
  const [creating, setCreating] = useState(false);
  const [selectionMode, setSelectionMode] = useState(false);
  const [selectedRunIds, setSelectedRunIds] = useState<string[]>([]);
  const [expandedYears, setExpandedYears] = useState<Set<number>>(new Set());
  const [expandedWeeks, setExpandedWeeks] = useState<Set<string>>(new Set());
  const [expandedProjects, setExpandedProjects] = useState<Set<string>>(new Set());
  const [renameRunId, setRenameRunId] = useState<string | null>(null);
  const [renameTitle, setRenameTitle] = useState('');

  const btbPreviewTitle = useMemo(() => (newName.trim() ? buildBtbTitle(newName) : ''), [newName]);
  const expansionInitialized = useRef(false);
  const isFirstFocus = useRef(true);

  const load = useCallback(async (mode: 'initial' | 'refresh' = 'initial') => {
    if (mode === 'initial') setInitialLoading(true);
    else setRefreshing(true);
    setError(null);
    try {
      const bundle = await getActiveTemplateBundle();
      setTemplateId(bundle.template.templateId);
      setTemplateName(bundle.template.templateName);
      setTemplateReady(bundle.template.status === 'ready');
      setSetupModel(bundle.setupModel);
      setRuns(await listRuns());
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Bautagebuch konnte nicht geladen werden.');
    } finally {
      if (mode === 'initial') setInitialLoading(false);
      else setRefreshing(false);
    }
  }, []);

  useEffect(() => {
    void load('initial');
  }, [load]);

  useFocusEffect(
    useCallback(() => {
      if (isFirstFocus.current) {
        isFirstFocus.current = false;
        return;
      }
      void load('refresh');
    }, [load])
  );

  const runTree = useMemo(
    () => groupRunsByCalendar(runs, { setupModel }),
    [runs, setupModel]
  );

  useEffect(() => {
    if (runs.length === 0) {
      expansionInitialized.current = false;
      setExpandedYears(new Set());
      setExpandedWeeks(new Set());
      setExpandedProjects(new Set());
      return;
    }
    if (expansionInitialized.current) return;

    setExpandedYears(new Set(runTree.years.map((yearGroup) => yearGroup.year)));
    const weekKeys = runTree.years.flatMap((yearGroup) => yearGroup.weeks.map((week) => week.weekKey));
    setExpandedWeeks(new Set(weekKeys));
    const projectKeys = runTree.years.flatMap((yearGroup) =>
      yearGroup.weeks.flatMap((week) =>
        week.projects.map((project) => projectGroupKey(week.weekKey, project.projectKey))
      )
    );
    setExpandedProjects(new Set(projectKeys));
    expansionInitialized.current = true;
  }, [runTree, runs.length]);

  const startRun = async () => {
    if (!templateId || !setupModel || !templateReady) {
      Alert.alert(
        'Vorlage nicht bereit',
        'Die aktive Vorlage ist noch nicht startbereit. Bitte Setup abschließen.'
      );
      return;
    }
    setCreating(true);
    try {
      const title = buildBtbTitle(newName);
      let run = await createRun({
        templateId,
        title,
        setupVersion: Number(setupModel.version || 1)
      });
      const defaults = applyRunDefaultsFromModel(setupModel, run.values);
      if (defaults.changed) {
        const updated = await updateRun(run.runId, { values: defaults.values });
        if (updated) run = updated;
      }
      setNewName('');
      router.push(`/bautagebuch/run/${run.runId}`);
    } catch (err) {
      Alert.alert('Fehler', err instanceof Error ? err.message : 'BTB konnte nicht erstellt werden.');
    } finally {
      setCreating(false);
    }
  };

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
          void deleteRunCascade(runId).then(() => {
            setSelectedRunIds((current) => current.filter((id) => id !== runId));
            showToast('BTB gelöscht');
            void load('refresh');
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
            void Promise.all(selectedRunIds.map((runId) => deleteRunCascade(runId))).then(() => {
              setSelectedRunIds([]);
              setSelectionMode(false);
              showToast('Auswahl gelöscht');
              void load('refresh');
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
    const updated = await renameRun(renameRunId, renameTitle);
    if (!updated) {
      Alert.alert('Umbenennen', 'Bitte einen gültigen Titel eingeben.');
      return;
    }
    setRenameRunId(null);
    setRenameTitle('');
    showToast('BTB umbenannt');
    void load('refresh');
  };

  return (
    <Screen
      title={formatGreeting()}
      subtitle={formatTodayLabel()}
      scroll
      refreshing={refreshing}
      onRefresh={() => void load('refresh')}
    >
      {initialLoading ? (
        <View style={styles.center}>
          <ActivityIndicator color={colors.accent} />
          <Text style={styles.muted}>Vorlage und Daten werden geladen…</Text>
        </View>
      ) : null}

      {error ? (
        <Card style={styles.errorCard}>
          <Text style={styles.error}>{error}</Text>
          <PrimaryButton label="Erneut versuchen" variant="secondary" onPress={() => void load('initial')} />
        </Card>
      ) : null}

      {!initialLoading && !error ? (
        <>
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

          <Section
            title="Deine Bautagebücher"
            action={
              selectionMode ? (
                <View style={styles.listActions}>
                  <PrimaryButton
                    label="Abbrechen"
                    variant="ghost"
                    onPress={() => {
                      setSelectionMode(false);
                      setSelectedRunIds([]);
                    }}
                  />
                  <PrimaryButton
                    label={`Löschen (${selectedRunIds.length})`}
                    variant="secondary"
                    disabled={selectedRunIds.length === 0}
                    onPress={deleteSelected}
                  />
                </View>
              ) : (
                <PrimaryButton
                  label="Auswahl"
                  variant="ghost"
                  disabled={runs.length === 0}
                  onPress={() => setSelectionMode(true)}
                />
              )
            }
          >
            {runs.length === 0 ? (
              <EmptyState
                title="Noch keine BTB-Läufe"
                description="Tippe oben auf „BTB jetzt starten“, um dein erstes elektronisches Bautagebuch zu erfassen."
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

          <Section title="Vorlage">
            <Card style={styles.toolsCard}>
              <PrimaryButton
                label="Setup-Editor öffnen"
                variant="secondary"
                disabled={!templateId}
                onPress={() => router.push('/bautagebuch/setup')}
              />
            </Card>
          </Section>
        </>
      ) : null}

      {!initialLoading && !error && templateReady && newName.trim() ? (
        <Fab label="+" onPress={() => void startRun()} accessibilityLabel="Neues BTB" />
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
  previewTitle: { ...typography.bodyStrong, color: colors.accent },
  listActions: { flexDirection: 'row', gap: spacing.xs },
  toolsCard: { gap: spacing.sm },
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

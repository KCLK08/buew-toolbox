import { useCallback, useEffect, useMemo, useState } from 'react';
import {
  ActivityIndicator,
  Alert,
  Modal,
  Pressable,
  StyleSheet,
  Text,
  View
} from 'react-native';
import { useRouter } from 'expo-router';

import {
  Card,
  EmptyState,
  Fab,
  PrimaryButton,
  Screen,
  Section,
  StatCard,
  TextField
} from '../../../src/components/mobile';
import { colors, spacing, typography } from '../../../src/constants/theme';
import { useToast } from '../../../src/contexts/ToastContext';
import { BautagebuchRunCard } from '../../../src/native/bautagebuch/components/BautagebuchRunCard';
import {
  createRun,
  deleteRunCascade,
  listExports,
  listRuns,
  renameRun,
  updateRun
} from '../../../src/native/bautagebuch/db/database';
import { applyRunDefaultsFromModel } from '../../../src/native/bautagebuch/lib/run-defaults';
import { deleteCachedExport, shareCachedExport } from '../../../src/native/bautagebuch/services/exportService';
import {
  exportBautagebuchBackupZip,
  pickAndRestoreBautagebuchBackup
} from '../../../src/native/bautagebuch/services/backupExportService';
import { ensureBuiltinTemplate } from '../../../src/native/bautagebuch/services/templateService';
import type { BautagebuchExport, BautagebuchRun } from '../../../src/native/bautagebuch/types';

function groupRunsByWeek(runs: BautagebuchRun[]) {
  const groups = new Map<string, BautagebuchRun[]>();
  for (const run of runs) {
    const date = run.title.match(/\d{4}-\d{2}-\d{2}/)?.[0] || run.createdAt.slice(0, 10);
    const weekKey = `KW ${getIsoWeek(date)} · ${date.slice(0, 7)}`;
    const bucket = groups.get(weekKey) || [];
    bucket.push(run);
    groups.set(weekKey, bucket);
  }
  return [...groups.entries()];
}

function getIsoWeek(dateString: string): number {
  const date = new Date(`${dateString}T12:00:00`);
  const target = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  const day = target.getUTCDay() || 7;
  target.setUTCDate(target.getUTCDate() + 4 - day);
  const yearStart = new Date(Date.UTC(target.getUTCFullYear(), 0, 1));
  return Math.ceil((((target.getTime() - yearStart.getTime()) / 86400000) + 1) / 7);
}

function formatTodayLabel(): string {
  return new Date().toLocaleDateString('de-DE', {
    weekday: 'long',
    day: 'numeric',
    month: 'long'
  });
}

export default function BautagebuchHomeScreen() {
  const router = useRouter();
  const { showToast } = useToast();
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [runs, setRuns] = useState<BautagebuchRun[]>([]);
  const [newName, setNewName] = useState('');
  const [templateId, setTemplateId] = useState('');
  const [setupModel, setSetupModel] = useState<Record<string, unknown> | null>(null);
  const [creating, setCreating] = useState(false);
  const [exportsList, setExportsList] = useState<BautagebuchExport[]>([]);
  const [sharingExport, setSharingExport] = useState<string | null>(null);
  const [selectionMode, setSelectionMode] = useState(false);
  const [selectedRunIds, setSelectedRunIds] = useState<string[]>([]);
  const [expandedWeeks, setExpandedWeeks] = useState<Set<string>>(new Set());
  const [renameRunId, setRenameRunId] = useState<string | null>(null);
  const [renameTitle, setRenameTitle] = useState('');
  const [backupBusy, setBackupBusy] = useState(false);
  const [restoreBusy, setRestoreBusy] = useState(false);
  const [toolsOpen, setToolsOpen] = useState(false);

  const load = useCallback(async () => {
    setLoading(true);
    setError(null);
    try {
      const bundle = await ensureBuiltinTemplate();
      setTemplateId(bundle.templateId);
      setSetupModel(bundle.setupModel);
      setRuns(await listRuns(bundle.templateId));
      setExportsList(await listExports());
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Bautagebuch konnte nicht geladen werden.');
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    void load();
  }, [load]);

  const grouped = useMemo(() => groupRunsByWeek(runs), [runs]);
  const openRuns = useMemo(() => runs.filter((run) => run.status !== 'completed').length, [runs]);

  useEffect(() => {
    setExpandedWeeks(new Set(grouped.map(([week]) => week)));
  }, [grouped.length]);

  const startRun = async () => {
    if (!templateId || !setupModel) return;
    setCreating(true);
    try {
      const today = new Date().toISOString().slice(0, 10);
      const title = `BTB ${today} - ${newName.trim() || 'Baustelle'}`;
      let run = await createRun({ templateId, title, setupVersion: 6 });
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

  const toggleWeek = (week: string) => {
    setExpandedWeeks((current) => {
      const next = new Set(current);
      if (next.has(week)) next.delete(week);
      else next.add(week);
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
            showToast('BTB gelöscht');
            void load();
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
              void load();
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
    void load();
  };

  const shareExport = async (exportId: string) => {
    setSharingExport(exportId);
    try {
      await shareCachedExport(exportId);
    } catch (err) {
      Alert.alert('Export', err instanceof Error ? err.message : 'Teilen fehlgeschlagen.');
    } finally {
      setSharingExport(null);
    }
  };

  const exportBackup = async () => {
    setBackupBusy(true);
    try {
      await exportBautagebuchBackupZip();
      showToast('Backup erstellt');
    } catch (err) {
      Alert.alert('Backup', err instanceof Error ? err.message : 'Backup fehlgeschlagen.');
    } finally {
      setBackupBusy(false);
    }
  };

  const restoreBackup = () => {
    Alert.alert(
      'Backup wiederherstellen',
      'Das aktuelle Bautagebuch wird durch das ZIP-Backup ersetzt (Datenbank, Vorlagen, Fotos, Exporte). Fortfahren?',
      [
        { text: 'Abbrechen', style: 'cancel' },
        {
          text: 'Wiederherstellen',
          style: 'destructive',
          onPress: () => {
            void (async () => {
              setRestoreBusy(true);
              try {
                const result = await pickAndRestoreBautagebuchBackup();
                if (!result) return;
                showToast(
                  `Backup wiederhergestellt (${result.photoFileCount} Fotos, ${result.exportFileCount} Exporte)`
                );
                await load();
              } catch (err) {
                Alert.alert(
                  'Wiederherstellung',
                  err instanceof Error ? err.message : 'Wiederherstellung fehlgeschlagen.'
                );
              } finally {
                setRestoreBusy(false);
              }
            })();
          }
        }
      ]
    );
  };

  const removeExport = (exportId: string) => {
    Alert.alert('Export löschen', 'Gespeicherten Export wirklich entfernen?', [
      { text: 'Abbrechen', style: 'cancel' },
      {
        text: 'Löschen',
        style: 'destructive',
        onPress: () => {
          void deleteCachedExport(exportId).then(load);
        }
      }
    ]);
  };

  return (
    <Screen
      title="Bautagebuch"
      subtitle="Elektronisches Bautagebuch (eBTB)"
      scroll
      refreshing={loading}
      onRefresh={load}
    >
      {loading ? (
        <View style={styles.center}>
          <ActivityIndicator color={colors.accent} />
          <Text style={styles.muted}>Vorlage und Daten werden geladen…</Text>
        </View>
      ) : null}

      {error ? (
        <Card style={styles.errorCard}>
          <Text style={styles.error}>{error}</Text>
          <PrimaryButton label="Erneut versuchen" variant="secondary" onPress={() => void load()} />
        </Card>
      ) : null}

      {!loading && !error ? (
        <>
          <View style={styles.statsRow}>
            <StatCard title="Offen" value={String(openRuns)} icon="📝" />
            <StatCard title="Gesamt" value={String(runs.length)} icon="📚" />
            <StatCard title="Exporte" value={String(exportsList.length)} icon="📤" />
          </View>

          <Card>
            <Text style={styles.heroDate}>{formatTodayLabel()}</Text>
            <Text style={styles.heroTitle}>Neues Bautagebuch starten</Text>
            <Text style={styles.heroHint}>
              Gib der Baustelle einen kurzen Namen — Datum und Vorlage werden automatisch gesetzt.
            </Text>
            <TextField
              label="Baustelle / Strecke"
              hint="z. B. Strecke Nord, Tunnel Süd"
              value={newName}
              onChangeText={setNewName}
            />
            <PrimaryButton
              label={creating ? 'Wird erstellt…' : 'BTB jetzt starten'}
              disabled={creating || !templateId}
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
              grouped.map(([week, weekRuns]) => {
                const expanded = expandedWeeks.has(week);
                return (
                  <View key={week} style={styles.weekGroup}>
                    <Pressable style={styles.weekHeader} onPress={() => toggleWeek(week)}>
                      <Text style={styles.weekTitle}>{week}</Text>
                      <Text style={styles.weekToggle}>
                        {expanded ? '▾' : '▸'} {weekRuns.length}
                      </Text>
                    </Pressable>
                    {expanded
                      ? weekRuns.map((run) => (
                          <BautagebuchRunCard
                            key={run.runId}
                            run={run}
                            selectionMode={selectionMode}
                            selected={selectedRunIds.includes(run.runId)}
                            onPress={() => router.push(`/bautagebuch/run/${run.runId}`)}
                            onToggleSelect={() => toggleSelection(run.runId)}
                            onRename={() => openRename(run)}
                            onDelete={() => deleteRun(run.runId)}
                          />
                        ))
                      : null}
                  </View>
                );
              })
            )}
          </Section>

          {exportsList.length > 0 ? (
            <Section title={`Letzte Exporte (${exportsList.length})`}>
              {exportsList.slice(0, 5).map((item) => (
                <Card key={item.exportId} padded style={styles.exportRow}>
                  <Text style={styles.exportName} numberOfLines={2}>
                    {item.fileName}
                  </Text>
                  <View style={styles.exportActions}>
                    <PrimaryButton
                      label={sharingExport === item.exportId ? 'Teilen…' : 'PDF teilen'}
                      variant="secondary"
                      disabled={Boolean(sharingExport)}
                      onPress={() => void shareExport(item.exportId)}
                    />
                    <PrimaryButton label="Löschen" variant="ghost" onPress={() => removeExport(item.exportId)} />
                  </View>
                </Card>
              ))}
            </Section>
          ) : null}

          <Section title="Werkzeuge">
            <Pressable style={styles.toolsToggle} onPress={() => setToolsOpen((value) => !value)}>
              <Text style={styles.toolsToggleLabel}>
                {toolsOpen ? 'Weniger anzeigen' : 'Setup, Backup & Wiederherstellung'}
              </Text>
              <Text style={styles.weekToggle}>{toolsOpen ? '▾' : '▸'}</Text>
            </Pressable>
            {toolsOpen ? (
              <Card style={styles.toolsCard}>
                <PrimaryButton
                  label="Setup-Editor öffnen"
                  variant="secondary"
                  disabled={!templateId}
                  onPress={() => router.push('/bautagebuch/setup')}
                />
                <PrimaryButton
                  label={backupBusy ? 'Backup wird erstellt…' : 'Backup exportieren (ZIP)'}
                  variant="secondary"
                  disabled={backupBusy || restoreBusy || !templateId}
                  onPress={() => void exportBackup()}
                />
                <PrimaryButton
                  label={restoreBusy ? 'Wird wiederhergestellt…' : 'Backup wiederherstellen'}
                  variant="secondary"
                  disabled={backupBusy || restoreBusy}
                  onPress={restoreBackup}
                />
              </Card>
            ) : null}
          </Section>
        </>
      ) : null}

      <Fab label="+" onPress={() => void startRun()} accessibilityLabel="Neues BTB" />

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
  statsRow: { flexDirection: 'row', gap: spacing.sm },
  heroDate: { ...typography.caption, color: colors.accent, textTransform: 'capitalize' },
  heroTitle: { ...typography.subtitle, color: colors.ink },
  heroHint: { ...typography.caption, color: colors.muted, marginBottom: spacing.xs },
  listActions: { flexDirection: 'row', gap: spacing.xs },
  weekGroup: { gap: spacing.sm },
  weekHeader: {
    flexDirection: 'row',
    justifyContent: 'space-between',
    alignItems: 'center',
    minHeight: spacing.touchMin,
    paddingHorizontal: spacing.xxs
  },
  weekTitle: { ...typography.label, color: colors.muted },
  weekToggle: { ...typography.caption, color: colors.muted },
  exportRow: { gap: spacing.sm },
  exportName: { ...typography.body, color: colors.ink },
  exportActions: { flexDirection: 'row', flexWrap: 'wrap', gap: spacing.xs },
  toolsToggle: {
    flexDirection: 'row',
    alignItems: 'center',
    justifyContent: 'space-between',
    minHeight: spacing.touchMin,
    paddingHorizontal: spacing.xxs
  },
  toolsToggleLabel: { ...typography.bodyStrong, color: colors.accent2 },
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

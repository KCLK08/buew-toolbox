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

import { EmptyState, Fab, PrimaryButton, Screen, TextField } from '../../../src/components/mobile';
import { colors, typography } from '../../../src/constants/theme';
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
    <Screen title="Bautagebuch" subtitle="Elektronisches Bautagebuch (eBTB)" scroll refreshing={loading} onRefresh={load}>
      {loading ? (
        <View style={styles.center}>
          <ActivityIndicator color={colors.accent} />
          <Text style={styles.muted}>Vorlage und Daten werden geladen…</Text>
        </View>
      ) : null}

      {error ? <Text style={styles.error}>{error}</Text> : null}

      <View style={styles.newCard}>
        <Text style={styles.cardTitle}>Neues BTB starten</Text>
        <TextField
          label="Bezeichnung"
          hint="z. B. Strecke Nord"
          value={newName}
          onChangeText={setNewName}
        />
        <PrimaryButton
          label={creating ? 'Wird erstellt…' : 'BTB starten'}
          disabled={creating || !templateId}
          onPress={() => void startRun()}
        />
        <PrimaryButton
          label="Setup-Editor öffnen"
          variant="secondary"
          disabled={!templateId}
          onPress={() => router.push('/bautagebuch/setup')}
        />
      </View>

      {exportsList.length > 0 ? (
        <View style={styles.exportsCard}>
          <Text style={styles.cardTitle}>Exporte ({exportsList.length})</Text>
          {exportsList.map((item) => (
            <View key={item.exportId} style={styles.exportRow}>
              <Text style={styles.exportName}>{item.fileName}</Text>
              <View style={styles.exportActions}>
                <PrimaryButton
                  label={sharingExport === item.exportId ? 'Teilen…' : 'PDF teilen'}
                  variant="secondary"
                  disabled={Boolean(sharingExport)}
                  onPress={() => void shareExport(item.exportId)}
                />
                <PrimaryButton label="Löschen" variant="ghost" onPress={() => removeExport(item.exportId)} />
              </View>
            </View>
          ))}
        </View>
      ) : null}

      <View style={styles.listHeader}>
        <Text style={styles.cardTitle}>Bautagebücher</Text>
        <View style={styles.listActions}>
          {selectionMode ? (
            <>
              <PrimaryButton label="Abbrechen" variant="ghost" onPress={() => {
                setSelectionMode(false);
                setSelectedRunIds([]);
              }} />
              <PrimaryButton
                label={`Löschen (${selectedRunIds.length})`}
                variant="secondary"
                disabled={selectedRunIds.length === 0}
                onPress={deleteSelected}
              />
            </>
          ) : (
            <PrimaryButton
              label="Auswahl"
              variant="ghost"
              disabled={runs.length === 0}
              onPress={() => setSelectionMode(true)}
            />
          )}
        </View>
      </View>

      {runs.length === 0 && !loading ? (
        <EmptyState
          title="Noch keine BTB-Läufe"
          description="Starte ein neues elektronisches Bautagebuch mit der Vorlage-eBTB."
        />
      ) : null}

      {grouped.map(([week, weekRuns]) => {
        const expanded = expandedWeeks.has(week);
        return (
          <View key={week} style={styles.weekGroup}>
            <Pressable style={styles.weekHeader} onPress={() => toggleWeek(week)}>
              <Text style={styles.weekTitle}>{week}</Text>
              <Text style={styles.weekToggle}>{expanded ? '▾' : '▸'} {weekRuns.length}</Text>
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
      })}

      <Fab label="+" onPress={() => void startRun()} accessibilityLabel="Neues BTB" />

      <Modal visible={Boolean(renameRunId)} transparent animationType="fade">
        <View style={styles.modalBackdrop}>
          <View style={styles.modalCard}>
            <Text style={styles.cardTitle}>BTB umbenennen</Text>
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
  center: { alignItems: 'center', gap: 8, paddingVertical: 24 },
  muted: { ...typography.caption, color: colors.muted },
  error: { ...typography.body, color: colors.danger },
  newCard: { gap: 12, marginBottom: 8 },
  cardTitle: { ...typography.bodyStrong, color: colors.ink },
  exportsCard: { gap: 12, marginBottom: 8 },
  exportRow: { gap: 8, padding: 12, borderRadius: 12, backgroundColor: colors.panel, borderWidth: 1, borderColor: colors.border },
  exportName: { ...typography.body, color: colors.ink },
  exportActions: { flexDirection: 'row', flexWrap: 'wrap', gap: 8 },
  listHeader: { flexDirection: 'row', alignItems: 'center', justifyContent: 'space-between', marginTop: 8 },
  listActions: { flexDirection: 'row', gap: 8 },
  weekGroup: { gap: 8, marginTop: 8 },
  weekHeader: { flexDirection: 'row', justifyContent: 'space-between', alignItems: 'center' },
  weekTitle: { ...typography.label, color: colors.muted },
  weekToggle: { ...typography.caption, color: colors.muted },
  modalBackdrop: {
    flex: 1,
    backgroundColor: 'rgba(0,0,0,0.35)',
    justifyContent: 'center',
    padding: 24
  },
  modalCard: {
    backgroundColor: colors.panel,
    borderRadius: 16,
    padding: 16,
    gap: 12,
    borderWidth: 1,
    borderColor: colors.border
  },
  modalActions: { flexDirection: 'row', justifyContent: 'flex-end', gap: 8 }
});

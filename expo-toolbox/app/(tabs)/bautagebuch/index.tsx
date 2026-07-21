import { useCallback, useEffect, useState } from 'react';
import { ActivityIndicator, Alert, StyleSheet, Text, View } from 'react-native';
import { useRouter } from 'expo-router';

import { EmptyState, Fab, ListItem, PrimaryButton, Screen, TextField } from '../../../src/components/mobile';
import { colors, typography } from '../../../src/constants/theme';
import { createRun, listExports, listRuns } from '../../../src/native/bautagebuch/db/database';
import { ensureBuiltinTemplate } from '../../../src/native/bautagebuch/services/templateService';
import { deleteCachedExport, shareCachedExport } from '../../../src/native/bautagebuch/services/exportService';
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
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [runs, setRuns] = useState<BautagebuchRun[]>([]);
  const [newName, setNewName] = useState('');
  const [templateId, setTemplateId] = useState('');
  const [creating, setCreating] = useState(false);
  const [exportsList, setExportsList] = useState<BautagebuchExport[]>([]);
  const [sharingExport, setSharingExport] = useState<string | null>(null);

  const load = useCallback(async () => {
    setLoading(true);
    setError(null);
    try {
      const bundle = await ensureBuiltinTemplate();
      setTemplateId(bundle.templateId);
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

  const startRun = async () => {
    if (!templateId) return;
    setCreating(true);
    try {
      const today = new Date().toISOString().slice(0, 10);
      const title = `BTB ${today} - ${newName.trim() || 'Baustelle'}`;
      const run = await createRun({ templateId, title, setupVersion: 6 });
      setNewName('');
      router.push(`/bautagebuch/run/${run.runId}`);
    } catch (err) {
      Alert.alert('Fehler', err instanceof Error ? err.message : 'BTB konnte nicht erstellt werden.');
    } finally {
      setCreating(false);
    }
  };

  const grouped = groupRunsByWeek(runs);

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

      {runs.length === 0 && !loading ? (
        <EmptyState
          title="Noch keine BTB-Läufe"
          description="Starte ein neues elektronisches Bautagebuch mit der Vorlage-eBTB."
        />
      ) : null}

      {grouped.map(([week, weekRuns]) => (
        <View key={week} style={styles.weekGroup}>
          <Text style={styles.weekTitle}>{week}</Text>
          {weekRuns.map((run) => (
            <ListItem
              key={run.runId}
              title={run.title}
              subtitle={run.status === 'completed' ? 'Abgeschlossen' : 'Entwurf'}
              meta={new Date(run.updatedAt).toLocaleString('de-DE')}
              onPress={() => router.push(`/bautagebuch/run/${run.runId}`)}
            />
          ))}
        </View>
      ))}

      <Fab label="+" onPress={() => void startRun()} accessibilityLabel="Neues BTB" />
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
  weekGroup: { gap: 8, marginTop: 8 },
  weekTitle: { ...typography.label, color: colors.muted, marginTop: 8 }
});

import { useCallback, useEffect, useMemo, useState } from 'react';
import { ActivityIndicator, Alert, Pressable, ScrollView, StyleSheet, Text, TextInput, View } from 'react-native';
import { useFocusEffect, useRouter } from 'expo-router';

import { colors } from '@/theme/colors';
import { ensureBuiltinTemplate } from '@/lib/bootstrap';
import { createRun, deleteRunCascade, getSetupModel, listRuns, updateRun } from '@/lib/db';
import { applyRunDefaultsFromModel, buildRunTitleByConvention, formatWeekLabel, getWeekKey, normalizeRunNameInput } from '@/lib/run-utils';
import type { Run } from '@/types';

export default function HomeScreen() {
  const router = useRouter();
  const [loading, setLoading] = useState(true);
  const [busy, setBusy] = useState(false);
  const [runs, setRuns] = useState<Run[]>([]);
  const [templateId, setTemplateId] = useState<string | null>(null);
  const [newRunName, setNewRunName] = useState('');
  const [error, setError] = useState('');
  const [info, setInfo] = useState('');

  const loadData = useCallback(async () => {
    setLoading(true);
    setError('');
    try {
      const id = await ensureBuiltinTemplate();
      setTemplateId(id);
      setRuns(await listRuns());
    } catch (e) {
      setError((e as Error).message || 'Daten konnten nicht geladen werden.');
    } finally {
      setLoading(false);
    }
  }, []);

  useFocusEffect(
    useCallback(() => {
      loadData();
    }, [loadData])
  );

  const weekGroups = useMemo(() => {
    const groups = new Map<string, Run[]>();
    for (const run of runs) {
      const key = getWeekKey(run.updatedAt || run.createdAt);
      if (!groups.has(key)) groups.set(key, []);
      groups.get(key)!.push(run);
    }
    return [...groups.entries()].sort((a, b) => b[0].localeCompare(a[0]));
  }, [runs]);

  async function startNewRun() {
    if (!templateId) return;
    const normalized = normalizeRunNameInput(newRunName);
    if (!normalized) {
      setError('Bitte zuerst den Namen für das BTB eingeben.');
      return;
    }

    setBusy(true);
    setError('');
    try {
      const setupModel = await getSetupModel(templateId);
      if (!setupModel) throw new Error('Vorlage konnte nicht geladen werden.');

      const run = await createRun({
        templateId,
        title: buildRunTitleByConvention(normalized),
        setupVersion: setupModel.version,
      });

      const defaults = applyRunDefaultsFromModel(setupModel, run.values);
      if (defaults.changed) {
        await updateRun(run.runId, { values: defaults.values as Record<string, string | boolean> });
      }

      setNewRunName('');
      router.push(`/run/${run.runId}`);
    } catch (e) {
      setError((e as Error).message || 'BTB konnte nicht erstellt werden.');
    } finally {
      setBusy(false);
    }
  }

  function confirmDelete(run: Run) {
    Alert.alert('BTB löschen', `"${run.title}" wirklich löschen?`, [
      { text: 'Abbrechen', style: 'cancel' },
      {
        text: 'Löschen',
        style: 'destructive',
        onPress: async () => {
          await deleteRunCascade(run.runId);
          setInfo('BTB gelöscht.');
          await loadData();
        },
      },
    ]);
  }

  if (loading) {
    return (
      <View style={styles.centered}>
        <ActivityIndicator size="large" color={colors.primary} />
        <Text style={styles.muted}>Bautagebuch wird vorbereitet…</Text>
      </View>
    );
  }

  return (
    <ScrollView style={styles.container} contentContainerStyle={styles.content}>
      <View style={styles.hero}>
        <Text style={styles.heroTitle}>Bautagebuch</Text>
        <Text style={styles.heroSubtitle}>Offline BTB mit PDF-Vorlage, Formular und Fotodokumentation</Text>
      </View>

      {error ? <Text style={styles.error}>{error}</Text> : null}
      {info ? <Text style={styles.info}>{info}</Text> : null}

      <View style={styles.card}>
        <Text style={styles.cardTitle}>Neues Bautagebuch</Text>
        <TextInput
          style={styles.input}
          placeholder="Projekt / Baustelle"
          placeholderTextColor={colors.textMuted}
          value={newRunName}
          onChangeText={setNewRunName}
        />
        <Pressable style={[styles.primaryButton, busy && styles.disabled]} onPress={startNewRun} disabled={busy}>
          <Text style={styles.primaryButtonText}>{busy ? 'Wird erstellt…' : 'BTB starten'}</Text>
        </Pressable>
      </View>

      {weekGroups.map(([weekKey, weekRuns]) => (
        <View key={weekKey} style={styles.weekGroup}>
          <Text style={styles.weekLabel}>KW · {formatWeekLabel(weekKey)}</Text>
          {weekRuns.map((run) => (
            <Pressable key={run.runId} style={styles.runCard} onPress={() => router.push(`/run/${run.runId}`)}>
              <View style={styles.runHeader}>
                <Text style={styles.runTitle}>{run.title}</Text>
                <Pressable onPress={() => confirmDelete(run)} hitSlop={8}>
                  <Text style={styles.delete}>Löschen</Text>
                </Pressable>
              </View>
              <Text style={styles.runMeta}>
                Aktualisiert: {new Date(run.updatedAt).toLocaleString('de-DE')}
              </Text>
            </Pressable>
          ))}
        </View>
      ))}

      {runs.length === 0 ? <Text style={styles.muted}>Noch keine Bautagebücher. Starten Sie Ihr erstes BTB oben.</Text> : null}
    </ScrollView>
  );
}

const styles = StyleSheet.create({
  container: { flex: 1, backgroundColor: colors.background },
  content: { padding: 16, paddingBottom: 40 },
  centered: { flex: 1, alignItems: 'center', justifyContent: 'center', gap: 12, backgroundColor: colors.background },
  hero: { marginBottom: 16 },
  heroTitle: { color: colors.text, fontSize: 28, fontWeight: '800' },
  heroSubtitle: { color: colors.textMuted, fontSize: 15, lineHeight: 22, marginTop: 4 },
  card: { backgroundColor: colors.surface, borderRadius: 14, borderWidth: 1, borderColor: colors.border, padding: 16, marginBottom: 20 },
  cardTitle: { color: colors.text, fontSize: 16, fontWeight: '700', marginBottom: 10 },
  input: { backgroundColor: colors.background, borderColor: colors.border, borderRadius: 10, borderWidth: 1, color: colors.text, fontSize: 16, marginBottom: 12, minHeight: 44, paddingHorizontal: 12 },
  primaryButton: { backgroundColor: colors.primary, borderRadius: 10, paddingVertical: 14 },
  primaryButtonText: { color: '#fff', fontWeight: '700', textAlign: 'center' },
  disabled: { opacity: 0.6 },
  weekGroup: { marginBottom: 18 },
  weekLabel: { color: colors.accent, fontSize: 13, fontWeight: '700', marginBottom: 8 },
  runCard: { backgroundColor: colors.surface, borderRadius: 12, borderWidth: 1, borderColor: colors.border, marginBottom: 10, padding: 14 },
  runHeader: { flexDirection: 'row', justifyContent: 'space-between', gap: 12 },
  runTitle: { color: colors.text, flex: 1, fontSize: 16, fontWeight: '600' },
  delete: { color: colors.danger, fontSize: 13, fontWeight: '600' },
  runMeta: { color: colors.textMuted, fontSize: 12, marginTop: 6 },
  error: { backgroundColor: '#fdecec', borderRadius: 8, color: colors.danger, marginBottom: 12, padding: 10 },
  info: { backgroundColor: '#e8f4f2', borderRadius: 8, color: colors.primary, marginBottom: 12, padding: 10 },
  muted: { color: colors.textMuted, fontSize: 14, textAlign: 'center' },
});

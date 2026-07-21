import { useCallback, useEffect, useState } from 'react';
import { Alert, StyleSheet, Text, View } from 'react-native';
import { useLocalSearchParams, useRouter } from 'expo-router';

import { ListItem, PrimaryButton, Screen, TextField } from '../../src/components/mobile';
import { colors, typography } from '../../src/constants/theme';
import {
  addTemplate,
  defaultColumns,
  getTemplate,
  loadSettings,
  saveSettings,
  updateTemplate,
  type SiteReportColumn,
  type SiteReportTemplate
} from '../../src/native/sitereport/db/database';
import { nowIso } from '../../src/lib/ids';

function createColumnId() {
  return `col_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;
}

function withColumnIds(columns: SiteReportColumn[]): SiteReportColumn[] {
  return columns.map((col) => ({ ...col, id: col.id || createColumnId() }));
}

export default function SiteReportFormatBuilderScreen() {
  const router = useRouter();
  const { mode, templateId } = useLocalSearchParams<{ mode?: string; templateId?: string }>();
  const isEdit = mode === 'edit' && Boolean(templateId);

  const [templateName, setTemplateName] = useState('');
  const [columns, setColumns] = useState<SiteReportColumn[]>([]);
  const [newColName, setNewColName] = useState('');
  const [newColType, setNewColType] = useState<'text' | 'number'>('text');
  const [editingIndex, setEditingIndex] = useState<number | null>(null);
  const [editColName, setEditColName] = useState('');
  const [editColType, setEditColType] = useState<'text' | 'number'>('text');
  const [saving, setSaving] = useState(false);

  const load = useCallback(async () => {
    if (isEdit && templateId) {
      const tpl = await getTemplate(templateId);
      if (!tpl) {
        Alert.alert('Format', 'Vorlage nicht gefunden.');
        router.back();
        return;
      }
      setTemplateName(tpl.name);
      setColumns(withColumnIds(tpl.columns));
      return;
    }
    setTemplateName('');
    setColumns(withColumnIds([{ ...defaultColumns[0] }]));
  }, [isEdit, router, templateId]);

  useEffect(() => {
    void load();
  }, [load]);

  const addColumn = () => {
    if (!newColName.trim()) return;
    setColumns((prev) => [
      ...prev,
      { id: createColumnId(), name: newColName.trim(), type: newColType, isPhoto: false }
    ]);
    setNewColName('');
    setNewColType('text');
  };

  const removeColumn = (id: string) => {
    const target = columns.find((col) => col.id === id);
    if (!target || target.isPhoto) return;
    setColumns((prev) => prev.filter((col) => col.id !== id));
  };

  const moveColumn = (index: number, direction: -1 | 1) => {
    const nextIndex = index + direction;
    if (nextIndex < 0 || nextIndex >= columns.length) return;
    setColumns((prev) => {
      const next = prev.slice();
      const [moved] = next.splice(index, 1);
      next.splice(nextIndex, 0, moved);
      return next;
    });
  };

  const startEditColumn = (index: number) => {
    const col = columns[index];
    setEditingIndex(index);
    setEditColName(col.name);
    setEditColType(col.type);
  };

  const saveEditColumn = () => {
    if (editingIndex === null || !editColName.trim()) return;
    setColumns((prev) =>
      prev.map((col, idx) =>
        idx === editingIndex ? { ...col, name: editColName.trim(), type: editColType } : col
      )
    );
    setEditingIndex(null);
  };

  const save = async () => {
    if (!isEdit && !templateName.trim()) {
      Alert.alert('Format', 'Bitte gib einen Formatnamen ein.');
      return;
    }
    setSaving(true);
    try {
      const nextColumns = columns.map((col) => ({ ...col }));
      if (isEdit && templateId) {
        const existing = await getTemplate(templateId);
        if (!existing) throw new Error('Vorlage nicht gefunden.');
        const updated: SiteReportTemplate = { ...existing, columns: nextColumns };
        await updateTemplate(updated);
        const settings = await loadSettings();
        if (settings?.selectedTemplateId === templateId) {
          await saveSettings({ ...settings, columns: nextColumns });
        }
      } else {
        const record: SiteReportTemplate = {
          id: `tpl_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`,
          createdAt: nowIso(),
          name: templateName.trim(),
          columns: nextColumns
        };
        await addTemplate(record);
        await saveSettings({ selectedTemplateId: record.id, columns: nextColumns });
      }
      router.back();
    } catch (err) {
      Alert.alert('Format', err instanceof Error ? err.message : 'Speichern fehlgeschlagen.');
    } finally {
      setSaving(false);
    }
  };

  return (
    <Screen
      title={isEdit ? 'Format bearbeiten' : 'Neues Tabellenformat'}
      subtitle="Spalten für Excel- und PDF-Export definieren"
      showBack
      footer={<PrimaryButton label={saving ? 'Speichern…' : 'Format speichern'} disabled={saving} onPress={() => void save()} />}
    >
      {!isEdit ? (
        <TextField label="Formatname" value={templateName} onChangeText={setTemplateName} placeholder="z. B. Standard Baustelle" />
      ) : null}

      <Text style={styles.section}>Spalten</Text>
      {columns.map((col, index) => (
        <View key={col.id} style={styles.colCard}>
          {editingIndex === index ? (
            <>
              <TextField label="Spaltenname" value={editColName} onChangeText={setEditColName} />
              <ListItem
                title="Typ"
                subtitle={editColType === 'number' ? 'Zahl' : 'Text'}
                onPress={() => setEditColType((prev) => (prev === 'text' ? 'number' : 'text'))}
              />
              <View style={styles.row}>
                <PrimaryButton label="Übernehmen" onPress={saveEditColumn} />
                <PrimaryButton label="Abbrechen" variant="secondary" onPress={() => setEditingIndex(null)} />
              </View>
            </>
          ) : (
            <>
              <Text style={styles.colTitle}>{col.name}</Text>
              <Text style={styles.muted}>
                {col.isPhoto ? 'Foto-Spalte' : `Typ: ${col.type === 'number' ? 'Zahl' : 'Text'}`}
              </Text>
              <View style={styles.row}>
                {!col.isPhoto ? (
                  <>
                    <PrimaryButton label="Bearbeiten" variant="secondary" onPress={() => startEditColumn(index)} />
                    <PrimaryButton label="Entfernen" variant="ghost" onPress={() => removeColumn(col.id)} />
                  </>
                ) : null}
                <PrimaryButton label="↑" variant="ghost" onPress={() => moveColumn(index, -1)} />
                <PrimaryButton label="↓" variant="ghost" onPress={() => moveColumn(index, 1)} />
              </View>
            </>
          )}
        </View>
      ))}

      <Text style={styles.section}>Spalte hinzufügen</Text>
      <TextField label="Spaltenname" value={newColName} onChangeText={setNewColName} />
      <ListItem
        title="Typ"
        subtitle={newColType === 'number' ? 'Zahl' : 'Text'}
        onPress={() => setNewColType((prev) => (prev === 'text' ? 'number' : 'text'))}
      />
      <PrimaryButton label="Spalte hinzufügen" variant="secondary" onPress={addColumn} />
    </Screen>
  );
}

const styles = StyleSheet.create({
  section: { ...typography.bodyStrong, color: colors.ink, marginTop: 8 },
  colCard: { gap: 8, marginBottom: 12, padding: 12, borderRadius: 12, backgroundColor: colors.panel, borderWidth: 1, borderColor: colors.border },
  colTitle: { ...typography.bodyStrong, color: colors.ink },
  muted: { ...typography.body, color: colors.muted },
  row: { flexDirection: 'row', flexWrap: 'wrap', gap: 8 }
});

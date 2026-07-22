// @ts-nocheck
import { StyleSheet, Text, View } from 'react-native';

import { colors, spacing, typography } from '../../../constants/theme';
import { buildRunSections } from '../lib/setup-model.js';

function fieldKey(fieldId: string) {
  return `field:${fieldId}`;
}

function cellKey(cellId: string) {
  return `cell:${cellId}`;
}

function formatValue(value: unknown): string {
  if (value === true) return 'Ja';
  if (value === false) return 'Nein';
  if (value === null || value === undefined) return '—';
  const text = String(value).trim();
  return text || '—';
}

type Props = {
  setupModel: Record<string, unknown>;
  values: Record<string, unknown>;
  sectionIndex: number;
};

export function RunValuesPreview({ setupModel, values, sectionIndex }: Props) {
  const sections = buildRunSections(setupModel);
  const section = sections[sectionIndex];
  if (!section) return null;

  const rows: Array<{ key: string; label: string; value: string; multiline: boolean }> = [];

  if (section.kind === 'single') {
    for (const field of section.fields || []) {
      if (field.skipped) continue;
      rows.push({
        key: field.fieldId,
        label: field.label || field.fieldName,
        value: formatValue(values[fieldKey(field.fieldId)]),
        multiline: field.multiline === true
      });
    }
  }

  if (section.kind === 'table') {
    const visibleRowCount = Number(values[`__tableRows:${section.tableId || ''}`] ?? 1);
    const activeRows = (section.rows || []).slice(0, Math.max(1, visibleRowCount));
    for (const row of activeRows) {
      if (row.skipped) continue;
      for (const cell of row.cells || []) {
        if (cell.skipped) continue;
        const column = (section.columns || []).find((entry) => entry.columnId === cell.columnId);
        if (column?.skipped) continue;
        rows.push({
          key: cell.cellId,
          label: `${column?.label || cell.columnId} (Zeile ${row.index ?? '?'})`,
          value: formatValue(values[cellKey(cell.cellId)]),
          multiline: column?.multiline === true || cell.multiline === true
        });
      }
    }
  }

  const filled = rows.filter((row) => row.value !== '—');
  if (filled.length === 0) {
    return (
      <View style={styles.card}>
        <Text style={styles.title}>PDF-Vorschau (aktueller Abschnitt)</Text>
        <Text style={styles.muted}>Noch keine Werte für diesen Abschnitt eingetragen.</Text>
      </View>
    );
  }

  return (
    <View style={styles.card}>
      <Text style={styles.title}>PDF-Vorschau (aktueller Abschnitt)</Text>
      {filled.map((row) => (
        <View key={row.key} style={styles.row}>
          <Text style={styles.label}>{row.label}</Text>
          <Text style={[styles.value, row.multiline ? styles.valueMultiline : null]}>{row.value}</Text>
        </View>
      ))}
    </View>
  );
}

const styles = StyleSheet.create({
  card: {
    gap: spacing.xs,
    padding: spacing.md,
    borderRadius: 12,
    backgroundColor: colors.panel,
    borderWidth: 1,
    borderColor: colors.border
  },
  title: { ...typography.label, color: colors.muted },
  muted: { ...typography.caption, color: colors.muted },
  row: { gap: 2, paddingVertical: 4, borderTopWidth: 1, borderTopColor: colors.border },
  label: { ...typography.caption, color: colors.muted },
  value: { ...typography.body, color: colors.ink },
  valueMultiline: {
    lineHeight: 22
  }
});

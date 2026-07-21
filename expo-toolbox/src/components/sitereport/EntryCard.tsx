import { Image, Pressable, StyleSheet, Text, View } from 'react-native';

import { colors, spacing, typography } from '../../constants/theme';
import type { SiteReportColumn, SiteReportEntry } from '../../native/sitereport/db/database';
import { Card, PrimaryButton, StatusBadge } from '../mobile';

type Props = {
  entry: SiteReportEntry;
  columns: SiteReportColumn[];
  onEdit?: () => void;
  onDelete?: () => void;
};

function statusTone(value: string): 'warning' | 'info' | 'success' | 'neutral' {
  const lower = value.toLowerCase();
  if (lower === 'offen') return 'warning';
  if (lower === 'bearbeitung') return 'info';
  if (lower === 'erledigt') return 'success';
  return 'neutral';
}

export function EntryCard({ entry, columns, onEdit, onDelete }: Props) {
  const dataColumns = columns.filter((col) => !col.isPhoto);
  const statusCol = dataColumns.find((col) => col.name.toLowerCase() === 'status');
  const statusValue = statusCol ? String(entry.fields[statusCol.name] ?? '') : '';
  const otherColumns = dataColumns.filter((col) => col.name.toLowerCase() !== 'status');

  return (
    <Card style={styles.card}>
      {entry.photoPath ? (
        <Image source={{ uri: entry.photoPath }} style={styles.photo} resizeMode="cover" />
      ) : (
        <View style={styles.photoPlaceholder}>
          <Text style={styles.placeholderText}>Kein Foto</Text>
        </View>
      )}
      <View style={styles.body}>
        {statusValue ? <StatusBadge label={statusValue} tone={statusTone(statusValue)} /> : null}
        {otherColumns.map((col) => {
          const value = entry.fields[col.name];
          if (value === undefined || value === '') return null;
          return (
            <View key={col.id} style={styles.field}>
              <Text style={styles.fieldLabel}>{col.name}</Text>
              <Text style={styles.fieldValue}>{String(value)}</Text>
            </View>
          );
        })}
        <View style={styles.actions}>
          {onEdit ? <PrimaryButton label="Bearbeiten" variant="secondary" onPress={onEdit} /> : null}
          {onDelete ? <PrimaryButton label="Löschen" variant="ghost" onPress={onDelete} /> : null}
        </View>
      </View>
    </Card>
  );
}

const styles = StyleSheet.create({
  card: {
    marginBottom: spacing.md,
    padding: 0,
    overflow: 'hidden'
  },
  photo: {
    width: '100%',
    height: 200,
    backgroundColor: colors.border
  },
  photoPlaceholder: {
    height: 120,
    backgroundColor: colors.border,
    alignItems: 'center',
    justifyContent: 'center'
  },
  placeholderText: {
    ...typography.caption,
    color: colors.muted
  },
  body: {
    padding: spacing.cardPadding,
    gap: spacing.sm
  },
  field: {
    gap: 2
  },
  fieldLabel: {
    ...typography.label,
    color: colors.muted
  },
  fieldValue: {
    ...typography.body,
    color: colors.ink
  },
  actions: {
    flexDirection: 'row',
    flexWrap: 'wrap',
    gap: spacing.xs,
    marginTop: spacing.xs
  }
});

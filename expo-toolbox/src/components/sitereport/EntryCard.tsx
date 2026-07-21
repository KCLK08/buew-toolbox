import { Image, Pressable, StyleSheet, Text, View } from 'react-native';

import { colors, shadows, spacing, typography } from '../../constants/theme';
import type { SiteReportColumn, SiteReportEntry } from '../../native/sitereport/db/database';
import { Card, StatusBadge } from '../mobile';
import { hapticSelection } from '../../lib/haptics';

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
  const otherColumns = dataColumns.filter((col) => col.name.toLowerCase() !== 'status').slice(0, 3);

  return (
    <Card style={styles.card} padded={false}>
      {entry.photoPath ? (
        <Image source={{ uri: entry.photoPath }} style={styles.photo} resizeMode="cover" />
      ) : (
        <View style={styles.photoPlaceholder}>
          <Text style={styles.placeholderIcon}>📷</Text>
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
              <Text style={styles.fieldValue} numberOfLines={2}>
                {String(value)}
              </Text>
            </View>
          );
        })}
        <View style={styles.actions}>
          {onEdit ? (
            <Pressable
              style={({ pressed }) => [styles.actionBtn, pressed ? styles.actionPressed : null]}
              onPress={() => {
                void hapticSelection();
                onEdit();
              }}
            >
              <Text style={styles.actionLabel}>Bearbeiten</Text>
            </Pressable>
          ) : null}
          {onDelete ? (
            <Pressable
              style={({ pressed }) => [styles.actionBtn, styles.actionDanger, pressed ? styles.actionPressed : null]}
              onPress={() => {
                void hapticSelection();
                onDelete();
              }}
            >
              <Text style={[styles.actionLabel, styles.actionDangerLabel]}>Löschen</Text>
            </Pressable>
          ) : null}
        </View>
      </View>
    </Card>
  );
}

const styles = StyleSheet.create({
  card: {
    marginBottom: spacing.md,
    overflow: 'hidden',
    ...shadows.card
  },
  photo: {
    width: '100%',
    height: 220,
    backgroundColor: colors.border
  },
  photoPlaceholder: {
    height: 160,
    backgroundColor: colors.bg,
    alignItems: 'center',
    justifyContent: 'center',
    gap: spacing.xs
  },
  placeholderIcon: {
    fontSize: 32,
    opacity: 0.5
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
    ...typography.bodyStrong,
    color: colors.ink
  },
  actions: {
    flexDirection: 'row',
    gap: spacing.sm,
    marginTop: spacing.xs
  },
  actionBtn: {
    flex: 1,
    minHeight: spacing.touchMin,
    borderRadius: spacing.buttonRadius,
    backgroundColor: colors.bg,
    alignItems: 'center',
    justifyContent: 'center',
    borderWidth: 1,
    borderColor: colors.border
  },
  actionDanger: {
    backgroundColor: 'transparent',
    borderColor: 'transparent'
  },
  actionPressed: {
    opacity: 0.85
  },
  actionLabel: {
    ...typography.label,
    color: colors.ink
  },
  actionDangerLabel: {
    color: colors.danger
  }
});

import { StyleSheet, Text, View } from 'react-native';

import { colors, spacing, typography } from '../../constants/theme';
import type { SiteReportColumn } from '../../native/sitereport/db/database';
import { Card, PrimaryButton, TextField } from '../mobile';

type Props = {
  column: SiteReportColumn;
  index: number;
  total: number;
  editing: boolean;
  editName: string;
  editType: 'text' | 'number';
  onEditNameChange: (value: string) => void;
  onEditTypeToggle: () => void;
  onStartEdit: () => void;
  onSaveEdit: () => void;
  onCancelEdit: () => void;
  onRemove: () => void;
  onMoveUp: () => void;
  onMoveDown: () => void;
};

export function FormatColumnCard({
  column,
  index,
  total,
  editing,
  editName,
  editType,
  onEditNameChange,
  onEditTypeToggle,
  onStartEdit,
  onSaveEdit,
  onCancelEdit,
  onRemove,
  onMoveUp,
  onMoveDown
}: Props) {
  return (
    <Card style={styles.card}>
      {editing ? (
        <View style={styles.gap}>
          <TextField label="Spaltenname" value={editName} onChangeText={onEditNameChange} />
          {!column.isPhoto ? (
            <PrimaryButton
              label={`Typ: ${editType === 'number' ? 'Zahl' : 'Text'}`}
              variant="secondary"
              onPress={onEditTypeToggle}
            />
          ) : null}
          <View style={styles.actions}>
            <PrimaryButton label="Speichern" onPress={onSaveEdit} />
            <PrimaryButton label="Abbrechen" variant="ghost" onPress={onCancelEdit} />
          </View>
        </View>
      ) : (
        <View style={styles.row}>
          {!column.isPhoto ? (
            <View style={styles.handle}>
              <Text style={styles.handleIcon}>≡</Text>
            </View>
          ) : (
            <View style={styles.photoHandle}>
              <Text style={styles.photoHandleIcon}>📷</Text>
            </View>
          )}
          <View style={styles.content}>
            <View style={styles.header}>
              <Text style={styles.position}>#{index + 1}</Text>
              <Text style={styles.name}>{column.name}</Text>
            </View>
            <Text style={styles.meta}>
              {column.isPhoto ? 'Foto-Spalte (fixiert)' : `Typ: ${column.type === 'number' ? 'Zahl' : 'Text'}`}
            </Text>
            {!column.isPhoto ? (
              <View style={styles.actions}>
                <PrimaryButton label="Bearbeiten" variant="secondary" onPress={onStartEdit} />
                <PrimaryButton label="Entfernen" variant="ghost" onPress={onRemove} />
              </View>
            ) : (
              <Text style={styles.photoPill}>Foto</Text>
            )}
          </View>
          <View style={styles.reorder}>
            <PrimaryButton label="↑" variant="ghost" onPress={onMoveUp} disabled={index === 0} />
            <PrimaryButton label="↓" variant="ghost" onPress={onMoveDown} disabled={index >= total - 1} />
          </View>
        </View>
      )}
    </Card>
  );
}

const styles = StyleSheet.create({
  card: {
    marginBottom: spacing.sm
  },
  gap: {
    gap: spacing.sm
  },
  row: {
    flexDirection: 'row',
    alignItems: 'flex-start',
    gap: spacing.sm
  },
  handle: {
    width: 32,
    paddingTop: 4,
    alignItems: 'center'
  },
  handleIcon: {
    fontSize: 22,
    color: colors.muted,
    lineHeight: 26
  },
  photoHandle: {
    width: 32,
    paddingTop: 4,
    alignItems: 'center'
  },
  photoHandleIcon: {
    fontSize: 18
  },
  content: {
    flex: 1,
    gap: 4
  },
  header: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: spacing.xs
  },
  position: {
    ...typography.caption,
    color: colors.muted,
    backgroundColor: colors.bg,
    paddingHorizontal: spacing.xs,
    paddingVertical: 2,
    borderRadius: 6,
    overflow: 'hidden'
  },
  name: {
    ...typography.bodyStrong,
    color: colors.ink,
    flex: 1
  },
  meta: {
    ...typography.caption,
    color: colors.muted
  },
  photoPill: {
    ...typography.label,
    color: colors.accent,
    backgroundColor: colors.badgeBg,
    paddingHorizontal: spacing.sm,
    paddingVertical: 4,
    borderRadius: 999,
    overflow: 'hidden',
    alignSelf: 'flex-start',
    marginTop: spacing.xs
  },
  reorder: {
    gap: 0
  },
  actions: {
    flexDirection: 'row',
    flexWrap: 'wrap',
    gap: spacing.xs,
    marginTop: spacing.xs
  }
});

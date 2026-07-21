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
        <>
          <Text style={styles.name}>{column.name}</Text>
          <Text style={styles.meta}>
            {column.isPhoto ? 'Foto-Spalte (fest)' : `Typ: ${column.type === 'number' ? 'Zahl' : 'Text'}`}
          </Text>
          <View style={styles.actions}>
            {!column.isPhoto ? (
              <>
                <PrimaryButton label="Bearbeiten" variant="secondary" onPress={onStartEdit} />
                <PrimaryButton label="Entfernen" variant="ghost" onPress={onRemove} />
              </>
            ) : (
              <Text style={styles.photoPill}>Foto</Text>
            )}
            <PrimaryButton label="↑" variant="ghost" onPress={onMoveUp} disabled={index === 0} />
            <PrimaryButton label="↓" variant="ghost" onPress={onMoveDown} disabled={index >= total - 1} />
          </View>
        </>
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
  name: {
    ...typography.bodyStrong,
    color: colors.ink
  },
  meta: {
    ...typography.caption,
    color: colors.muted,
    marginTop: 2,
    marginBottom: spacing.sm
  },
  photoPill: {
    ...typography.label,
    color: colors.accent,
    backgroundColor: colors.badgeBg,
    paddingHorizontal: spacing.sm,
    paddingVertical: 4,
    borderRadius: 999,
    overflow: 'hidden',
    alignSelf: 'flex-start'
  },
  actions: {
    flexDirection: 'row',
    flexWrap: 'wrap',
    gap: spacing.xs
  }
});

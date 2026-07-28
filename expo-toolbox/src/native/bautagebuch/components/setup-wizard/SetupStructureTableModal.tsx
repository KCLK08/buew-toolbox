import { useEffect, useState } from 'react';
import { Alert, Modal, Pressable, ScrollView, StyleSheet, Text, View } from 'react-native';
import { MaterialCommunityIcons } from '@expo/vector-icons';
import { useSafeAreaInsets } from 'react-native-safe-area-context';

import { PrimaryButton, TextField } from '../../../../components/mobile';
import { colors, spacing, typography } from '../../../../constants/theme';
import { systemBottomInset } from '../../../../navigation/systemInsets';

type ColumnDraft = {
  id?: string;
  name: string;
};

type Props = {
  visible: boolean;
  initialName?: string;
  initialColumns?: ColumnDraft[];
  readOnly?: boolean;
  onClose: () => void;
  onSave: (input: { name: string; columns: ColumnDraft[] }) => void;
};

export function SetupStructureTableModal({
  visible,
  initialName = '',
  initialColumns = [],
  readOnly = false,
  onClose,
  onSave
}: Props) {
  const insets = useSafeAreaInsets();
  const [name, setName] = useState(initialName);
  const [columns, setColumns] = useState<ColumnDraft[]>(initialColumns);

  useEffect(() => {
    if (!visible) return;
    setName(initialName);
    setColumns(initialColumns.length > 0 ? initialColumns : [{ name: '' }]);
  }, [visible, initialName, initialColumns]);

  const addColumn = () => {
    setColumns((current) => [...current, { name: '' }]);
  };

  const updateColumn = (index: number, value: string) => {
    setColumns((current) =>
      current.map((column, columnIndex) =>
        columnIndex === index ? { ...column, name: value } : column
      )
    );
  };

  const removeColumn = (index: number) => {
    setColumns((current) => current.filter((_, columnIndex) => columnIndex !== index));
  };

  const submit = () => {
    const trimmedColumns = columns
      .map((column) => ({ ...column, name: column.name.trim() }))
      .filter((column) => column.name.length > 0);
    if (!name.trim()) {
      Alert.alert('Tabelle', 'Bitte einen Namen eingeben.');
      return;
    }
    if (trimmedColumns.length === 0) {
      Alert.alert('Tabelle', 'Lege mindestens eine Spalte an.');
      return;
    }
    onSave({ name: name.trim(), columns: trimmedColumns });
  };

  return (
    <Modal visible={visible} animationType="slide" onRequestClose={onClose}>
      <View style={[styles.root, { paddingTop: insets.top }]}>
        <View style={styles.header}>
          <Pressable accessibilityRole="button" onPress={onClose}>
            <Text style={styles.cancel}>Abbrechen</Text>
          </Pressable>
          <Text style={styles.title}>{initialName ? 'Tabelle bearbeiten' : 'Tabelle hinzufügen'}</Text>
          <View style={styles.headerSpacer} />
        </View>

        <ScrollView
          style={styles.body}
          contentContainerStyle={[
            styles.bodyContent,
            { paddingBottom: systemBottomInset(insets) + spacing.xl }
          ]}
          keyboardShouldPersistTaps="handled"
        >
          <TextField
            label="Name"
            value={name}
            onChangeText={setName}
            placeholder="z. B. Arbeitsleistungen"
            editable={!readOnly}
            autoCapitalize="sentences"
          />

          <View style={styles.columnsSection}>
            <Text style={styles.columnsTitle}>Spalten definieren</Text>
            {columns.map((column, index) => (
              <View key={column.id || `col_${index}`} style={styles.columnRow}>
                <TextField
                  label={`Spalte ${index + 1}`}
                  value={column.name}
                  onChangeText={(value) => updateColumn(index, value)}
                  placeholder="Spaltenname"
                  editable={!readOnly}
                />
                {!readOnly && columns.length > 1 ? (
                  <Pressable
                    accessibilityRole="button"
                    accessibilityLabel="Spalte entfernen"
                    style={styles.removeBtn}
                    onPress={() => removeColumn(index)}
                  >
                    <MaterialCommunityIcons name="close-circle" size={22} color={colors.muted} />
                  </Pressable>
                ) : null}
              </View>
            ))}
            {!readOnly ? (
              <Pressable accessibilityRole="button" style={styles.addColumnBtn} onPress={addColumn}>
                <MaterialCommunityIcons name="plus" size={18} color={colors.accent} />
                <Text style={styles.addColumnLabel}>Spalte hinzufügen</Text>
              </Pressable>
            ) : null}
          </View>
        </ScrollView>

        {!readOnly ? (
          <View style={[styles.footer, { paddingBottom: systemBottomInset(insets) + spacing.sm }]}>
            <PrimaryButton label="Speichern" onPress={submit} />
          </View>
        ) : null}
      </View>
    </Modal>
  );
}

const styles = StyleSheet.create({
  root: {
    flex: 1,
    backgroundColor: colors.bg
  },
  header: {
    flexDirection: 'row',
    alignItems: 'center',
    justifyContent: 'space-between',
    paddingHorizontal: spacing.pageX,
    paddingVertical: spacing.sm,
    borderBottomWidth: 1,
    borderBottomColor: colors.border,
    backgroundColor: colors.panel
  },
  cancel: {
    ...typography.bodyStrong,
    color: colors.accent,
    minWidth: 88
  },
  title: {
    ...typography.subtitle,
    color: colors.ink
  },
  headerSpacer: {
    minWidth: 88
  },
  body: {
    flex: 1
  },
  bodyContent: {
    padding: spacing.pageX,
    gap: spacing.md
  },
  columnsSection: {
    gap: spacing.sm
  },
  columnsTitle: {
    ...typography.label,
    color: colors.muted
  },
  columnRow: {
    gap: spacing.xxs
  },
  removeBtn: {
    alignSelf: 'flex-end',
    padding: spacing.xxs
  },
  addColumnBtn: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: spacing.xs,
    minHeight: spacing.touchMin,
    paddingHorizontal: spacing.sm,
    borderRadius: 12,
    borderWidth: 1,
    borderColor: colors.border,
    backgroundColor: colors.panelElevated
  },
  addColumnLabel: {
    ...typography.bodyStrong,
    color: colors.accent
  },
  footer: {
    paddingHorizontal: spacing.pageX,
    paddingTop: spacing.sm,
    borderTopWidth: 1,
    borderTopColor: colors.border,
    backgroundColor: colors.panel
  }
});

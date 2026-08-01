import { useEffect, useRef, useState } from 'react';
import { Alert, Modal, Pressable, StyleSheet, Text, View } from 'react-native';
import { MaterialCommunityIcons } from '@expo/vector-icons';
import { useSafeAreaInsets } from 'react-native-safe-area-context';

import { PrimaryButton, TextField } from '../../../../components/mobile';
import { colors, spacing, typography } from '../../../../constants/theme';
import { hapticSelection } from '../../../../lib/haptics';
import { systemBottomInset } from '../../../../navigation/systemInsets';
import { SetupModalKeyboardFrame } from './SetupModalKeyboardFrame';
import { SetupScrollView } from './SetupScrollView';

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

function normalizeColumnCount(value: string): number | null {
  const trimmed = value.trim();
  if (!trimmed) return null;
  const parsed = Number(trimmed);
  if (!Number.isFinite(parsed) || parsed < 0) return null;
  return Math.min(20, Math.floor(parsed));
}

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
  const [columnCountInput, setColumnCountInput] = useState('');
  const [columns, setColumns] = useState<ColumnDraft[]>(initialColumns);
  const wasVisibleRef = useRef(false);
  const isEditing = Boolean(initialName);

  useEffect(() => {
    if (visible && !wasVisibleRef.current) {
      setName(initialName);
      const nextColumns = initialColumns.map((column) => ({ ...column }));
      setColumns(nextColumns);
      setColumnCountInput(nextColumns.length > 0 ? String(nextColumns.length) : '');
    }
    wasVisibleRef.current = visible;
  }, [visible, initialName, initialColumns]);

  const applyColumnCount = (count: number | null) => {
    if (count === null) return;
    setColumns((current) => {
      if (count === current.length) return current;
      if (count < current.length) {
        return current.slice(0, count);
      }
      const next = [...current];
      while (next.length < count) {
        next.push({ name: '' });
      }
      return next;
    });
  };

  const handleColumnCountChange = (value: string) => {
    setColumnCountInput(value);
    const count = normalizeColumnCount(value);
    if (count !== null) {
      applyColumnCount(count);
    }
  };

  const addColumn = () => {
    void hapticSelection();
    setColumns((current) => {
      const next = [...current, { name: '' }];
      setColumnCountInput(String(next.length));
      return next;
    });
  };

  const updateColumn = (index: number, value: string) => {
    setColumns((current) =>
      current.map((column, columnIndex) =>
        columnIndex === index ? { ...column, name: value } : column
      )
    );
  };

  const removeColumn = (index: number) => {
    void hapticSelection();
    setColumns((current) => {
      const next = current.filter((_, columnIndex) => columnIndex !== index);
      setColumnCountInput(next.length > 0 ? String(next.length) : '');
      return next;
    });
  };

  const submit = () => {
    const trimmedColumns = columns
      .map((column) => ({ ...column, name: column.name.trim() }))
      .filter((column) => column.name.length > 0);
    if (!name.trim()) {
      Alert.alert('Tabelle', 'Bitte einen Namen eingeben.');
      return;
    }
    onSave({ name: name.trim(), columns: trimmedColumns });
  };

  return (
    <Modal visible={visible} animationType="slide" onRequestClose={onClose}>
      <SetupModalKeyboardFrame>
        <View style={[styles.root, { paddingTop: insets.top }]}>
        <View style={styles.header}>
          <Pressable accessibilityRole="button" style={styles.headerBtn} onPress={onClose}>
            <Text style={styles.cancel}>Abbrechen</Text>
          </Pressable>
          <Text style={styles.title}>{isEditing ? 'Tabelle bearbeiten' : 'Tabelle hinzufügen'}</Text>
          <View style={styles.headerBtn} />
        </View>

        <SetupScrollView
          style={styles.body}
          contentContainerStyle={[
            styles.bodyContent,
            { paddingBottom: systemBottomInset(insets) + spacing.xl }
          ]}
        >
          <View style={styles.hero}>
            <View style={styles.heroIconWrap}>
              <MaterialCommunityIcons name="table" size={28} color={colors.info} />
            </View>
            <Text style={styles.heroTitle}>Tabellenbereich</Text>
            <Text style={styles.heroCopy}>
              Definiere zuerst die Tabellenstruktur. PDF-Felder werden in Schritt 2 den Spalten
              zugeordnet.
            </Text>
          </View>

          <TextField
            label="Tabellenname"
            value={name}
            onChangeText={setName}
            placeholder="z. B. Arbeitskräfte"
            editable={!readOnly}
            autoCapitalize="sentences"
          />

          <TextField
            label="Anzahl der Spalten (optional)"
            value={columnCountInput}
            onChangeText={handleColumnCountChange}
            placeholder="z. B. 4"
            editable={!readOnly}
            keyboardType="number-pad"
            hint="Leer lassen, wenn die Spalten später einzeln hinzugefügt werden."
          />

          {columns.length > 0 ? (
            <View style={styles.columnsSection}>
              <View style={styles.columnsHeader}>
                <Text style={styles.columnsTitle}>Spalten benennen</Text>
                <Text style={styles.columnsMeta}>{columns.filter((c) => c.name.trim()).length} benannt</Text>
              </View>
              {columns.map((column, index) => (
                <View key={column.id || `col_${index}`} style={styles.columnRow}>
                  <View style={styles.columnIndex}>
                    <Text style={styles.columnIndexText}>{index + 1}</Text>
                  </View>
                  <View style={styles.columnField}>
                    <TextField
                      label={`Spalte ${index + 1}`}
                      value={column.name}
                      onChangeText={(value) => updateColumn(index, value)}
                      placeholder="Spaltenname"
                      editable={!readOnly}
                    />
                  </View>
                  {!readOnly ? (
                    <Pressable
                      accessibilityRole="button"
                      accessibilityLabel="Spalte entfernen"
                      style={styles.removeBtn}
                      onPress={() => removeColumn(index)}
                    >
                      <MaterialCommunityIcons name="minus-circle-outline" size={24} color={colors.muted} />
                    </Pressable>
                  ) : null}
                </View>
              ))}
            </View>
          ) : null}

          {!readOnly ? (
            <Pressable accessibilityRole="button" style={styles.addColumnBtn} onPress={addColumn}>
              <MaterialCommunityIcons name="plus-circle-outline" size={22} color={colors.accent} />
              <Text style={styles.addColumnLabel}>Spalte hinzufügen</Text>
            </Pressable>
          ) : null}
        </SetupScrollView>

        {!readOnly ? (
          <View style={[styles.footer, { paddingBottom: systemBottomInset(insets) + spacing.sm }]}>
            <PrimaryButton label="Speichern" onPress={submit} />
          </View>
        ) : null}
        </View>
      </SetupModalKeyboardFrame>
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
  headerBtn: {
    minWidth: 88
  },
  cancel: {
    ...typography.bodyStrong,
    color: colors.accent
  },
  title: {
    ...typography.subtitle,
    color: colors.ink
  },
  body: {
    flex: 1
  },
  bodyContent: {
    padding: spacing.pageX,
    gap: spacing.md
  },
  hero: {
    alignItems: 'center',
    gap: spacing.xs,
    paddingVertical: spacing.sm
  },
  heroIconWrap: {
    width: 56,
    height: 56,
    borderRadius: 16,
    alignItems: 'center',
    justifyContent: 'center',
    backgroundColor: 'rgba(42, 95, 143, 0.12)'
  },
  heroTitle: {
    ...typography.subtitle,
    color: colors.ink
  },
  heroCopy: {
    ...typography.body,
    color: colors.muted,
    textAlign: 'center',
    lineHeight: 22
  },
  columnsSection: {
    gap: spacing.sm
  },
  columnsHeader: {
    flexDirection: 'row',
    alignItems: 'center',
    justifyContent: 'space-between'
  },
  columnsTitle: {
    ...typography.label,
    color: colors.muted
  },
  columnsMeta: {
    ...typography.caption,
    color: colors.muted
  },
  columnRow: {
    flexDirection: 'row',
    alignItems: 'flex-start',
    gap: spacing.sm
  },
  columnIndex: {
    width: 28,
    height: 28,
    marginTop: 28,
    borderRadius: 999,
    alignItems: 'center',
    justifyContent: 'center',
    backgroundColor: colors.panel,
    borderWidth: 1,
    borderColor: colors.border
  },
  columnIndexText: {
    ...typography.caption,
    color: colors.muted,
    fontFamily: 'SpaceGrotesk_600SemiBold'
  },
  columnField: {
    flex: 1
  },
  removeBtn: {
    marginTop: 30,
    padding: spacing.xxs
  },
  addColumnBtn: {
    flexDirection: 'row',
    alignItems: 'center',
    justifyContent: 'center',
    gap: spacing.xs,
    minHeight: spacing.touchMin + 4,
    paddingHorizontal: spacing.md,
    borderRadius: spacing.cardRadius,
    borderWidth: 1,
    borderColor: colors.border,
    borderStyle: 'dashed',
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

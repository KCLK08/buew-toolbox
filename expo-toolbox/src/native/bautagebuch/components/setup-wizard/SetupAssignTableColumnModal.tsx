import { useEffect, useRef, useState } from 'react';
import { Modal, Pressable, ScrollView, StyleSheet, Text, View } from 'react-native';
import { MaterialCommunityIcons } from '@expo/vector-icons';
import { useSafeAreaInsets } from 'react-native-safe-area-context';

import { PrimaryButton, SingleLineText } from '../../../../components/mobile';
import { colors, shadows, spacing, typography } from '../../../../constants/theme';
import { hapticSelection } from '../../../../lib/haptics';
import { systemBottomInset } from '../../../../navigation/systemInsets';
import type { SetupStructureTableColumn } from '../../types';

type Props = {
  visible: boolean;
  tableName: string;
  columns: SetupStructureTableColumn[];
  suggestedFieldName: string;
  readOnly?: boolean;
  onClose: () => void;
  onConfirm: (input: { columnId?: string; newColumnName?: string; fieldLabel: string }) => void;
};

export function SetupAssignTableColumnModal({
  visible,
  tableName,
  columns,
  suggestedFieldName,
  readOnly = false,
  onClose,
  onConfirm
}: Props) {
  const insets = useSafeAreaInsets();
  const [mode, setMode] = useState<'pick' | 'new'>('pick');
  const [selectedColumnId, setSelectedColumnId] = useState<string | null>(null);
  const wasVisibleRef = useRef(false);

  useEffect(() => {
    if (visible && !wasVisibleRef.current) {
      setMode(columns.length > 0 ? 'pick' : 'new');
      setSelectedColumnId(columns[0]?.id || null);
    }
    wasVisibleRef.current = visible;
  }, [visible, columns]);

  const submit = () => {
    if (mode === 'new') {
      onConfirm({
        newColumnName: suggestedFieldName,
        fieldLabel: suggestedFieldName
      });
      return;
    }
    if (!selectedColumnId) return;
    onConfirm({
      columnId: selectedColumnId,
      fieldLabel: suggestedFieldName
    });
  };

  return (
    <Modal visible={visible} animationType="slide" onRequestClose={onClose}>
      <View style={[styles.root, { paddingTop: insets.top }]}>
        <View style={styles.header}>
          <Pressable accessibilityRole="button" style={styles.headerBtn} onPress={onClose}>
            <Text style={styles.cancel}>Abbrechen</Text>
          </Pressable>
          <Text style={styles.title}>Spalte wählen</Text>
          <View style={styles.headerBtn} />
        </View>

        <ScrollView
          style={styles.body}
          contentContainerStyle={[
            styles.bodyContent,
            { paddingBottom: systemBottomInset(insets) + spacing.xl }
          ]}
          keyboardShouldPersistTaps="handled"
        >
          <View style={styles.hero}>
            <View style={styles.heroIconWrap}>
              <MaterialCommunityIcons name="table" size={28} color={colors.info} />
            </View>
            <Text style={styles.heroTitle}>Tabelle</Text>
            <Text style={styles.tableName}>{tableName}</Text>
            <Text style={styles.fieldHint}>
              Feld „{suggestedFieldName}" einer Spalte zuordnen
            </Text>
          </View>

          <Pressable
            accessibilityRole="button"
            style={[styles.option, mode === 'new' ? styles.optionActive : null, shadows.card]}
            onPress={() => {
              void hapticSelection();
              setMode('new');
            }}
          >
            <MaterialCommunityIcons
              name="table-column"
              size={22}
              color={mode === 'new' ? colors.accent2 : colors.muted}
            />
            <View style={styles.optionCopy}>
              <Text style={mode === 'new' ? styles.optionTitleActive : styles.optionTitle}>
                Neue Spalte erzeugen
              </Text>
              <Text style={styles.optionHint}>Spalte „{suggestedFieldName}" anlegen</Text>
            </View>
          </Pressable>

          {columns.length > 0 ? (
            <View style={styles.columnList}>
              <Text style={styles.sectionLabel}>Vorhandene Spalten</Text>
              {columns.map((column) => {
                const active = mode === 'pick' && selectedColumnId === column.id;
                return (
                  <Pressable
                    key={column.id}
                    accessibilityRole="button"
                    style={[styles.columnRow, active ? styles.columnRowActive : null]}
                    onPress={() => {
                      void hapticSelection();
                      setMode('pick');
                      setSelectedColumnId(column.id);
                    }}
                  >
                    <MaterialCommunityIcons
                      name={active ? 'radiobox-marked' : 'radiobox-blank'}
                      size={20}
                      color={active ? colors.accent : colors.muted}
                    />
                    <SingleLineText style={active ? styles.columnLabelActive : styles.columnLabel}>
                      {column.name}
                    </SingleLineText>
                  </Pressable>
                );
              })}
            </View>
          ) : null}
        </ScrollView>

        {!readOnly ? (
          <View style={[styles.footer, { paddingBottom: systemBottomInset(insets) + spacing.sm }]}>
            <PrimaryButton
              label="Übernehmen"
              onPress={submit}
              disabled={mode === 'pick' && !selectedColumnId}
            />
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
    gap: spacing.xxs,
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
    ...typography.caption,
    color: colors.muted
  },
  tableName: {
    ...typography.subtitle,
    color: colors.ink,
    textAlign: 'center'
  },
  fieldHint: {
    ...typography.body,
    color: colors.muted,
    textAlign: 'center',
    lineHeight: 22
  },
  option: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: spacing.sm,
    padding: spacing.md,
    borderRadius: 14,
    borderWidth: 1,
    borderColor: colors.border,
    backgroundColor: colors.panel
  },
  optionActive: {
    borderColor: colors.accent,
    backgroundColor: colors.badgeBg
  },
  optionCopy: {
    flex: 1,
    gap: 2
  },
  optionTitle: {
    ...typography.bodyStrong,
    color: colors.ink
  },
  optionTitleActive: {
    ...typography.bodyStrong,
    color: colors.accent2
  },
  optionHint: {
    ...typography.caption,
    color: colors.muted
  },
  sectionLabel: {
    ...typography.label,
    color: colors.muted,
    marginBottom: spacing.xxs
  },
  columnList: {
    gap: spacing.xxs
  },
  columnRow: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: spacing.sm,
    minHeight: spacing.touchMin,
    paddingHorizontal: spacing.sm,
    borderRadius: 12,
    borderWidth: 1,
    borderColor: colors.border,
    backgroundColor: colors.panel
  },
  columnRowActive: {
    borderColor: colors.accent,
    backgroundColor: colors.badgeBg
  },
  columnLabel: {
    ...typography.body,
    color: colors.ink,
    flex: 1
  },
  columnLabelActive: {
    ...typography.bodyStrong,
    color: colors.accent2,
    flex: 1
  },
  footer: {
    paddingHorizontal: spacing.pageX,
    paddingTop: spacing.sm,
    borderTopWidth: 1,
    borderTopColor: colors.border,
    backgroundColor: colors.panel
  }
});

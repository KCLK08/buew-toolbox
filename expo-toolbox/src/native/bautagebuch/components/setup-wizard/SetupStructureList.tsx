import { Alert, Pressable, ScrollView, StyleSheet, Text, View } from 'react-native';
import { MaterialCommunityIcons } from '@expo/vector-icons';

import { PrimaryButton } from '../../../../components/mobile';
import { colors, shadows, spacing, typography } from '../../../../constants/theme';
import type { SetupStructureItem } from '../../types';

type Props = {
  items: SetupStructureItem[];
  readOnly?: boolean;
  onAddGroup: () => void;
  onAddTable: () => void;
  onEdit: (item: SetupStructureItem) => void;
  onDelete: (item: SetupStructureItem) => void;
  onMove: (id: string, direction: -1 | 1) => void;
};

function itemIcon(item: SetupStructureItem): keyof typeof MaterialCommunityIcons.glyphMap {
  return item.type === 'group' ? 'file-document-outline' : 'table';
}

export function SetupStructureList({
  items,
  readOnly = false,
  onAddGroup,
  onAddTable,
  onEdit,
  onDelete,
  onMove
}: Props) {
  const confirmDelete = (item: SetupStructureItem) => {
    Alert.alert(
      item.type === 'group' ? 'Gruppe löschen' : 'Tabelle löschen',
      `„${item.name}" wirklich entfernen?`,
      [
        { text: 'Abbrechen', style: 'cancel' },
        { text: 'Löschen', style: 'destructive', onPress: () => onDelete(item) }
      ]
    );
  };

  return (
    <ScrollView
      style={styles.scroll}
      contentContainerStyle={styles.content}
      keyboardShouldPersistTaps="handled"
      showsVerticalScrollIndicator={false}
    >
      <Text style={styles.heading}>Meine Struktur</Text>

      {!readOnly ? (
        <View style={styles.actions}>
          <Pressable accessibilityRole="button" style={styles.actionBtn} onPress={onAddGroup}>
            <MaterialCommunityIcons name="plus" size={20} color={colors.accent} />
            <Text style={styles.actionLabel}>Gruppe hinzufügen</Text>
          </Pressable>
          <Pressable accessibilityRole="button" style={styles.actionBtn} onPress={onAddTable}>
            <MaterialCommunityIcons name="plus" size={20} color={colors.accent} />
            <Text style={styles.actionLabel}>Tabelle hinzufügen</Text>
          </Pressable>
        </View>
      ) : null}

      {items.length === 0 ? (
        <View style={styles.emptyCard}>
          <Text style={styles.emptyTitle}>Noch keine Bereiche</Text>
          <Text style={styles.emptyCopy}>
            Lege Gruppen und Tabellen an, die dein digitales Bautagebuch strukturieren.
          </Text>
        </View>
      ) : (
        <View style={styles.list}>
          {items.map((item, index) => (
            <View key={item.id} style={styles.card}>
              <View style={styles.cardHeader}>
                <MaterialCommunityIcons name={itemIcon(item)} size={22} color={colors.accent2} />
                <Pressable
                  accessibilityRole="button"
                  style={styles.cardTitleWrap}
                  onPress={() => onEdit(item)}
                  disabled={readOnly}
                >
                  <Text style={styles.cardTitle}>{item.name}</Text>
                </Pressable>
                {!readOnly ? (
                  <View style={styles.cardActions}>
                    <Pressable accessibilityRole="button" onPress={() => onEdit(item)} hitSlop={8}>
                      <MaterialCommunityIcons name="pencil-outline" size={20} color={colors.muted} />
                    </Pressable>
                    <Pressable accessibilityRole="button" onPress={() => confirmDelete(item)} hitSlop={8}>
                      <MaterialCommunityIcons name="trash-can-outline" size={20} color={colors.danger} />
                    </Pressable>
                  </View>
                ) : null}
              </View>

              <View style={styles.cardBody}>
                <Text style={styles.metaLabel}>Typ</Text>
                <Text style={styles.metaValue}>
                  {item.type === 'group' ? 'Gruppe' : 'Tabelle'}
                </Text>

                {item.type === 'group' && item.description ? (
                  <>
                    <Text style={styles.metaLabel}>Beschreibung</Text>
                    <Text style={styles.metaValue}>{item.description}</Text>
                  </>
                ) : null}

                {item.type === 'table' ? (
                  <>
                    <Text style={styles.metaLabel}>Spalten</Text>
                    <View style={styles.columnList}>
                      {item.columns.map((column) => (
                        <Text key={column.id} style={styles.columnItem}>
                          • {column.name}
                        </Text>
                      ))}
                    </View>
                  </>
                ) : null}
              </View>

              {!readOnly ? (
                <View style={styles.reorderRow}>
                  <MaterialCommunityIcons name="drag-vertical" size={22} color={colors.muted} />
                  <Text style={styles.reorderHint}>Reihenfolge</Text>
                  <View style={styles.moveActions}>
                    <PrimaryButton
                      label="↑"
                      variant="ghost"
                      disabled={index === 0}
                      onPress={() => onMove(item.id, -1)}
                    />
                    <PrimaryButton
                      label="↓"
                      variant="ghost"
                      disabled={index === items.length - 1}
                      onPress={() => onMove(item.id, 1)}
                    />
                  </View>
                </View>
              ) : null}
            </View>
          ))}
        </View>
      )}
    </ScrollView>
  );
}

const styles = StyleSheet.create({
  scroll: {
    flex: 1
  },
  content: {
    padding: spacing.pageX,
    gap: spacing.md,
    paddingBottom: spacing.xl
  },
  heading: {
    ...typography.title,
    color: colors.ink
  },
  actions: {
    gap: spacing.sm
  },
  actionBtn: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: spacing.sm,
    minHeight: spacing.touchMin + 4,
    paddingHorizontal: spacing.md,
    borderRadius: spacing.cardRadius,
    borderWidth: 1,
    borderColor: colors.border,
    backgroundColor: colors.panelElevated,
    ...shadows.card
  },
  actionLabel: {
    ...typography.bodyStrong,
    color: colors.accent2
  },
  emptyCard: {
    borderRadius: spacing.cardRadius,
    borderWidth: 1,
    borderColor: colors.border,
    backgroundColor: colors.panel,
    padding: spacing.lg,
    gap: spacing.xs
  },
  emptyTitle: {
    ...typography.subtitle,
    color: colors.ink
  },
  emptyCopy: {
    ...typography.body,
    color: colors.muted
  },
  list: {
    gap: spacing.sm
  },
  card: {
    borderRadius: spacing.cardRadius,
    borderWidth: 1,
    borderColor: colors.border,
    backgroundColor: colors.panelElevated,
    overflow: 'hidden',
    ...shadows.card
  },
  cardHeader: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: spacing.sm,
    paddingHorizontal: spacing.md,
    paddingTop: spacing.md,
    paddingBottom: spacing.xs
  },
  cardTitleWrap: {
    flex: 1,
    minWidth: 0
  },
  cardTitle: {
    ...typography.subtitle,
    color: colors.ink
  },
  cardActions: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: spacing.sm
  },
  cardBody: {
    paddingHorizontal: spacing.md,
    paddingBottom: spacing.sm,
    gap: spacing.xxs
  },
  metaLabel: {
    ...typography.caption,
    color: colors.muted,
    marginTop: spacing.xxs
  },
  metaValue: {
    ...typography.body,
    color: colors.ink
  },
  columnList: {
    gap: 2,
    paddingTop: 2
  },
  columnItem: {
    ...typography.body,
    color: colors.ink
  },
  reorderRow: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: spacing.sm,
    paddingHorizontal: spacing.md,
    paddingVertical: spacing.sm,
    borderTopWidth: 1,
    borderTopColor: colors.border,
    backgroundColor: colors.panel
  },
  reorderHint: {
    ...typography.caption,
    color: colors.muted,
    flex: 1
  },
  moveActions: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: 2
  }
});

import { Alert, Pressable, ScrollView, StyleSheet, Text, View } from 'react-native';
import { MaterialCommunityIcons } from '@expo/vector-icons';

import { colors, shadows, spacing, typography } from '../../../../constants/theme';
import { hapticLight, hapticSelection } from '../../../../lib/haptics';
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

type ItemVisual = {
  icon: keyof typeof MaterialCommunityIcons.glyphMap;
  tint: string;
  bg: string;
  typeLabel: string;
};

function itemVisual(item: SetupStructureItem): ItemVisual {
  if (item.type === 'group') {
    return {
      icon: 'file-document-outline',
      tint: colors.accent2,
      bg: 'rgba(36, 50, 64, 0.1)',
      typeLabel: 'Gruppe'
    };
  }
  return {
    icon: 'table',
    tint: colors.info,
    bg: 'rgba(42, 95, 143, 0.12)',
    typeLabel: 'Tabelle'
  };
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
        {
          text: 'Löschen',
          style: 'destructive',
          onPress: () => {
            void hapticLight();
            onDelete(item);
          }
        }
      ]
    );
  };

  const handleAddGroup = () => {
    void hapticSelection();
    onAddGroup();
  };

  const handleAddTable = () => {
    void hapticSelection();
    onAddTable();
  };

  const handleMove = (id: string, direction: -1 | 1) => {
    void hapticSelection();
    onMove(id, direction);
  };

  return (
    <ScrollView
      style={styles.scroll}
      contentContainerStyle={styles.content}
      keyboardShouldPersistTaps="handled"
      showsVerticalScrollIndicator={false}
    >
      <View style={styles.intro}>
        <Text style={styles.heading}>Meine Struktur</Text>
        <Text style={styles.introCopy}>
          Lege die Bereiche fest, die später im Bautagebuch erscheinen. Die Reihenfolge kannst du
          jederzeit anpassen.
        </Text>
      </View>

      {!readOnly ? (
        <View style={styles.actions}>
          <Pressable
            accessibilityRole="button"
            style={({ pressed }) => [styles.actionTile, styles.actionTileGroup, pressed ? styles.actionPressed : null]}
            onPress={handleAddGroup}
          >
            <View style={[styles.actionIconWrap, styles.actionIconGroup]}>
              <MaterialCommunityIcons name="folder-plus-outline" size={24} color={colors.accent2} />
            </View>
            <Text style={styles.actionTitle}>Gruppe</Text>
            <Text style={styles.actionCopy}>Formularbereich</Text>
          </Pressable>
          <Pressable
            accessibilityRole="button"
            style={({ pressed }) => [styles.actionTile, styles.actionTileTable, pressed ? styles.actionPressed : null]}
            onPress={handleAddTable}
          >
            <View style={[styles.actionIconWrap, styles.actionIconTable]}>
              <MaterialCommunityIcons name="table-plus" size={24} color={colors.info} />
            </View>
            <Text style={styles.actionTitle}>Tabelle</Text>
            <Text style={styles.actionCopy}>Zeilen & Spalten</Text>
          </Pressable>
        </View>
      ) : null}

      {items.length === 0 ? (
        <View style={styles.emptyCard}>
          <View style={styles.emptyIconWrap}>
            <MaterialCommunityIcons name="shape-plus-outline" size={32} color={colors.accent} />
          </View>
          <Text style={styles.emptyTitle}>Noch keine Bereiche</Text>
          <Text style={styles.emptyCopy}>
            Schau dir die PDF-Vorlage an und lege passende Gruppen oder Tabellen an.
          </Text>
        </View>
      ) : (
        <View style={styles.list}>
          <View style={styles.listHeader}>
            <Text style={styles.listHeading}>Reihenfolge im Bautagebuch</Text>
            <View style={styles.countBadge}>
              <Text style={styles.countText}>{items.length}</Text>
            </View>
          </View>

          {items.map((item, index) => {
            const visual = itemVisual(item);
            return (
              <Pressable
                key={item.id}
                accessibilityRole="button"
                style={({ pressed }) => [styles.card, pressed && !readOnly ? styles.cardPressed : null]}
                onPress={() => {
                  if (readOnly) return;
                  void hapticSelection();
                  onEdit(item);
                }}
                disabled={readOnly}
              >
                <View style={[styles.accentStrip, { backgroundColor: visual.tint }]} />
                <View style={styles.cardInner}>
                  <View style={styles.cardTop}>
                    <View style={styles.indexBadge}>
                      <Text style={styles.indexText}>{index + 1}</Text>
                    </View>
                    <View style={[styles.iconWrap, { backgroundColor: visual.bg }]}>
                      <MaterialCommunityIcons name={visual.icon} size={22} color={visual.tint} />
                    </View>
                    <View style={styles.cardCopy}>
                      <Text style={styles.cardTitle} numberOfLines={2}>
                        {item.name}
                      </Text>
                      <View style={styles.typePill}>
                        <Text style={styles.typePillText}>{visual.typeLabel}</Text>
                      </View>
                    </View>
                    {!readOnly ? (
                      <View style={styles.cardActions}>
                        <Pressable
                          accessibilityRole="button"
                          accessibilityLabel="Bearbeiten"
                          style={styles.iconBtn}
                          onPress={(event) => {
                            event.stopPropagation();
                            void hapticSelection();
                            onEdit(item);
                          }}
                          hitSlop={8}
                        >
                          <MaterialCommunityIcons name="pencil-outline" size={20} color={colors.muted} />
                        </Pressable>
                        <Pressable
                          accessibilityRole="button"
                          accessibilityLabel="Löschen"
                          style={styles.iconBtn}
                          onPress={(event) => {
                            event.stopPropagation();
                            confirmDelete(item);
                          }}
                          hitSlop={8}
                        >
                          <MaterialCommunityIcons name="trash-can-outline" size={20} color={colors.danger} />
                        </Pressable>
                      </View>
                    ) : null}
                  </View>

                  {item.type === 'group' && item.description ? (
                    <Text style={styles.description} numberOfLines={3}>
                      {item.description}
                    </Text>
                  ) : null}

                  {item.type === 'table' ? (
                    <View style={styles.columnWrap}>
                      <Text style={styles.columnLabel}>Spalten</Text>
                      <View style={styles.columnChips}>
                        {item.columns.map((column) => (
                          <View key={column.id} style={styles.columnChip}>
                            <Text style={styles.columnChipText}>{column.name}</Text>
                          </View>
                        ))}
                      </View>
                    </View>
                  ) : null}

                  {!readOnly ? (
                    <View style={styles.reorderRow}>
                      <MaterialCommunityIcons name="drag-vertical" size={20} color={colors.muted} />
                      <Text style={styles.reorderHint}>Position im Bautagebuch</Text>
                      <View style={styles.moveActions}>
                        <Pressable
                          accessibilityRole="button"
                          accessibilityLabel="Nach oben"
                          style={[styles.moveBtn, index === 0 ? styles.moveBtnDisabled : null]}
                          disabled={index === 0}
                          onPress={(event) => {
                            event.stopPropagation();
                            handleMove(item.id, -1);
                          }}
                        >
                          <MaterialCommunityIcons
                            name="chevron-up"
                            size={22}
                            color={index === 0 ? colors.borderStrong : colors.accent2}
                          />
                        </Pressable>
                        <Pressable
                          accessibilityRole="button"
                          accessibilityLabel="Nach unten"
                          style={[
                            styles.moveBtn,
                            index === items.length - 1 ? styles.moveBtnDisabled : null
                          ]}
                          disabled={index === items.length - 1}
                          onPress={(event) => {
                            event.stopPropagation();
                            handleMove(item.id, 1);
                          }}
                        >
                          <MaterialCommunityIcons
                            name="chevron-down"
                            size={22}
                            color={index === items.length - 1 ? colors.borderStrong : colors.accent2}
                          />
                        </Pressable>
                      </View>
                    </View>
                  ) : null}
                </View>
              </Pressable>
            );
          })}
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
    paddingBottom: spacing.xxl
  },
  intro: {
    gap: spacing.xxs
  },
  heading: {
    ...typography.title,
    color: colors.ink
  },
  introCopy: {
    ...typography.body,
    color: colors.muted,
    lineHeight: 22
  },
  actions: {
    flexDirection: 'row',
    gap: spacing.sm
  },
  actionTile: {
    flex: 1,
    minHeight: 112,
    borderRadius: spacing.cardRadius,
    borderWidth: 1,
    padding: spacing.md,
    gap: spacing.xxs,
    ...shadows.card
  },
  actionTileGroup: {
    borderColor: colors.border,
    backgroundColor: colors.panelElevated
  },
  actionTileTable: {
    borderColor: colors.border,
    backgroundColor: colors.panelElevated
  },
  actionPressed: {
    opacity: 0.92,
    transform: [{ scale: 0.98 }]
  },
  actionIconWrap: {
    width: 44,
    height: 44,
    borderRadius: 12,
    alignItems: 'center',
    justifyContent: 'center',
    marginBottom: spacing.xxs
  },
  actionIconGroup: {
    backgroundColor: 'rgba(36, 50, 64, 0.1)'
  },
  actionIconTable: {
    backgroundColor: 'rgba(42, 95, 143, 0.12)'
  },
  actionTitle: {
    ...typography.bodyStrong,
    color: colors.ink
  },
  actionCopy: {
    ...typography.caption,
    color: colors.muted
  },
  emptyCard: {
    alignItems: 'center',
    borderRadius: spacing.cardRadius,
    borderWidth: 1,
    borderColor: colors.border,
    backgroundColor: colors.panel,
    padding: spacing.xl,
    gap: spacing.sm
  },
  emptyIconWrap: {
    width: 64,
    height: 64,
    borderRadius: 999,
    alignItems: 'center',
    justifyContent: 'center',
    backgroundColor: colors.badgeBg
  },
  emptyTitle: {
    ...typography.subtitle,
    color: colors.ink,
    textAlign: 'center'
  },
  emptyCopy: {
    ...typography.body,
    color: colors.muted,
    textAlign: 'center',
    lineHeight: 22
  },
  list: {
    gap: spacing.sm
  },
  listHeader: {
    flexDirection: 'row',
    alignItems: 'center',
    justifyContent: 'space-between',
    gap: spacing.sm
  },
  listHeading: {
    ...typography.label,
    color: colors.muted,
    flex: 1
  },
  countBadge: {
    minWidth: 28,
    height: 28,
    borderRadius: 999,
    alignItems: 'center',
    justifyContent: 'center',
    backgroundColor: colors.accent2
  },
  countText: {
    ...typography.caption,
    color: colors.white,
    fontFamily: 'SpaceGrotesk_600SemiBold'
  },
  card: {
    borderRadius: spacing.cardRadius,
    borderWidth: 1,
    borderColor: colors.border,
    backgroundColor: colors.panelElevated,
    overflow: 'hidden',
    ...shadows.card
  },
  cardPressed: {
    opacity: 0.96
  },
  accentStrip: {
    height: 3
  },
  cardInner: {
    padding: spacing.md,
    gap: spacing.sm
  },
  cardTop: {
    flexDirection: 'row',
    alignItems: 'flex-start',
    gap: spacing.sm
  },
  indexBadge: {
    width: 24,
    height: 24,
    borderRadius: 999,
    alignItems: 'center',
    justifyContent: 'center',
    backgroundColor: colors.bg,
    borderWidth: 1,
    borderColor: colors.border
  },
  indexText: {
    ...typography.caption,
    color: colors.muted,
    fontFamily: 'SpaceGrotesk_600SemiBold'
  },
  iconWrap: {
    width: 40,
    height: 40,
    borderRadius: 12,
    alignItems: 'center',
    justifyContent: 'center'
  },
  cardCopy: {
    flex: 1,
    minWidth: 0,
    gap: spacing.xxs
  },
  cardTitle: {
    ...typography.subtitle,
    color: colors.ink
  },
  typePill: {
    alignSelf: 'flex-start',
    paddingHorizontal: spacing.sm,
    paddingVertical: 3,
    borderRadius: 999,
    backgroundColor: colors.bg,
    borderWidth: 1,
    borderColor: colors.border
  },
  typePillText: {
    ...typography.caption,
    color: colors.muted,
    fontFamily: 'SpaceGrotesk_600SemiBold'
  },
  cardActions: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: 2
  },
  iconBtn: {
    width: 36,
    height: 36,
    borderRadius: 10,
    alignItems: 'center',
    justifyContent: 'center'
  },
  description: {
    ...typography.body,
    color: colors.muted,
    lineHeight: 22,
    paddingLeft: 36
  },
  columnWrap: {
    gap: spacing.xxs,
    paddingLeft: 36
  },
  columnLabel: {
    ...typography.caption,
    color: colors.muted
  },
  columnChips: {
    flexDirection: 'row',
    flexWrap: 'wrap',
    gap: spacing.xxs
  },
  columnChip: {
    paddingHorizontal: spacing.sm,
    paddingVertical: 4,
    borderRadius: 999,
    backgroundColor: colors.bg,
    borderWidth: 1,
    borderColor: colors.border
  },
  columnChipText: {
    ...typography.caption,
    color: colors.ink
  },
  reorderRow: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: spacing.sm,
    paddingTop: spacing.xs,
    borderTopWidth: 1,
    borderTopColor: colors.border
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
  },
  moveBtn: {
    width: 40,
    height: 40,
    borderRadius: 10,
    alignItems: 'center',
    justifyContent: 'center',
    backgroundColor: colors.bg,
    borderWidth: 1,
    borderColor: colors.border
  },
  moveBtnDisabled: {
    opacity: 0.45
  }
});

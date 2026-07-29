import { useMemo, useState, type ReactNode } from 'react';
import {
  LayoutAnimation,
  Modal,
  Platform,
  Pressable,
  ScrollView,
  StyleSheet,
  Text,
  UIManager,
  View
} from 'react-native';
import { MaterialCommunityIcons } from '@expo/vector-icons';
import { useSafeAreaInsets } from 'react-native-safe-area-context';

import { PrimaryButton, SingleLineText } from '../../../../components/mobile';
import { colors, shadows, spacing, typography } from '../../../../constants/theme';
import { hapticSelection } from '../../../../lib/haptics';
import { systemBottomInset } from '../../../../navigation/systemInsets';
import {
  buildFieldDetailRows,
  buildGroupFieldSummaryRows,
  buildTemplateDetailOverview,
  formatTemplateUpdatedAt,
  structureItemIcon,
  structureItemSummary,
  type TemplateDetailField,
  type TemplateDetailGroup,
  type TemplateDetailItem
} from '../../lib/template-detail';
import type { DetectedField } from '../../types';

if (Platform.OS === 'android' && UIManager.setLayoutAnimationEnabledExperimental) {
  UIManager.setLayoutAnimationEnabledExperimental(true);
}

type Props = {
  templateName: string;
  setupModel: Record<string, unknown>;
  detectedFields: DetectedField[];
  readOnly?: boolean;
  onBack: () => void;
  onEdit?: () => void;
};

function DetailInfoSheet({
  visible,
  title,
  onClose,
  children
}: {
  visible: boolean;
  title: string;
  onClose: () => void;
  children: ReactNode;
}) {
  const insets = useSafeAreaInsets();

  return (
    <Modal visible={visible} animationType="slide" onRequestClose={onClose}>
      <View style={[styles.sheetRoot, { paddingTop: insets.top }]}>
        <View style={styles.sheetHeader}>
          <Pressable accessibilityRole="button" style={styles.sheetCloseBtn} onPress={onClose}>
            <Text style={styles.sheetCloseLabel}>Schließen</Text>
          </Pressable>
          <Text style={styles.sheetTitle}>{title}</Text>
          <View style={styles.sheetCloseBtn} />
        </View>
        <ScrollView
          contentContainerStyle={[
            styles.sheetBody,
            { paddingBottom: systemBottomInset(insets) + spacing.lg }
          ]}
          showsVerticalScrollIndicator={false}
        >
          {children}
        </ScrollView>
      </View>
    </Modal>
  );
}

function InfoRows({ rows }: { rows: Array<{ label: string; value: string; checked?: boolean }> }) {
  return (
    <View style={styles.infoBlock}>
      {rows.map((row) => (
        <View key={row.label} style={styles.infoRow}>
          <Text style={styles.infoLabel}>{row.label}</Text>
          {row.checked !== undefined ? (
            <MaterialCommunityIcons
              name={row.checked ? 'check-circle' : 'close-circle-outline'}
              size={20}
              color={row.checked ? colors.success : colors.muted}
            />
          ) : (
            <Text style={styles.infoValue}>{row.value}</Text>
          )}
        </View>
      ))}
    </View>
  );
}

function StructureCard({
  item,
  index,
  expanded,
  onToggle,
  onOpenGroup,
  onOpenField
}: {
  item: TemplateDetailItem;
  index: number;
  expanded: boolean;
  onToggle: () => void;
  onOpenGroup?: (group: TemplateDetailGroup) => void;
  onOpenField?: (field: TemplateDetailField, groupName: string) => void;
}) {
  const iconName = structureItemIcon(
    item.kind === 'group'
      ? { id: item.id, name: item.name, type: 'group', order: index }
      : {
          id: item.id,
          name: item.name,
          type: 'table',
          order: index,
          columns: item.columns.map((col, colIndex) => ({
            id: col.id,
            name: col.name,
            order: colIndex
          }))
        },
    index
  ) as keyof typeof MaterialCommunityIcons.glyphMap;

  const summary =
    item.kind === 'group'
      ? structureItemSummary(
          { id: item.id, name: item.name, type: 'group', order: index },
          item.fieldCount
        )
      : 'Tabelle';

  return (
    <View style={[styles.card, shadows.card]}>
      <Pressable
        accessibilityRole="button"
        style={styles.cardHeader}
        onPress={() => {
          void hapticSelection();
          LayoutAnimation.configureNext(LayoutAnimation.Presets.easeInEaseOut);
          onToggle();
        }}
      >
        <View style={styles.cardIconWrap}>
          <MaterialCommunityIcons name={iconName} size={22} color={colors.accent2} />
        </View>
        <View style={styles.cardTitleBlock}>
          <SingleLineText style={styles.cardTitle}>{item.name}</SingleLineText>
          <Text style={styles.cardMeta}>{summary}</Text>
        </View>
        <MaterialCommunityIcons
          name={expanded ? 'chevron-up' : 'chevron-down'}
          size={22}
          color={colors.muted}
        />
      </Pressable>

      {expanded ? (
        <View style={styles.cardBody}>
          {item.kind === 'group' ? (
            <>
              {item.description ? (
                <Text style={styles.cardDescription}>{item.description}</Text>
              ) : null}
              <Text style={styles.cardSectionLabel}>Felder:</Text>
              {item.fields.length === 0 ? (
                <Text style={styles.emptyHint}>Keine Felder zugeordnet</Text>
              ) : (
                item.fields.map((field) => (
                  <Pressable
                    key={field.fieldId}
                    accessibilityRole="button"
                    style={styles.fieldRow}
                    onPress={() => onOpenField?.(field, item.name)}
                  >
                    <Text style={styles.fieldRowLabel}>{field.label}</Text>
                    <MaterialCommunityIcons name="chevron-right" size={18} color={colors.muted} />
                  </Pressable>
                ))
              )}
              <Pressable
                accessibilityRole="button"
                style={styles.cardDetailLink}
                onPress={() => onOpenGroup?.(item)}
              >
                <Text style={styles.cardDetailLinkLabel}>Gruppendetails anzeigen</Text>
              </Pressable>
            </>
          ) : (
            <>
              <View style={styles.tableMetaRow}>
                <Text style={styles.tableMetaLabel}>Typ:</Text>
                <Text style={styles.tableMetaValue}>Tabelle</Text>
              </View>
              <Text style={styles.cardSectionLabel}>Spalten:</Text>
              {item.columns.map((column) => (
                <View key={column.id} style={styles.columnRow}>
                  <Text style={styles.columnLabel}>{column.name}</Text>
                </View>
              ))}
            </>
          )}
        </View>
      ) : null}
    </View>
  );
}

export function TemplateDetailView({
  templateName,
  setupModel,
  detectedFields,
  readOnly = false,
  onBack,
  onEdit
}: Props) {
  const insets = useSafeAreaInsets();
  const overview = useMemo(
    () => buildTemplateDetailOverview(setupModel, detectedFields),
    [setupModel, detectedFields]
  );
  const [expandedIds, setExpandedIds] = useState<Set<string>>(new Set());
  const [activeGroup, setActiveGroup] = useState<TemplateDetailGroup | null>(null);
  const [activeField, setActiveField] = useState<{
    field: TemplateDetailField;
    groupName: string;
  } | null>(null);

  const toggleExpanded = (id: string) => {
    setExpandedIds((prev) => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  };

  return (
    <View style={[styles.root, { paddingTop: insets.top }]}>
      <View style={styles.header}>
        <Pressable accessibilityRole="button" style={styles.backBtn} onPress={onBack}>
          <MaterialCommunityIcons name="chevron-left" size={22} color={colors.accent} />
          <Text style={styles.backLabel}>Vorlagen</Text>
        </Pressable>
        {!readOnly && onEdit ? (
          <Pressable accessibilityRole="button" style={styles.editBtn} onPress={onEdit}>
            <MaterialCommunityIcons name="pencil-outline" size={18} color={colors.accent2} />
            <Text style={styles.editLabel}>Bearbeiten</Text>
          </Pressable>
        ) : (
          <View style={styles.editBtnPlaceholder} />
        )}
      </View>

      <ScrollView
        contentContainerStyle={[
          styles.content,
          { paddingBottom: systemBottomInset(insets) + spacing.xl }
        ]}
        showsVerticalScrollIndicator={false}
      >
        <Text style={styles.templateName}>{templateName}</Text>
        <Text style={styles.updatedAt}>
          Letzte Änderung:{'\n'}
          {formatTemplateUpdatedAt(overview.updatedAt)}
        </Text>

        <View style={styles.divider} />

        <Text style={styles.sectionHeading}>Struktur</Text>

        <View style={styles.cardList}>
          {overview.items.map((item, index) => (
            <StructureCard
              key={item.id}
              item={item}
              index={index}
              expanded={expandedIds.has(item.id)}
              onToggle={() => toggleExpanded(item.id)}
              onOpenGroup={setActiveGroup}
              onOpenField={(field, groupName) => setActiveField({ field, groupName })}
            />
          ))}
        </View>

        {overview.items.length === 0 ? (
          <View style={styles.emptyCard}>
            <Text style={styles.emptyHint}>Noch keine Struktur vorhanden.</Text>
            {!readOnly && onEdit ? (
              <PrimaryButton label="Einrichtung starten" variant="secondary" onPress={onEdit} />
            ) : null}
          </View>
        ) : null}
      </ScrollView>

      <DetailInfoSheet
        visible={Boolean(activeGroup)}
        title={activeGroup?.name || 'Gruppe'}
        onClose={() => setActiveGroup(null)}
      >
        {activeGroup?.fields.map((field) => (
          <Pressable
            key={field.fieldId}
            accessibilityRole="button"
            style={styles.groupFieldBlock}
            onPress={() => {
              const groupName = activeGroup.name;
              setActiveGroup(null);
              setActiveField({ field, groupName });
            }}
          >
            <Text style={styles.groupFieldTitle}>{field.label}</Text>
            <InfoRows rows={buildGroupFieldSummaryRows(field)} />
          </Pressable>
        ))}
      </DetailInfoSheet>

      <DetailInfoSheet
        visible={Boolean(activeField)}
        title={activeField?.field.label || 'Feld'}
        onClose={() => setActiveField(null)}
      >
        {activeField ? (
          <>
            <Text style={styles.fieldDetailHeading}>Feld: {activeField.field.label}</Text>
            {activeField.groupName ? (
              <Text style={styles.fieldDetailGroup}>Gruppe: {activeField.groupName}</Text>
            ) : null}
            <Text style={styles.settingsHeading}>Einstellungen:</Text>
            <InfoRows rows={buildFieldDetailRows(activeField.field.config, detectedFields)} />
          </>
        ) : null}
      </DetailInfoSheet>
    </View>
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
    paddingBottom: spacing.sm
  },
  backBtn: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: spacing.xxs,
    minHeight: spacing.touchMin
  },
  backLabel: {
    ...typography.bodyStrong,
    color: colors.accent
  },
  editBtn: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: spacing.xxs,
    minHeight: spacing.touchMin,
    paddingHorizontal: spacing.sm,
    borderRadius: 999,
    backgroundColor: colors.panelElevated,
    borderWidth: 1,
    borderColor: colors.border
  },
  editBtnPlaceholder: {
    width: 96
  },
  editLabel: {
    ...typography.caption,
    color: colors.accent2,
    fontFamily: 'SpaceGrotesk_600SemiBold'
  },
  content: {
    paddingHorizontal: spacing.pageX,
    gap: spacing.sm
  },
  templateName: {
    ...typography.title,
    color: colors.ink,
    fontSize: 28
  },
  updatedAt: {
    ...typography.caption,
    color: colors.muted,
    lineHeight: 20
  },
  divider: {
    height: 1,
    backgroundColor: colors.border,
    marginVertical: spacing.sm
  },
  sectionHeading: {
    ...typography.subtitle,
    color: colors.ink,
    marginBottom: spacing.xxs
  },
  cardList: {
    gap: spacing.cardGap
  },
  card: {
    borderRadius: spacing.cardRadius,
    backgroundColor: colors.panelElevated,
    borderWidth: 1,
    borderColor: colors.border,
    overflow: 'hidden'
  },
  cardHeader: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: spacing.sm,
    padding: spacing.md,
    minHeight: spacing.touchMin + spacing.sm
  },
  cardIconWrap: {
    width: 40,
    height: 40,
    borderRadius: spacing.iconRadius,
    alignItems: 'center',
    justifyContent: 'center',
    backgroundColor: colors.badgeBg
  },
  cardTitleBlock: {
    flex: 1,
    gap: 2,
    minWidth: 0
  },
  cardTitle: {
    ...typography.bodyStrong,
    color: colors.ink,
    fontSize: 17
  },
  cardMeta: {
    ...typography.caption,
    color: colors.muted
  },
  cardBody: {
    paddingHorizontal: spacing.md,
    paddingBottom: spacing.md,
    gap: spacing.xs,
    borderTopWidth: 1,
    borderTopColor: colors.border
  },
  cardDescription: {
    ...typography.caption,
    color: colors.muted
  },
  cardSectionLabel: {
    ...typography.label,
    color: colors.muted,
    marginTop: spacing.xxs
  },
  fieldRow: {
    flexDirection: 'row',
    alignItems: 'center',
    justifyContent: 'space-between',
    minHeight: spacing.touchMin,
    paddingVertical: spacing.xxs
  },
  fieldRowLabel: {
    ...typography.body,
    color: colors.ink,
    flex: 1
  },
  cardDetailLink: {
    marginTop: spacing.xs,
    minHeight: spacing.touchMin,
    justifyContent: 'center'
  },
  cardDetailLinkLabel: {
    ...typography.caption,
    color: colors.accent,
    fontFamily: 'SpaceGrotesk_600SemiBold'
  },
  tableMetaRow: {
    flexDirection: 'row',
    gap: spacing.xs
  },
  tableMetaLabel: {
    ...typography.caption,
    color: colors.muted
  },
  tableMetaValue: {
    ...typography.caption,
    color: colors.ink,
    fontFamily: 'SpaceGrotesk_600SemiBold'
  },
  columnRow: {
    paddingVertical: spacing.xxs
  },
  columnLabel: {
    ...typography.body,
    color: colors.ink
  },
  emptyCard: {
    padding: spacing.lg,
    alignItems: 'center',
    gap: spacing.md
  },
  emptyHint: {
    ...typography.body,
    color: colors.muted,
    textAlign: 'center'
  },
  sheetRoot: {
    flex: 1,
    backgroundColor: colors.bg
  },
  sheetHeader: {
    flexDirection: 'row',
    alignItems: 'center',
    justifyContent: 'space-between',
    paddingHorizontal: spacing.pageX,
    paddingBottom: spacing.sm,
    borderBottomWidth: 1,
    borderBottomColor: colors.border,
    backgroundColor: colors.panel
  },
  sheetCloseBtn: {
    minWidth: 72,
    minHeight: spacing.touchMin,
    justifyContent: 'center'
  },
  sheetCloseLabel: {
    ...typography.bodyStrong,
    color: colors.accent
  },
  sheetTitle: {
    ...typography.subtitle,
    color: colors.ink,
    flex: 1,
    textAlign: 'center'
  },
  sheetBody: {
    padding: spacing.pageX,
    gap: spacing.md
  },
  groupFieldBlock: {
    padding: spacing.md,
    borderRadius: spacing.cardRadius,
    backgroundColor: colors.panelElevated,
    borderWidth: 1,
    borderColor: colors.border,
    gap: spacing.sm
  },
  groupFieldTitle: {
    ...typography.bodyStrong,
    color: colors.ink,
    fontSize: 17
  },
  fieldDetailHeading: {
    ...typography.subtitle,
    color: colors.ink
  },
  fieldDetailGroup: {
    ...typography.caption,
    color: colors.muted
  },
  settingsHeading: {
    ...typography.label,
    color: colors.muted,
    marginTop: spacing.sm
  },
  infoBlock: {
    gap: spacing.sm
  },
  infoRow: {
    flexDirection: 'row',
    alignItems: 'center',
    justifyContent: 'space-between',
    gap: spacing.sm,
    minHeight: 32
  },
  infoLabel: {
    ...typography.caption,
    color: colors.muted,
    flex: 1
  },
  infoValue: {
    ...typography.body,
    color: colors.ink,
    textAlign: 'right',
    flex: 1
  }
});

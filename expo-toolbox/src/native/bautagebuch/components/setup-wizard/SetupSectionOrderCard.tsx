import { Pressable, StyleSheet, Text, View } from 'react-native';
import { MaterialCommunityIcons } from '@expo/vector-icons';

import { PrimaryButton } from '../../../../components/mobile';
import { colors, shadows, spacing, typography } from '../../../../constants/theme';
import {
  sectionEntryKey,
  type OrderedSectionEntry
} from '../../lib/setup-section-order';

type Props = {
  sections: OrderedSectionEntry[];
  selectedKey?: string | null;
  readOnly?: boolean;
  onMove: (index: number, direction: -1 | 1) => void;
  onSelect?: (entry: OrderedSectionEntry) => void;
};

export function SetupSectionOrderCard({
  sections,
  selectedKey = null,
  readOnly = false,
  onMove,
  onSelect
}: Props) {
  if (sections.length <= 1) return null;

  return (
    <View style={styles.card}>
      <View style={styles.accentStrip} />
      <View style={styles.inner}>
        <View style={styles.header}>
          <View style={styles.iconWrap}>
            <MaterialCommunityIcons name="sort-variant" size={22} color={colors.accent} />
          </View>
          <View style={styles.headerCopy}>
            <Text style={styles.title}>Reihenfolge im Bautagebuch</Text>
            <Text style={styles.subtitle}>
              {sections.length} Abschnitte · bestimmt die Abfolge im BTB-Assistenten
            </Text>
          </View>
        </View>

        <Text style={styles.hint}>
          Gruppen und Tabellen von oben nach unten anordnen. Die oberste Position erscheint zuerst
          beim Ausfüllen.
        </Text>

        <View style={styles.list}>
          {sections.map((entry, index) => {
            const key = sectionEntryKey(entry);
            const selected = selectedKey === key;
            const RowWrapper = onSelect ? Pressable : View;

            return (
              <RowWrapper
                key={key}
                style={[styles.row, selected ? styles.rowSelected : null]}
                {...(onSelect
                  ? {
                      onPress: () => onSelect(entry),
                      accessibilityRole: 'button' as const
                    }
                  : {})}
              >
                <View style={styles.indexBadge}>
                  <Text style={styles.indexText}>{index + 1}</Text>
                </View>
                <View style={styles.rowCopy}>
                  <Text style={styles.rowTitle} numberOfLines={2}>
                    {entry.label}
                  </Text>
                  <Text style={styles.rowMeta}>{entry.typeLabel}</Text>
                </View>
                {!readOnly ? (
                  <View style={styles.moveActions}>
                    <PrimaryButton
                      label="↑"
                      variant="ghost"
                      disabled={index === 0}
                      onPress={() => onMove(index, -1)}
                    />
                    <PrimaryButton
                      label="↓"
                      variant="ghost"
                      disabled={index === sections.length - 1}
                      onPress={() => onMove(index, 1)}
                    />
                  </View>
                ) : null}
              </RowWrapper>
            );
          })}
        </View>
      </View>
    </View>
  );
}

const styles = StyleSheet.create({
  card: {
    borderRadius: spacing.cardRadius,
    backgroundColor: colors.panelElevated,
    borderWidth: 1,
    borderColor: colors.accent,
    overflow: 'hidden',
    ...shadows.card
  },
  accentStrip: {
    height: 4,
    backgroundColor: colors.accent
  },
  inner: {
    gap: spacing.sm,
    padding: spacing.md
  },
  header: {
    flexDirection: 'row',
    alignItems: 'flex-start',
    gap: spacing.sm
  },
  iconWrap: {
    width: 40,
    height: 40,
    borderRadius: spacing.iconRadius,
    alignItems: 'center',
    justifyContent: 'center',
    backgroundColor: colors.badgeBg
  },
  headerCopy: {
    flex: 1,
    gap: 2
  },
  title: {
    ...typography.subtitle,
    color: colors.ink
  },
  subtitle: {
    ...typography.caption,
    color: colors.muted
  },
  hint: {
    ...typography.caption,
    color: colors.muted,
    lineHeight: 18
  },
  list: {
    gap: spacing.xs
  },
  row: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: spacing.sm,
    padding: spacing.sm,
    borderRadius: 12,
    borderWidth: 1,
    borderColor: colors.border,
    backgroundColor: colors.panel
  },
  rowSelected: {
    borderColor: colors.accent,
    backgroundColor: colors.badgeBg
  },
  indexBadge: {
    width: 28,
    height: 28,
    borderRadius: 999,
    alignItems: 'center',
    justifyContent: 'center',
    backgroundColor: colors.accent2
  },
  indexText: {
    ...typography.label,
    color: colors.white
  },
  rowCopy: {
    flex: 1,
    gap: 2,
    minWidth: 0
  },
  rowTitle: {
    ...typography.bodyStrong,
    color: colors.ink
  },
  rowMeta: {
    ...typography.caption,
    color: colors.muted
  },
  moveActions: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: 2
  }
});

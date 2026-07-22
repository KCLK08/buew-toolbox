import { Pressable, ScrollView, StyleSheet, Text, View } from 'react-native';

import { colors, spacing, typography } from '../../../../constants/theme';

type GroupItem = {
  sectionId: string;
  label: string;
  count: number;
};

type Props = {
  groups: GroupItem[];
  activeSectionId: string;
  horizontal?: boolean;
  onSelect: (sectionId: string) => void;
};

export function SetupGroupNav({
  groups,
  activeSectionId,
  horizontal = true,
  onSelect
}: Props) {
  return (
    <ScrollView
      horizontal={horizontal}
      showsHorizontalScrollIndicator={false}
      contentContainerStyle={[styles.row, horizontal ? null : styles.column]}
    >
      {groups.map((group) => {
        const active = group.sectionId === activeSectionId;
        return (
          <Pressable
            key={group.sectionId}
            style={[styles.chip, active ? styles.chipActive : null]}
            onPress={() => onSelect(group.sectionId)}
          >
            <Text style={[styles.chipLabel, active ? styles.chipLabelActive : null]}>{group.label}</Text>
            <Text style={[styles.count, active ? styles.countActive : null]}>({group.count})</Text>
          </Pressable>
        );
      })}
    </ScrollView>
  );
}

const styles = StyleSheet.create({
  row: {
    gap: spacing.sm,
    paddingHorizontal: spacing.pageX,
    paddingVertical: spacing.xs
  },
  column: {
    flexDirection: 'column',
    alignItems: 'stretch'
  },
  chip: {
    minHeight: spacing.touchMin,
    borderRadius: 14,
    borderWidth: 1,
    borderColor: colors.border,
    backgroundColor: colors.panel,
    paddingHorizontal: spacing.md,
    paddingVertical: spacing.sm,
    flexDirection: 'row',
    alignItems: 'center',
    gap: spacing.xs
  },
  chipActive: {
    borderColor: colors.accent,
    backgroundColor: colors.badgeBg
  },
  chipLabel: {
    ...typography.bodyStrong,
    color: colors.ink
  },
  chipLabelActive: {
    color: colors.accent
  },
  count: {
    ...typography.caption,
    color: colors.muted
  },
  countActive: {
    color: colors.accent
  }
});

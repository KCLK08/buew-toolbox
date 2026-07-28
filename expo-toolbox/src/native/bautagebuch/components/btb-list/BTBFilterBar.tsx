import { useState } from 'react';
import { Pressable, ScrollView, StyleSheet, Text } from 'react-native';

import { BottomSheet, BottomSheetOption } from '../../../../components/mobile';
import { colors, spacing, typography } from '../../../../constants/theme';
import {
  type BtbListFilters,
  type BtbSortOrder,
  type ProjectListItem
} from '../../lib/btb-filter';
import { formatRunCount } from '../../lib/group-runs-by-calendar';

type FilterSheet = 'sort' | 'view' | 'project' | null;

type Props = {
  filters: BtbListFilters;
  projects: ProjectListItem[];
  onChange: (next: BtbListFilters) => void;
};

function FilterChip({
  label,
  active,
  onPress
}: {
  label: string;
  active?: boolean;
  onPress: () => void;
}) {
  return (
    <Pressable
      accessibilityRole="button"
      style={[styles.chip, active ? styles.chipActive : null]}
      onPress={onPress}
    >
      <Text style={[styles.chipLabel, active ? styles.chipLabelActive : null]} numberOfLines={1}>
        {label}
      </Text>
    </Pressable>
  );
}

export function BTBFilterBar({ filters, projects, onChange }: Props) {
  const [sheet, setSheet] = useState<FilterSheet>(null);

  const patch = (partial: Partial<BtbListFilters>) => onChange({ ...filters, ...partial });

  const selectedProjectLabel =
    projects.find((entry) => entry.projectKey === filters.projectKey)?.projectLabel ?? null;
  const projectChipLabel = selectedProjectLabel
    ? `Projekt: ${selectedProjectLabel}`
    : 'Projekt: Alle';

  return (
    <>
      <ScrollView
        horizontal
        showsHorizontalScrollIndicator={false}
        contentContainerStyle={styles.row}
      >
        <FilterChip
          label={filters.groupMode === 'project' ? 'Ansicht: Projekt' : 'Ansicht: Kalender'}
          active
          onPress={() => setSheet('view')}
        />
        <FilterChip
          label={projectChipLabel}
          active={Boolean(filters.projectKey)}
          onPress={() => setSheet('project')}
        />
        <FilterChip
          label={filters.sortOrder === 'newest' ? 'Sortieren: Neueste' : 'Sortieren: Älteste'}
          onPress={() => setSheet('sort')}
        />
      </ScrollView>

      <BottomSheet visible={sheet === 'view'} title="Ansicht" onClose={() => setSheet(null)}>
        <BottomSheetOption
          label="Kalender"
          description="Jahr → KW → Projekt → BTB"
          selected={filters.groupMode === 'calendar'}
          onPress={() => {
            patch({ groupMode: 'calendar' });
            setSheet(null);
          }}
        />
        <BottomSheetOption
          label="Projekt"
          description="Projektliste mit Drill-down"
          selected={filters.groupMode === 'project'}
          onPress={() => {
            patch({ groupMode: 'project' });
            setSheet(null);
          }}
        />
      </BottomSheet>

      <BottomSheet visible={sheet === 'project'} title="Projekt filtern" onClose={() => setSheet(null)}>
        <BottomSheetOption
          label="Alle Projekte"
          description="Kein Projektfilter aktiv"
          selected={!filters.projectKey}
          onPress={() => {
            patch({ projectKey: null });
            setSheet(null);
          }}
        />
        {projects.map((project) => (
          <BottomSheetOption
            key={project.projectKey}
            label={project.projectLabel}
            description={formatRunCount(project.runCount)}
            selected={filters.projectKey === project.projectKey}
            onPress={() => {
              patch({ projectKey: project.projectKey });
              setSheet(null);
            }}
          />
        ))}
      </BottomSheet>

      <BottomSheet visible={sheet === 'sort'} title="Sortierung" onClose={() => setSheet(null)}>
        {(['newest', 'oldest'] as BtbSortOrder[]).map((order) => (
          <BottomSheetOption
            key={order}
            label={order === 'newest' ? 'Neueste zuerst' : 'Älteste zuerst'}
            selected={filters.sortOrder === order}
            onPress={() => {
              patch({ sortOrder: order });
              setSheet(null);
            }}
          />
        ))}
      </BottomSheet>
    </>
  );
}

const styles = StyleSheet.create({
  row: {
    gap: spacing.xs,
    paddingVertical: spacing.xxs
  },
  chip: {
    minHeight: spacing.touchMin,
    justifyContent: 'center',
    paddingHorizontal: spacing.md,
    paddingVertical: spacing.sm,
    borderRadius: 999,
    borderWidth: 1,
    borderColor: colors.border,
    backgroundColor: colors.panel,
    maxWidth: 240
  },
  chipActive: {
    borderColor: colors.accent,
    backgroundColor: colors.badgeBg
  },
  chipLabel: {
    ...typography.label,
    color: colors.ink
  },
  chipLabelActive: {
    color: colors.accent
  }
});

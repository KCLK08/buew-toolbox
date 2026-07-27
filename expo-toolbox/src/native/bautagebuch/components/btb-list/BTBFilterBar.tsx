import { useMemo, useState } from 'react';
import { Pressable, ScrollView, StyleSheet, Text, View } from 'react-native';

import { BottomSheet, BottomSheetOption } from '../../../../components/mobile';
import { colors, spacing, typography } from '../../../../constants/theme';
import {
  DEFAULT_BTB_FILTERS,
  formatGroupLabel,
  listAvailableWeeks,
  listAvailableYears,
  type BtbListFilters,
  type BtbSortOrder
} from '../../lib/btb-filter';
import type { BautagebuchRunTree } from '../../lib/group-runs-by-calendar';

type FilterSheet = 'sort' | 'project' | 'year' | 'week' | 'view' | null;

type Props = {
  filters: BtbListFilters;
  baseTree: BautagebuchRunTree;
  projectOptions: Array<{ projectKey: string; projectLabel: string; runCount: number }>;
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

export function BTBFilterBar({ filters, baseTree, projectOptions, onChange }: Props) {
  const [sheet, setSheet] = useState<FilterSheet>(null);

  const years = useMemo(() => listAvailableYears(baseTree), [baseTree]);
  const weeks = useMemo(
    () => listAvailableWeeks(baseTree, filters.year),
    [baseTree, filters.year]
  );

  const selectedProject = projectOptions.find((entry) => entry.projectKey === filters.projectKey);
  const selectedWeek = weeks.find((entry) => entry.weekKey === filters.weekKey);

  const patch = (partial: Partial<BtbListFilters>) => onChange({ ...filters, ...partial });

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
          label={filters.sortOrder === 'newest' ? 'Sortieren: Neueste' : 'Sortieren: Älteste'}
          onPress={() => setSheet('sort')}
        />
        <FilterChip
          label={
            selectedProject
              ? `Projekt: ${selectedProject.projectLabel}`
              : 'Filter Projekt'
          }
          active={Boolean(filters.projectKey)}
          onPress={() => setSheet('project')}
        />
        <FilterChip
          label={filters.year ? `Jahr: ${filters.year}` : 'Filter Jahr'}
          active={filters.year !== null}
          onPress={() => setSheet('year')}
        />
        <FilterChip
          label={
            selectedWeek
              ? formatGroupLabel(selectedWeek.weekLabel, selectedWeek.runCount)
              : 'Filter KW'
          }
          active={Boolean(filters.weekKey)}
          onPress={() => setSheet('week')}
        />
        {filters.projectKey || filters.year !== null || filters.weekKey ? (
          <FilterChip
            label="Zurücksetzen"
            onPress={() => onChange({ ...filters, ...DEFAULT_BTB_FILTERS, groupMode: filters.groupMode, sortOrder: filters.sortOrder })}
          />
        ) : null}
      </ScrollView>

      <BottomSheet visible={sheet === 'view'} title="Ansicht" onClose={() => setSheet(null)}>
        <BottomSheetOption
          label="Kalender"
          description="Jahr → KW → Projekt → BTB"
          selected={filters.groupMode === 'calendar'}
          onPress={() => {
            patch({ groupMode: 'calendar', projectKey: null });
            setSheet(null);
          }}
        />
        <BottomSheetOption
          label="Projekt"
          description="Projektliste mit Drill-down"
          selected={filters.groupMode === 'project'}
          onPress={() => {
            patch({ groupMode: 'project', projectKey: null, year: null, weekKey: null });
            setSheet(null);
          }}
        />
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

      <BottomSheet visible={sheet === 'project'} title="Projekt filtern" onClose={() => setSheet(null)}>
        <BottomSheetOption
          label="Alle Projekte"
          selected={!filters.projectKey}
          onPress={() => {
            patch({ projectKey: null, year: null, weekKey: null });
            setSheet(null);
          }}
        />
        {projectOptions.map((project) => (
          <BottomSheetOption
            key={project.projectKey}
            label={formatGroupLabel(project.projectLabel, project.runCount)}
            selected={filters.projectKey === project.projectKey}
            onPress={() => {
              patch({
                projectKey: project.projectKey,
                groupMode: filters.groupMode === 'project' ? 'project' : 'calendar',
                year: null,
                weekKey: null
              });
              setSheet(null);
            }}
          />
        ))}
      </BottomSheet>

      <BottomSheet visible={sheet === 'year'} title="Jahr filtern" onClose={() => setSheet(null)}>
        <BottomSheetOption
          label="Alle Jahre"
          selected={filters.year === null}
          onPress={() => {
            patch({ year: null, weekKey: null });
            setSheet(null);
          }}
        />
        {years.map((year) => {
          const count = baseTree.years.find((entry) => entry.year === year)?.runCount || 0;
          return (
            <BottomSheetOption
              key={year}
              label={formatGroupLabel(String(year), count)}
              selected={filters.year === year}
              onPress={() => {
                patch({ year, weekKey: null });
                setSheet(null);
              }}
            />
          );
        })}
      </BottomSheet>

      <BottomSheet visible={sheet === 'week'} title="Kalenderwoche filtern" onClose={() => setSheet(null)}>
        <BottomSheetOption
          label="Alle Kalenderwochen"
          selected={!filters.weekKey}
          onPress={() => {
            patch({ weekKey: null });
            setSheet(null);
          }}
        />
        {weeks.map((week) => (
          <BottomSheetOption
            key={week.weekKey}
            label={formatGroupLabel(week.weekLabel, week.runCount)}
            description={week.dateRangeLabel}
            selected={filters.weekKey === week.weekKey}
            onPress={() => {
              patch({ weekKey: week.weekKey, year: week.weekYear || filters.year });
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

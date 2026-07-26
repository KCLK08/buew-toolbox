import { useEffect, useRef } from 'react';
import { Pressable, ScrollView, StyleSheet, Text, View } from 'react-native';

import { Card } from '../../../components/mobile';
import { colors, spacing, typography } from '../../../constants/theme';

export type SectionNavItem = {
  sectionId: string;
  label: string;
  progress: 'todo' | 'progress' | 'done';
  missingCount?: number;
};

type Props = {
  sections: SectionNavItem[];
  activeIndex: number;
  onSelect: (index: number) => void;
  totalMissingRequired?: number;
};

const CHIP_WIDTH = 148;

export function RunSectionNav({
  sections,
  activeIndex,
  onSelect,
  totalMissingRequired = 0
}: Props) {
  const scrollRef = useRef<ScrollView>(null);
  const total = sections.length;
  const current = activeIndex + 1;
  const doneCount = sections.filter((entry) => entry.progress === 'done').length;
  const percent = total > 0 ? Math.round((doneCount / total) * 100) : 0;
  const barWidth = `${Math.max(4, percent)}%` as `${number}%`;
  const currentLabel = sections[activeIndex]?.label || '';

  useEffect(() => {
    scrollRef.current?.scrollTo({
      x: Math.max(0, activeIndex * (CHIP_WIDTH + spacing.sm) - spacing.pageX),
      animated: true
    });
  }, [activeIndex]);

  return (
    <Card style={styles.card} padded={false}>
      <View style={styles.progressBlock}>
        <View style={styles.progressTop}>
          <Text style={styles.stepCounter}>
            Schritt {current} von {total}
          </Text>
          <Text style={styles.currentSection} numberOfLines={1}>
            {currentLabel}
          </Text>
        </View>
        <View style={styles.track}>
          <View style={[styles.fill, { width: barWidth }]} />
        </View>
        <View style={styles.metaRow}>
          <Text style={styles.meta}>{percent}% abgeschlossen</Text>
          {totalMissingRequired > 0 ? (
            <Text style={styles.metaMissing}>
              {totalMissingRequired} Pflichtfeld{totalMissingRequired === 1 ? '' : 'er'} offen
            </Text>
          ) : (
            <Text style={styles.metaReady}>Bereit zum Export</Text>
          )}
        </View>
      </View>

      <ScrollView
        ref={scrollRef}
        horizontal
        showsHorizontalScrollIndicator={false}
        contentContainerStyle={styles.navRow}
      >
        {sections.map((entry, index) => {
          const active = index === activeIndex;
          const done = entry.progress === 'done';
          const inProgress = entry.progress === 'progress';
          const missing = entry.missingCount || 0;

          return (
            <Pressable
              key={entry.sectionId}
              style={[
                styles.chip,
                active ? styles.chipActive : null,
                done && !active ? styles.chipDone : null,
                inProgress && !active ? styles.chipInProgress : null
              ]}
              onPress={() => onSelect(index)}
            >
              <View style={styles.chipTop}>
                <View
                  style={[
                    styles.stepBadge,
                    active ? styles.stepBadgeActive : null,
                    done && !active ? styles.stepBadgeDone : null
                  ]}
                >
                  <Text
                    style={[
                      styles.stepBadgeText,
                      active ? styles.stepBadgeTextActive : null,
                      done && !active ? styles.stepBadgeTextDone : null
                    ]}
                  >
                    {done && !active ? '✓' : index + 1}
                  </Text>
                </View>
                {missing > 0 ? (
                  <View style={styles.missingBadge}>
                    <Text style={styles.missingBadgeText}>{missing}</Text>
                  </View>
                ) : null}
              </View>
              <Text
                style={[styles.chipLabel, active ? styles.chipLabelActive : null]}
                numberOfLines={1}
                ellipsizeMode="tail"
              >
                {entry.label}
              </Text>
            </Pressable>
          );
        })}
      </ScrollView>
    </Card>
  );
}

const styles = StyleSheet.create({
  card: {
    overflow: 'hidden'
  },
  progressBlock: {
    gap: spacing.xs,
    paddingHorizontal: spacing.cardPadding,
    paddingTop: spacing.cardPadding,
    paddingBottom: spacing.sm,
    borderBottomWidth: 1,
    borderBottomColor: colors.border
  },
  progressTop: {
    gap: 2
  },
  stepCounter: {
    ...typography.label,
    color: colors.muted
  },
  currentSection: {
    ...typography.subtitle,
    color: colors.ink
  },
  track: {
    height: 8,
    borderRadius: 999,
    backgroundColor: colors.border,
    overflow: 'hidden',
    marginTop: spacing.xxs
  },
  fill: {
    height: '100%',
    borderRadius: 999,
    backgroundColor: colors.accent
  },
  metaRow: {
    flexDirection: 'row',
    justifyContent: 'space-between',
    alignItems: 'center',
    gap: spacing.sm
  },
  meta: {
    ...typography.caption,
    color: colors.muted
  },
  metaMissing: {
    ...typography.caption,
    color: colors.danger,
    fontFamily: 'SpaceGrotesk_600SemiBold'
  },
  metaReady: {
    ...typography.caption,
    color: colors.success,
    fontFamily: 'SpaceGrotesk_600SemiBold'
  },
  navRow: {
    gap: spacing.sm,
    paddingHorizontal: spacing.cardPadding,
    paddingVertical: spacing.sm
  },
  chip: {
    width: CHIP_WIDTH,
    minHeight: 52,
    borderRadius: spacing.inputRadius,
    borderWidth: 1,
    borderColor: colors.border,
    backgroundColor: colors.panel,
    paddingHorizontal: spacing.sm,
    paddingVertical: spacing.sm,
    gap: spacing.xs
  },
  chipActive: {
    borderColor: colors.accent,
    backgroundColor: colors.badgeBg
  },
  chipDone: {
    borderColor: colors.success,
    backgroundColor: 'rgba(47, 107, 69, 0.08)'
  },
  chipInProgress: {
    borderColor: colors.warning
  },
  chipTop: {
    flexDirection: 'row',
    alignItems: 'center',
    justifyContent: 'space-between'
  },
  stepBadge: {
    width: 24,
    height: 24,
    borderRadius: 12,
    borderWidth: 1,
    borderColor: colors.borderStrong,
    backgroundColor: colors.panelElevated,
    alignItems: 'center',
    justifyContent: 'center'
  },
  stepBadgeActive: {
    borderColor: colors.accent,
    backgroundColor: colors.accent
  },
  stepBadgeDone: {
    borderColor: colors.success,
    backgroundColor: colors.success
  },
  stepBadgeText: {
    ...typography.caption,
    color: colors.muted,
    fontFamily: 'SpaceGrotesk_600SemiBold',
    fontSize: 11
  },
  stepBadgeTextActive: {
    color: colors.white
  },
  stepBadgeTextDone: {
    color: colors.white
  },
  missingBadge: {
    minWidth: 20,
    height: 20,
    borderRadius: 10,
    paddingHorizontal: 6,
    backgroundColor: colors.danger,
    alignItems: 'center',
    justifyContent: 'center'
  },
  missingBadgeText: {
    ...typography.caption,
    color: colors.white,
    fontFamily: 'SpaceGrotesk_600SemiBold',
    fontSize: 11
  },
  chipLabel: {
    ...typography.caption,
    color: colors.ink,
    fontFamily: 'SpaceGrotesk_400Regular'
  },
  chipLabelActive: {
    color: colors.accent,
    fontFamily: 'SpaceGrotesk_600SemiBold'
  }
});

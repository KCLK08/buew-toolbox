import { Pressable, StyleSheet, Text, View } from 'react-native';
import { MaterialCommunityIcons } from '@expo/vector-icons';

import { colors, shadows, spacing, typography } from '../../../../constants/theme';
import { hapticSelection } from '../../../../lib/haptics';

export type SetupStructureViewTab = 'pdf' | 'structure';

type Props = {
  activeTab: SetupStructureViewTab;
  templateName?: string;
  onTabChange: (tab: SetupStructureViewTab) => void;
  onBack: () => void;
};

export function SetupStructureHeader({ activeTab, templateName, onTabChange, onBack }: Props) {
  const selectTab = (tab: SetupStructureViewTab) => {
    if (tab === activeTab) return;
    void hapticSelection();
    onTabChange(tab);
  };

  return (
    <View style={styles.root}>
      <View style={styles.topRow}>
        <Pressable accessibilityRole="button" style={styles.backBtn} onPress={onBack}>
          <MaterialCommunityIcons name="chevron-left" size={22} color={colors.accent} />
          <Text style={styles.backLabel}>Zurück</Text>
        </Pressable>
        <View style={styles.stepBadge}>
          <MaterialCommunityIcons name="numeric-1-circle" size={16} color={colors.accent} />
          <Text style={styles.step}>Schritt 1 von 3</Text>
        </View>
      </View>

      <View style={styles.titleBlock}>
        <Text style={styles.title}>Setup</Text>
        {templateName ? (
          <Text style={styles.templateName} numberOfLines={1}>
            {templateName}
          </Text>
        ) : null}
        <Text style={styles.subtitle}>Aus welchen Bereichen besteht dein digitales Bautagebuch?</Text>
      </View>

      <View style={styles.tabs}>
        <Pressable
          accessibilityRole="tab"
          accessibilityState={{ selected: activeTab === 'pdf' }}
          style={[styles.tab, activeTab === 'pdf' ? styles.tabActive : null]}
          onPress={() => selectTab('pdf')}
        >
          <MaterialCommunityIcons
            name="file-pdf-box"
            size={18}
            color={activeTab === 'pdf' ? colors.accent : colors.muted}
          />
          <Text style={activeTab === 'pdf' ? styles.tabLabelActive : styles.tabLabel}>PDF</Text>
        </Pressable>
        <Pressable
          accessibilityRole="tab"
          accessibilityState={{ selected: activeTab === 'structure' }}
          style={[styles.tab, activeTab === 'structure' ? styles.tabActive : null]}
          onPress={() => selectTab('structure')}
        >
          <MaterialCommunityIcons
            name="view-list"
            size={18}
            color={activeTab === 'structure' ? colors.accent : colors.muted}
          />
          <Text style={activeTab === 'structure' ? styles.tabLabelActive : styles.tabLabel}>Struktur</Text>
        </Pressable>
      </View>
    </View>
  );
}

const styles = StyleSheet.create({
  root: {
    backgroundColor: colors.panel,
    borderBottomWidth: 1,
    borderBottomColor: colors.border,
    paddingHorizontal: spacing.pageX,
    paddingTop: spacing.xs,
    paddingBottom: spacing.sm,
    gap: spacing.sm
  },
  topRow: {
    flexDirection: 'row',
    alignItems: 'center',
    justifyContent: 'space-between',
    minHeight: spacing.touchMin
  },
  backBtn: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: 2,
    minHeight: spacing.touchMin,
    marginLeft: -4
  },
  backLabel: {
    ...typography.bodyStrong,
    color: colors.accent
  },
  stepBadge: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: spacing.xxs,
    paddingHorizontal: spacing.sm,
    paddingVertical: spacing.xxs,
    borderRadius: 999,
    backgroundColor: colors.badgeBg
  },
  step: {
    ...typography.caption,
    color: colors.accent,
    fontFamily: 'SpaceGrotesk_600SemiBold'
  },
  titleBlock: {
    gap: 4
  },
  title: {
    ...typography.title,
    color: colors.ink
  },
  templateName: {
    ...typography.bodyStrong,
    color: colors.accent2
  },
  subtitle: {
    ...typography.body,
    color: colors.muted,
    lineHeight: 22
  },
  tabs: {
    flexDirection: 'row',
    gap: spacing.xs,
    padding: 4,
    borderRadius: 14,
    backgroundColor: colors.bg,
    borderWidth: 1,
    borderColor: colors.border
  },
  tab: {
    flex: 1,
    flexDirection: 'row',
    alignItems: 'center',
    justifyContent: 'center',
    gap: spacing.xxs,
    minHeight: spacing.touchMin,
    borderRadius: 10
  },
  tabActive: {
    backgroundColor: colors.panelElevated,
    ...shadows.card
  },
  tabLabel: {
    ...typography.bodyStrong,
    color: colors.muted
  },
  tabLabelActive: {
    ...typography.bodyStrong,
    color: colors.accent2
  }
});

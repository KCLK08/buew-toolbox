import { Pressable, StyleSheet, Text, View } from 'react-native';

import { colors, spacing, typography } from '../../../../constants/theme';

export type SetupStructureViewTab = 'pdf' | 'structure';

type Props = {
  activeTab: SetupStructureViewTab;
  onTabChange: (tab: SetupStructureViewTab) => void;
  onBack: () => void;
};

export function SetupStructureHeader({ activeTab, onTabChange, onBack }: Props) {
  return (
    <View style={styles.root}>
      <Pressable accessibilityRole="button" style={styles.backBtn} onPress={onBack}>
        <Text style={styles.backLabel}>← Zurück</Text>
      </Pressable>

      <View style={styles.titleBlock}>
        <Text style={styles.title}>Setup</Text>
        <Text style={styles.step}>Schritt 1 von 3</Text>
      </View>

      <View style={styles.tabs}>
        <Pressable
          accessibilityRole="tab"
          accessibilityState={{ selected: activeTab === 'pdf' }}
          style={[styles.tab, activeTab === 'pdf' ? styles.tabActive : null]}
          onPress={() => onTabChange('pdf')}
        >
          <Text style={activeTab === 'pdf' ? styles.tabLabelActive : styles.tabLabel}>PDF</Text>
        </Pressable>
        <Pressable
          accessibilityRole="tab"
          accessibilityState={{ selected: activeTab === 'structure' }}
          style={[styles.tab, activeTab === 'structure' ? styles.tabActive : null]}
          onPress={() => onTabChange('structure')}
        >
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
    paddingTop: spacing.sm,
    paddingBottom: spacing.sm,
    gap: spacing.sm
  },
  backBtn: {
    minHeight: spacing.touchMin,
    justifyContent: 'center',
    alignSelf: 'flex-start'
  },
  backLabel: {
    ...typography.bodyStrong,
    color: colors.accent
  },
  titleBlock: {
    gap: 2
  },
  title: {
    ...typography.title,
    color: colors.ink
  },
  step: {
    ...typography.caption,
    color: colors.accent,
    fontFamily: 'SpaceGrotesk_600SemiBold'
  },
  tabs: {
    flexDirection: 'row',
    gap: spacing.xs,
    padding: 4,
    borderRadius: 999,
    backgroundColor: colors.panelElevated,
    borderWidth: 1,
    borderColor: colors.border
  },
  tab: {
    flex: 1,
    minHeight: spacing.touchMin,
    borderRadius: 999,
    alignItems: 'center',
    justifyContent: 'center'
  },
  tabActive: {
    backgroundColor: colors.badgeBg,
    borderWidth: 1,
    borderColor: colors.accent
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

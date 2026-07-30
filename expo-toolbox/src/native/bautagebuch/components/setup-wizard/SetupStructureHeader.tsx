import { Alert, Pressable, StyleSheet, Text, View } from 'react-native';
import { MaterialCommunityIcons } from '@expo/vector-icons';
import { useSafeAreaInsets } from 'react-native-safe-area-context';

import { colors, shadows, spacing, typography } from '../../../../constants/theme';
import { hapticSelection } from '../../../../lib/haptics';
import { systemTopInset } from '../../../../navigation/systemInsets';

export type SetupStructureViewTab = 'pdf' | 'structure';

const INFO_TEXT =
  'Definiere hier die Bereiche deines digitalen Bautagebuchs anhand der PDF-Vorlage. Lege Gruppen und Tabellen an — die Feldzuordnung folgt in Schritt 2.';

type Props = {
  activeTab: SetupStructureViewTab;
  onTabChange: (tab: SetupStructureViewTab) => void;
  onBack: () => void;
  applyTopInset?: boolean;
};

export function SetupStructureHeader({
  activeTab,
  onTabChange,
  onBack,
  applyTopInset = true
}: Props) {
  const insets = useSafeAreaInsets();
  const topInset = applyTopInset ? systemTopInset(insets) : 0;
  const selectTab = (tab: SetupStructureViewTab) => {
    if (tab === activeTab) return;
    void hapticSelection();
    onTabChange(tab);
  };

  const showInfo = () => {
    Alert.alert('Setup Schritt 1', INFO_TEXT);
  };

  return (
    <View style={[styles.root, { paddingTop: topInset + 2 }]}>
      <View style={styles.topRow}>
        <Pressable accessibilityRole="button" style={styles.backBtn} onPress={onBack}>
          <MaterialCommunityIcons name="chevron-left" size={20} color={colors.accent} />
          <Text style={styles.backLabel}>Zurück</Text>
        </Pressable>

        <View style={styles.titleRow}>
          <Text style={styles.title}>Setup</Text>
          <Text style={styles.dot}>·</Text>
          <Text style={styles.step}>Schritt 1 von 3</Text>
          <Pressable
            accessibilityRole="button"
            accessibilityLabel="Informationen zum Setup"
            style={styles.infoBtn}
            onPress={showInfo}
            hitSlop={8}
          >
            <MaterialCommunityIcons name="information-outline" size={18} color={colors.muted} />
          </Pressable>
        </View>

        <View style={styles.topSpacer} />
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
            size={16}
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
            size={16}
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
    paddingBottom: spacing.xs,
    gap: spacing.xs
  },
  topRow: {
    flexDirection: 'row',
    alignItems: 'center',
    minHeight: 40
  },
  backBtn: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: 2,
    minWidth: 80,
    marginLeft: -4
  },
  backLabel: {
    ...typography.caption,
    color: colors.accent,
    fontFamily: 'SpaceGrotesk_600SemiBold'
  },
  titleRow: {
    flex: 1,
    flexDirection: 'row',
    alignItems: 'center',
    justifyContent: 'center',
    gap: spacing.xxs
  },
  title: {
    ...typography.bodyStrong,
    color: colors.ink
  },
  dot: {
    ...typography.bodyStrong,
    color: colors.muted
  },
  step: {
    ...typography.caption,
    color: colors.muted,
    fontFamily: 'SpaceGrotesk_600SemiBold'
  },
  infoBtn: {
    width: 28,
    height: 28,
    alignItems: 'center',
    justifyContent: 'center',
    marginLeft: 2
  },
  topSpacer: {
    minWidth: 80
  },
  tabs: {
    flexDirection: 'row',
    gap: 4,
    padding: 3,
    borderRadius: 12,
    backgroundColor: colors.bg,
    borderWidth: 1,
    borderColor: colors.border
  },
  tab: {
    flex: 1,
    flexDirection: 'row',
    alignItems: 'center',
    justifyContent: 'center',
    gap: 4,
    minHeight: 36,
    borderRadius: 9
  },
  tabActive: {
    backgroundColor: colors.panelElevated,
    ...shadows.card
  },
  tabLabel: {
    ...typography.caption,
    color: colors.muted,
    fontFamily: 'SpaceGrotesk_600SemiBold'
  },
  tabLabelActive: {
    ...typography.caption,
    color: colors.accent2,
    fontFamily: 'SpaceGrotesk_600SemiBold'
  }
});

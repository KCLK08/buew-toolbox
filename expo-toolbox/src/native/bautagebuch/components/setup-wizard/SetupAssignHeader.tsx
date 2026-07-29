import { Alert, Pressable, StyleSheet, Text, View } from 'react-native';
import { MaterialCommunityIcons } from '@expo/vector-icons';

import { colors, shadows, spacing, typography } from '../../../../constants/theme';
import { hapticSelection } from '../../../../lib/haptics';

export type SetupAssignViewTab = 'pdf' | 'assign';

const INFO_TEXT =
  'Ordne jedes erkannte PDF-Feld einer Gruppe oder Tabelle zu. Die Struktur hast du in Schritt 1 festgelegt — hier entstehen die digitalen Felder.';

type Props = {
  activeTab: SetupAssignViewTab;
  onTabChange: (tab: SetupAssignViewTab) => void;
  onBack: () => void;
  onOpenFields?: () => void;
};

export function SetupAssignHeader({ activeTab, onTabChange, onBack, onOpenFields }: Props) {
  const selectTab = (tab: SetupAssignViewTab) => {
    if (tab === activeTab) return;
    void hapticSelection();
    onTabChange(tab);
  };

  const showInfo = () => {
    Alert.alert('Setup Schritt 2', INFO_TEXT);
  };

  return (
    <View style={styles.root}>
      <View style={styles.topRow}>
        <Pressable accessibilityRole="button" style={styles.backBtn} onPress={onBack}>
          <MaterialCommunityIcons name="chevron-left" size={20} color={colors.accent} />
          <Text style={styles.backLabel}>Setup</Text>
        </Pressable>

        <View style={styles.titleRow}>
          <Text style={styles.step}>Schritt 2 von 3</Text>
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

        {onOpenFields ? (
          <Pressable accessibilityRole="button" style={styles.fieldsBtn} onPress={onOpenFields}>
            <Text style={styles.fieldsLabel}>Felder</Text>
          </Pressable>
        ) : (
          <View style={styles.topSpacer} />
        )}
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
          accessibilityState={{ selected: activeTab === 'assign' }}
          style={[styles.tab, activeTab === 'assign' ? styles.tabActive : null]}
          onPress={() => selectTab('assign')}
        >
          <MaterialCommunityIcons
            name="shape-outline"
            size={16}
            color={activeTab === 'assign' ? colors.accent : colors.muted}
          />
          <Text style={activeTab === 'assign' ? styles.tabLabelActive : styles.tabLabel}>Zuordnung</Text>
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
    paddingTop: 2,
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
    minWidth: 72,
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
  step: {
    ...typography.bodyStrong,
    color: colors.ink
  },
  infoBtn: {
    width: 28,
    height: 28,
    alignItems: 'center',
    justifyContent: 'center'
  },
  fieldsBtn: {
    minWidth: 72,
    minHeight: 36,
    alignItems: 'flex-end',
    justifyContent: 'center',
    paddingHorizontal: spacing.xxs
  },
  fieldsLabel: {
    ...typography.caption,
    color: colors.accent,
    fontFamily: 'SpaceGrotesk_600SemiBold'
  },
  topSpacer: {
    minWidth: 72
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

import { Pressable, StyleSheet, Text, View } from 'react-native';
import { MaterialCommunityIcons } from '@expo/vector-icons';
import { useSafeAreaInsets } from 'react-native-safe-area-context';

import { colors, spacing, typography } from '../../../../constants/theme';
import { hapticSelection } from '../../../../lib/haptics';
import { systemTopInset } from '../../../../navigation/systemInsets';
import type { SetupWizardStep } from '../../types';

type StepDef = {
  step: SetupWizardStep;
  number: number;
  label: string;
  shortLabel: string;
};

const STEPS: StepDef[] = [
  { step: 'structure', number: 1, label: 'Struktur', shortLabel: 'Struktur' },
  { step: 'assign', number: 2, label: 'Zuordnung', shortLabel: 'Zuordnung' },
  { step: 'fields', number: 3, label: 'Einstellungen', shortLabel: 'Einstellungen' }
];

type Props = {
  activeStep: SetupWizardStep;
  onSelectStep: (step: SetupWizardStep) => void;
};

export function SetupWizardStepNav({ activeStep, onSelectStep }: Props) {
  const insets = useSafeAreaInsets();
  const topInset = systemTopInset(insets);

  return (
    <View style={[styles.root, { paddingTop: topInset }]}>
      {STEPS.map((entry, index) => {
        const active = entry.step === activeStep;
        return (
          <View key={entry.step} style={styles.itemWrap}>
            {index > 0 ? <View style={styles.divider} /> : null}
            <Pressable
              accessibilityRole="tab"
              accessibilityState={{ selected: active }}
              style={[styles.item, active ? styles.itemActive : null]}
              onPress={() => {
                if (active) return;
                void hapticSelection();
                onSelectStep(entry.step);
              }}
            >
              <View style={[styles.badge, active ? styles.badgeActive : null]}>
                <Text style={active ? styles.badgeTextActive : styles.badgeText}>{entry.number}</Text>
              </View>
              <Text style={active ? styles.labelActive : styles.label}>{entry.shortLabel}</Text>
              {active ? (
                <MaterialCommunityIcons name="chevron-down" size={14} color={colors.accent} />
              ) : null}
            </Pressable>
          </View>
        );
      })}
    </View>
  );
}

const styles = StyleSheet.create({
  root: {
    flexDirection: 'row',
    alignItems: 'center',
    justifyContent: 'space-between',
    paddingHorizontal: spacing.pageX,
    paddingVertical: spacing.sm,
    backgroundColor: colors.panel,
    borderBottomWidth: 1,
    borderBottomColor: colors.border
  },
  itemWrap: {
    flex: 1,
    flexDirection: 'row',
    alignItems: 'center'
  },
  divider: {
    width: 1,
    height: 28,
    backgroundColor: colors.border,
    marginRight: spacing.xxs
  },
  item: {
    flex: 1,
    flexDirection: 'row',
    alignItems: 'center',
    justifyContent: 'center',
    gap: spacing.xxs,
    minHeight: spacing.touchMin,
    paddingHorizontal: spacing.xxs,
    borderRadius: 10
  },
  itemActive: {
    backgroundColor: colors.badgeBg
  },
  badge: {
    width: 22,
    height: 22,
    borderRadius: 999,
    alignItems: 'center',
    justifyContent: 'center',
    backgroundColor: colors.border
  },
  badgeActive: {
    backgroundColor: colors.accent
  },
  badgeText: {
    ...typography.caption,
    color: colors.muted,
    fontFamily: 'SpaceGrotesk_600SemiBold',
    fontSize: 11
  },
  badgeTextActive: {
    ...typography.caption,
    color: colors.white,
    fontFamily: 'SpaceGrotesk_600SemiBold',
    fontSize: 11
  },
  label: {
    ...typography.caption,
    color: colors.muted
  },
  labelActive: {
    ...typography.caption,
    color: colors.accent2,
    fontFamily: 'SpaceGrotesk_600SemiBold'
  }
});

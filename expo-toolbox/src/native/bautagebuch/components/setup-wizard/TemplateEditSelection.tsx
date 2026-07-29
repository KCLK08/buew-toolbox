import { Pressable, ScrollView, StyleSheet, Text, View } from 'react-native';
import { MaterialCommunityIcons } from '@expo/vector-icons';
import { useSafeAreaInsets } from 'react-native-safe-area-context';

import { colors, shadows, spacing, typography } from '../../../../constants/theme';
import { hapticSelection } from '../../../../lib/haptics';
import { systemBottomInset } from '../../../../navigation/systemInsets';
import type { SetupWizardStep } from '../../types';

type EditOption = {
  step: SetupWizardStep;
  number: number;
  icon: keyof typeof MaterialCommunityIcons.glyphMap;
  title: string;
  subtitle: string;
};

const OPTIONS: EditOption[] = [
  {
    step: 'structure',
    number: 1,
    icon: 'folder-outline',
    title: 'Struktur',
    subtitle: 'Gruppen und Tabellen\nverwalten'
  },
  {
    step: 'assign',
    number: 2,
    icon: 'link-variant',
    title: 'Feldzuordnung',
    subtitle: 'PDF-Felder Gruppen\nund Tabellen zuordnen'
  },
  {
    step: 'fields',
    number: 3,
    icon: 'tune-variant',
    title: 'Feldeinstellungen',
    subtitle: 'Verhalten und Eigenschaften\nder Felder ändern'
  }
];

type Props = {
  templateName: string;
  onBack: () => void;
  onSelectStep: (step: SetupWizardStep) => void;
};

export function TemplateEditSelection({ templateName, onBack, onSelectStep }: Props) {
  const insets = useSafeAreaInsets();

  return (
    <View style={[styles.root, { paddingTop: insets.top }]}>
      <View style={styles.header}>
        <Pressable accessibilityRole="button" style={styles.backBtn} onPress={onBack}>
          <MaterialCommunityIcons name="chevron-left" size={22} color={colors.accent} />
          <Text style={styles.backLabel}>Vorlage</Text>
        </Pressable>
      </View>

      <ScrollView
        contentContainerStyle={[
          styles.content,
          { paddingBottom: systemBottomInset(insets) + spacing.xl }
        ]}
        showsVerticalScrollIndicator={false}
      >
        <Text style={styles.title}>Was möchtest du ändern?</Text>
        <Text style={styles.subtitle}>{templateName}</Text>

        <View style={styles.list}>
          {OPTIONS.map((option) => (
            <Pressable
              key={option.step}
              accessibilityRole="button"
              style={[styles.card, shadows.card]}
              onPress={() => {
                void hapticSelection();
                onSelectStep(option.step);
              }}
            >
              <View style={styles.numberBadge}>
                <Text style={styles.numberText}>{option.number}</Text>
              </View>
              <View style={styles.iconWrap}>
                <MaterialCommunityIcons name={option.icon} size={26} color={colors.accent2} />
              </View>
              <View style={styles.cardCopy}>
                <Text style={styles.cardTitle}>{option.title}</Text>
                <Text style={styles.cardSubtitle}>{option.subtitle}</Text>
              </View>
              <MaterialCommunityIcons name="chevron-right" size={22} color={colors.muted} />
            </Pressable>
          ))}
        </View>
      </ScrollView>
    </View>
  );
}

const styles = StyleSheet.create({
  root: {
    flex: 1,
    backgroundColor: colors.bg
  },
  header: {
    paddingHorizontal: spacing.pageX,
    paddingBottom: spacing.sm
  },
  backBtn: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: spacing.xxs,
    minHeight: spacing.touchMin,
    alignSelf: 'flex-start'
  },
  backLabel: {
    ...typography.bodyStrong,
    color: colors.accent
  },
  content: {
    paddingHorizontal: spacing.pageX,
    gap: spacing.md
  },
  title: {
    ...typography.title,
    color: colors.ink,
    fontSize: 26
  },
  subtitle: {
    ...typography.body,
    color: colors.muted,
    marginTop: -spacing.xs
  },
  list: {
    gap: spacing.md,
    marginTop: spacing.sm
  },
  card: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: spacing.sm,
    padding: spacing.md,
    borderRadius: spacing.cardRadius,
    backgroundColor: colors.panelElevated,
    borderWidth: 1,
    borderColor: colors.border,
    minHeight: 88
  },
  numberBadge: {
    position: 'absolute',
    top: spacing.sm,
    left: spacing.sm,
    width: 24,
    height: 24,
    borderRadius: 999,
    alignItems: 'center',
    justifyContent: 'center',
    backgroundColor: colors.accent
  },
  numberText: {
    ...typography.caption,
    color: colors.white,
    fontFamily: 'SpaceGrotesk_600SemiBold'
  },
  iconWrap: {
    width: 48,
    height: 48,
    marginLeft: spacing.lg,
    borderRadius: spacing.iconRadius,
    alignItems: 'center',
    justifyContent: 'center',
    backgroundColor: colors.badgeBg
  },
  cardCopy: {
    flex: 1,
    gap: spacing.xxs
  },
  cardTitle: {
    ...typography.subtitle,
    color: colors.ink
  },
  cardSubtitle: {
    ...typography.caption,
    color: colors.muted,
    lineHeight: 18
  }
});

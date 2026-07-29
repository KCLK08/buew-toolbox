import { useEffect, useRef } from 'react';
import { Animated, Pressable, ScrollView, StyleSheet, Text, View } from 'react-native';
import { MaterialCommunityIcons } from '@expo/vector-icons';
import { useSafeAreaInsets } from 'react-native-safe-area-context';

import { PrimaryButton } from '../../../../components/mobile';
import { colors, shadows, spacing, typography } from '../../../../constants/theme';
import { systemBottomInset } from '../../../../navigation/systemInsets';

type InfoCard = {
  icon: keyof typeof MaterialCommunityIcons.glyphMap;
  tint: string;
  bg: string;
  title: string;
  copy: string;
};

const INFO_CARDS: InfoCard[] = [
  {
    icon: 'clipboard-check-outline',
    tint: colors.accent,
    bg: colors.badgeBg,
    title: 'Feld prüfen',
    copy: 'Gehe alle erstellten Felder nacheinander durch und überprüfe die vorgeschlagenen Einstellungen.'
  },
  {
    icon: 'tune-variant',
    tint: colors.accent2,
    bg: 'rgba(36, 50, 64, 0.1)',
    title: 'Verhalten festlegen',
    copy: 'Lege fest, wie sich jedes Feld später verhält, z. B. Datentyp, Pflichtfeld oder Auswahlmöglichkeiten.'
  },
  {
    icon: 'shape-outline',
    tint: colors.info,
    bg: 'rgba(42, 95, 143, 0.12)',
    title: 'Individuelle Anpassung',
    copy: 'Jeder Feldtyp besitzt eigene Einstellungsmöglichkeiten. Es werden immer nur die Optionen angezeigt, die für das jeweilige Feld sinnvoll sind.'
  }
];

type Props = {
  onBack: () => void;
  onStart: () => void;
};

export function SetupFieldSettingsOnboarding({ onBack, onStart }: Props) {
  const insets = useSafeAreaInsets();
  const fade = useRef(new Animated.Value(0)).current;
  const slide = useRef(new Animated.Value(16)).current;
  const gearSpin = useRef(new Animated.Value(0)).current;

  useEffect(() => {
    Animated.parallel([
      Animated.timing(fade, {
        toValue: 1,
        duration: 420,
        useNativeDriver: true
      }),
      Animated.timing(slide, {
        toValue: 0,
        duration: 420,
        useNativeDriver: true
      })
    ]).start();

    Animated.loop(
      Animated.timing(gearSpin, {
        toValue: 1,
        duration: 4200,
        useNativeDriver: true
      })
    ).start();
  }, [fade, slide, gearSpin]);

  const gearRotate = gearSpin.interpolate({
    inputRange: [0, 1],
    outputRange: ['0deg', '360deg']
  });

  return (
    <View style={styles.root}>
      <View style={[styles.header, { paddingTop: spacing.xs }]}>
        <Pressable accessibilityRole="button" style={styles.backBtn} onPress={onBack}>
          <MaterialCommunityIcons name="chevron-left" size={20} color={colors.accent} />
          <Text style={styles.backLabel}>Zurück</Text>
        </Pressable>
        <View style={styles.headerCopy}>
          <Text style={styles.step}>Schritt 3 von 3</Text>
          <Text style={styles.subtitle}>Feldeinstellungen</Text>
        </View>
        <View style={styles.headerSpacer} />
      </View>

      <ScrollView
        style={styles.scroll}
        contentContainerStyle={[
          styles.content,
          { paddingBottom: systemBottomInset(insets) + spacing.touchMin + 24 }
        ]}
        showsVerticalScrollIndicator={false}
      >
        <Animated.View style={[styles.hero, { opacity: fade, transform: [{ translateY: slide }] }]}>
          <View style={styles.illustration}>
            <View style={styles.fieldPanel}>
              <View style={styles.fieldLabelLine} />
              <View style={styles.fieldInput}>
                <MaterialCommunityIcons name="form-textbox" size={18} color={colors.accent2} />
                <View style={styles.fieldInputLine} />
              </View>
              <View style={styles.fieldLabelLineShort} />
            </View>

            <Animated.View style={[styles.gearBadge, { transform: [{ rotate: gearRotate }] }]}>
              <MaterialCommunityIcons name="cog" size={28} color={colors.accent} />
            </Animated.View>
          </View>

          <Text style={styles.headline}>Passe das Verhalten deiner Felder an</Text>
          <Text style={styles.description}>
            Alle PDF-Felder wurden bereits den passenden Gruppen und Tabellen zugeordnet.
          </Text>
          <Text style={styles.description}>
            Jetzt legst du fest, wie sich jedes Feld im digitalen Bautagebuch verhalten soll.
          </Text>
          <Text style={styles.descriptionMuted}>
            Je nach Feldtyp stehen unterschiedliche Einstellungen zur Verfügung.
          </Text>
        </Animated.View>

        <Animated.View style={[styles.cards, { opacity: fade, transform: [{ translateY: slide }] }]}>
          {INFO_CARDS.map((card) => (
            <View key={card.title} style={styles.infoCard}>
              <View style={[styles.infoIconWrap, { backgroundColor: card.bg }]}>
                <MaterialCommunityIcons name={card.icon} size={22} color={card.tint} />
              </View>
              <View style={styles.infoCopy}>
                <Text style={styles.infoTitle}>{card.title}</Text>
                <Text style={styles.infoText}>{card.copy}</Text>
              </View>
            </View>
          ))}
        </Animated.View>

        <Animated.View style={{ opacity: fade, transform: [{ translateY: slide }] }}>
          <View style={styles.tipBox}>
            <View style={styles.tipHeader}>
              <MaterialCommunityIcons name="lightbulb-on-outline" size={18} color={colors.accent} />
              <Text style={styles.tipTitle}>Tipp</Text>
            </View>
            <Text style={styles.tipText}>
              Du kannst jederzeit zwischen der PDF-Ansicht und der Einstellungsansicht wechseln. Das
              aktuell bearbeitete Feld bleibt dabei markiert, sodass du die Position im
              Originalformular jederzeit nachvollziehen kannst.
            </Text>
          </View>
        </Animated.View>
      </ScrollView>

      <View style={[styles.footer, { paddingBottom: systemBottomInset(insets) + spacing.sm }]}>
        <PrimaryButton label="Feldeinstellungen starten" onPress={onStart} />
      </View>
    </View>
  );
}

const styles = StyleSheet.create({
  root: {
    flex: 1,
    backgroundColor: colors.bg
  },
  header: {
    flexDirection: 'row',
    alignItems: 'flex-start',
    paddingHorizontal: spacing.pageX,
    paddingBottom: spacing.sm,
    borderBottomWidth: 1,
    borderBottomColor: colors.border,
    backgroundColor: colors.panel
  },
  backBtn: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: 2,
    minWidth: 80,
    minHeight: spacing.touchMin,
    marginLeft: -4
  },
  backLabel: {
    ...typography.caption,
    color: colors.accent,
    fontFamily: 'SpaceGrotesk_600SemiBold'
  },
  headerCopy: {
    flex: 1,
    alignItems: 'center',
    gap: 2,
    paddingTop: 2
  },
  headerSpacer: {
    minWidth: 80
  },
  step: {
    ...typography.caption,
    color: colors.accent,
    fontFamily: 'SpaceGrotesk_600SemiBold'
  },
  subtitle: {
    ...typography.bodyStrong,
    color: colors.ink,
    textAlign: 'center'
  },
  scroll: {
    flex: 1
  },
  content: {
    padding: spacing.pageX,
    gap: spacing.lg
  },
  hero: {
    alignItems: 'center',
    gap: spacing.sm
  },
  illustration: {
    width: '100%',
    maxWidth: 280,
    alignItems: 'center',
    justifyContent: 'center',
    marginBottom: spacing.xs,
    paddingVertical: spacing.md,
    position: 'relative'
  },
  fieldPanel: {
    width: 220,
    minHeight: 132,
    borderRadius: 16,
    borderWidth: 1,
    borderColor: colors.border,
    backgroundColor: colors.panelElevated,
    padding: spacing.md,
    gap: spacing.sm,
    ...shadows.card
  },
  fieldLabelLine: {
    height: 8,
    width: '55%',
    borderRadius: 999,
    backgroundColor: colors.border
  },
  fieldLabelLineShort: {
    height: 8,
    width: '40%',
    borderRadius: 999,
    backgroundColor: colors.border
  },
  fieldInput: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: spacing.xs,
    minHeight: 44,
    borderRadius: 10,
    borderWidth: 2,
    borderColor: colors.accent2,
    backgroundColor: colors.panel,
    paddingHorizontal: spacing.sm
  },
  fieldInputLine: {
    flex: 1,
    height: 6,
    borderRadius: 999,
    backgroundColor: colors.border
  },
  gearBadge: {
    position: 'absolute',
    right: 24,
    bottom: 8,
    width: 56,
    height: 56,
    borderRadius: 28,
    borderWidth: 2,
    borderColor: colors.accent,
    backgroundColor: colors.badgeBg,
    alignItems: 'center',
    justifyContent: 'center',
    ...shadows.card
  },
  headline: {
    ...typography.title,
    color: colors.ink,
    textAlign: 'center'
  },
  description: {
    ...typography.body,
    color: colors.ink,
    textAlign: 'center',
    lineHeight: 22
  },
  descriptionMuted: {
    ...typography.body,
    color: colors.muted,
    textAlign: 'center',
    lineHeight: 22
  },
  cards: {
    gap: spacing.sm
  },
  infoCard: {
    flexDirection: 'row',
    alignItems: 'flex-start',
    gap: spacing.sm,
    padding: spacing.md,
    borderRadius: spacing.cardRadius,
    borderWidth: 1,
    borderColor: colors.border,
    backgroundColor: colors.panelElevated,
    ...shadows.card
  },
  infoIconWrap: {
    width: 44,
    height: 44,
    borderRadius: 12,
    alignItems: 'center',
    justifyContent: 'center'
  },
  infoCopy: {
    flex: 1,
    gap: 2
  },
  infoTitle: {
    ...typography.bodyStrong,
    color: colors.ink
  },
  infoText: {
    ...typography.caption,
    color: colors.muted,
    lineHeight: 18
  },
  tipBox: {
    borderRadius: spacing.cardRadius,
    borderWidth: 1,
    borderColor: colors.accent,
    backgroundColor: colors.badgeBg,
    padding: spacing.md,
    gap: spacing.xs
  },
  tipHeader: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: spacing.xxs
  },
  tipTitle: {
    ...typography.bodyStrong,
    color: colors.accent2
  },
  tipText: {
    ...typography.body,
    color: colors.ink,
    lineHeight: 22
  },
  footer: {
    paddingHorizontal: spacing.pageX,
    paddingTop: spacing.sm,
    borderTopWidth: 1,
    borderTopColor: colors.border,
    backgroundColor: colors.panel
  }
});

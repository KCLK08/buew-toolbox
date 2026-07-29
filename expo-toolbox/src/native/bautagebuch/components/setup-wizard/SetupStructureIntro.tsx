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
    icon: 'file-pdf-box',
    tint: colors.accent,
    bg: colors.badgeBg,
    title: 'PDF ansehen',
    copy: 'Sieh dir die Vorlage an und verschaffe dir einen Überblick über ihren Aufbau.'
  },
  {
    icon: 'view-grid-plus',
    tint: colors.accent2,
    bg: 'rgba(36, 50, 64, 0.1)',
    title: 'Struktur erstellen',
    copy: 'Erstelle alle benötigten Gruppen und Tabellen für dein digitales Bautagebuch.'
  },
  {
    icon: 'link-variant-off',
    tint: colors.info,
    bg: 'rgba(42, 95, 143, 0.12)',
    title: 'Keine Feldzuordnung',
    copy: 'Die Zuordnung einzelner PDF-Felder erfolgt erst im nächsten Schritt.'
  }
];

type Props = {
  onBack: () => void;
  onStart: () => void;
};

export function SetupStructureIntro({ onBack, onStart }: Props) {
  const insets = useSafeAreaInsets();
  const fade = useRef(new Animated.Value(0)).current;
  const slide = useRef(new Animated.Value(16)).current;

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
  }, [fade, slide]);

  return (
    <View style={styles.root}>
      <View style={[styles.header, { paddingTop: spacing.xs }]}>
        <Pressable accessibilityRole="button" style={styles.backBtn} onPress={onBack}>
          <MaterialCommunityIcons name="chevron-left" size={20} color={colors.accent} />
          <Text style={styles.backLabel}>Zurück</Text>
        </Pressable>
        <View style={styles.headerCopy}>
          <Text style={styles.step}>Schritt 1 von 3</Text>
          <Text style={styles.subtitle}>Gruppen & Tabellen definieren</Text>
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
            <View style={[styles.illustrationCard, styles.illustrationCardBack]} />
            <View style={[styles.illustrationCard, styles.illustrationCardMid]} />
            <View style={[styles.illustrationCard, styles.illustrationCardFront]}>
              <MaterialCommunityIcons name="file-document-outline" size={28} color={colors.accent2} />
              <View style={styles.illustrationLines}>
                <View style={styles.illustrationLine} />
                <View style={[styles.illustrationLine, styles.illustrationLineShort]} />
              </View>
            </View>
            <View style={styles.illustrationTable}>
              <MaterialCommunityIcons name="table" size={20} color={colors.info} />
            </View>
          </View>

          <Text style={styles.headline}>Baue die Struktur deines Bautagebuchs auf</Text>
          <Text style={styles.description}>
            In diesem Schritt legst du fest, aus welchen Bereichen dein digitales Bautagebuch bestehen
            soll.
          </Text>
          <Text style={styles.description}>
            Nutze die PDF als Orientierung und erstelle die Gruppen und Tabellen, die später ausgefüllt
            werden sollen.
          </Text>
          <Text style={styles.descriptionMuted}>
            Du musst in diesem Schritt keine Felder auswählen oder zuordnen.
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
              Du kannst jederzeit zwischen der PDF-Ansicht und der Strukturansicht wechseln. Die PDF
              dient in diesem Schritt ausschließlich als Orientierung.
            </Text>
          </View>
        </Animated.View>
      </ScrollView>

      <View style={[styles.footer, { paddingBottom: systemBottomInset(insets) + spacing.sm }]}>
        <PrimaryButton label="Schritt starten" onPress={onStart} />
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
    width: 200,
    height: 140,
    marginBottom: spacing.xs,
    position: 'relative'
  },
  illustrationCard: {
    position: 'absolute',
    borderRadius: 14,
    borderWidth: 1,
    borderColor: colors.border,
    backgroundColor: colors.panelElevated
  },
  illustrationCardBack: {
    width: 120,
    height: 88,
    top: 18,
    left: 8,
    opacity: 0.45,
    transform: [{ rotate: '-6deg' }]
  },
  illustrationCardMid: {
    width: 120,
    height: 88,
    top: 10,
    left: 44,
    opacity: 0.7,
    transform: [{ rotate: '3deg' }]
  },
  illustrationCardFront: {
    width: 132,
    height: 96,
    top: 28,
    left: 58,
    padding: spacing.sm,
    gap: spacing.xs,
    ...shadows.card
  },
  illustrationLines: {
    gap: 6
  },
  illustrationLine: {
    height: 6,
    borderRadius: 999,
    backgroundColor: colors.border
  },
  illustrationLineShort: {
    width: '70%'
  },
  illustrationTable: {
    position: 'absolute',
    right: 4,
    bottom: 8,
    width: 44,
    height: 44,
    borderRadius: 12,
    alignItems: 'center',
    justifyContent: 'center',
    backgroundColor: 'rgba(42, 95, 143, 0.12)',
    borderWidth: 1,
    borderColor: colors.border
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

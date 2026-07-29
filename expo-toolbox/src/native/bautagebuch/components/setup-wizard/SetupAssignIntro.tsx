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
    icon: 'file-find-outline',
    tint: colors.accent,
    bg: colors.badgeBg,
    title: 'PDF als Orientierung',
    copy: 'In der PDF-Ansicht wird immer das aktuell zuzuordnende Feld hervorgehoben.'
  },
  {
    icon: 'folder-arrow-left-right-outline',
    tint: colors.accent2,
    bg: 'rgba(36, 50, 64, 0.1)',
    title: 'Gruppe auswählen',
    copy: 'Wechsle zur Strukturansicht und ordne das Feld einer zuvor erstellten Gruppe oder Tabelle zu.'
  },
  {
    icon: 'skip-next-outline',
    tint: colors.info,
    bg: 'rgba(42, 95, 143, 0.12)',
    title: 'Automatischer Fortschritt',
    copy: 'Nach jeder Zuordnung springt die App automatisch zum nächsten noch nicht zugeordneten Feld.'
  }
];

type Props = {
  onBack: () => void;
  onStart: () => void;
};

export function SetupAssignIntro({ onBack, onStart }: Props) {
  const insets = useSafeAreaInsets();
  const fade = useRef(new Animated.Value(0)).current;
  const slide = useRef(new Animated.Value(16)).current;
  const arrowPulse = useRef(new Animated.Value(0)).current;

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
      Animated.sequence([
        Animated.timing(arrowPulse, { toValue: 1, duration: 900, useNativeDriver: true }),
        Animated.timing(arrowPulse, { toValue: 0, duration: 900, useNativeDriver: true })
      ])
    ).start();
  }, [fade, slide, arrowPulse]);

  const arrowShift = arrowPulse.interpolate({
    inputRange: [0, 1],
    outputRange: [0, 6]
  });

  return (
    <View style={styles.root}>
      <View style={[styles.header, { paddingTop: spacing.xs }]}>
        <Pressable accessibilityRole="button" style={styles.backBtn} onPress={onBack}>
          <MaterialCommunityIcons name="chevron-left" size={20} color={colors.accent} />
          <Text style={styles.backLabel}>Zurück</Text>
        </Pressable>
        <View style={styles.headerCopy}>
          <Text style={styles.step}>Schritt 2 von 3</Text>
          <Text style={styles.subtitle}>Felder zuordnen</Text>
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
            <View style={styles.pdfPanel}>
              <View style={styles.pdfLine} />
              <View style={styles.pdfLineShort} />
              <View style={styles.pdfFieldHighlight}>
                <MaterialCommunityIcons name="form-textbox" size={16} color={colors.accent} />
              </View>
              <View style={styles.pdfLine} />
            </View>

            <Animated.View style={[styles.flowArrow, { transform: [{ translateX: arrowShift }] }]}>
              <MaterialCommunityIcons name="arrow-right-bold" size={28} color={colors.accent} />
            </Animated.View>

            <View style={styles.groupPanel}>
              <View style={styles.groupIconWrap}>
                <MaterialCommunityIcons name="folder-outline" size={22} color={colors.accent2} />
              </View>
              <Text style={styles.groupLabel}>Allgemeine Angaben</Text>
              <View style={styles.groupAssigned}>
                <MaterialCommunityIcons name="check-circle" size={14} color={colors.success} />
                <Text style={styles.groupAssignedText}>Feld zugeordnet</Text>
              </View>
            </View>
          </View>

          <Text style={styles.headline}>Ordne jedes PDF-Feld der passenden Gruppe zu</Text>
          <Text style={styles.description}>
            In diesem Schritt werden alle erkannten Felder deiner PDF nacheinander angezeigt.
          </Text>
          <Text style={styles.description}>
            Ordne jedes Feld der Gruppe oder Tabelle zu, in der die Information später gespeichert
            werden soll.
          </Text>
          <Text style={styles.descriptionMuted}>
            Du musst die Felder nur zuordnen – Einstellungen wie Pflichtfelder oder Standardwerte
            folgen erst im nächsten Schritt.
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
              Du kannst jederzeit zwischen der PDF-Ansicht und der Strukturansicht wechseln. Das
              aktuell ausgewählte Feld bleibt dabei erhalten, sodass du jederzeit nachvollziehen
              kannst, wo du dich befindest.
            </Text>
          </View>
        </Animated.View>
      </ScrollView>

      <View style={[styles.footer, { paddingBottom: systemBottomInset(insets) + spacing.sm }]}>
        <PrimaryButton label="Zuordnung starten" onPress={onStart} />
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
    maxWidth: 320,
    flexDirection: 'row',
    alignItems: 'center',
    justifyContent: 'center',
    gap: spacing.sm,
    marginBottom: spacing.xs,
    paddingVertical: spacing.sm
  },
  pdfPanel: {
    width: 108,
    minHeight: 120,
    borderRadius: 12,
    borderWidth: 1,
    borderColor: colors.border,
    backgroundColor: colors.panelElevated,
    padding: spacing.sm,
    gap: spacing.xs,
    ...shadows.card
  },
  pdfLine: {
    height: 6,
    borderRadius: 999,
    backgroundColor: colors.border
  },
  pdfLineShort: {
    height: 6,
    width: '65%',
    borderRadius: 999,
    backgroundColor: colors.border
  },
  pdfFieldHighlight: {
    height: 34,
    borderRadius: 8,
    borderWidth: 2,
    borderColor: colors.accent,
    backgroundColor: colors.badgeBg,
    alignItems: 'center',
    justifyContent: 'center'
  },
  flowArrow: {
    paddingHorizontal: 2
  },
  groupPanel: {
    width: 118,
    minHeight: 120,
    borderRadius: 12,
    borderWidth: 1,
    borderColor: colors.accent2,
    backgroundColor: colors.panelElevated,
    padding: spacing.sm,
    gap: spacing.xxs,
    alignItems: 'center',
    justifyContent: 'center',
    ...shadows.card
  },
  groupIconWrap: {
    width: 40,
    height: 40,
    borderRadius: 12,
    alignItems: 'center',
    justifyContent: 'center',
    backgroundColor: 'rgba(36, 50, 64, 0.1)'
  },
  groupLabel: {
    ...typography.caption,
    color: colors.ink,
    fontFamily: 'SpaceGrotesk_600SemiBold',
    textAlign: 'center'
  },
  groupAssigned: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: 4
  },
  groupAssignedText: {
    ...typography.caption,
    color: colors.success,
    fontSize: 11
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

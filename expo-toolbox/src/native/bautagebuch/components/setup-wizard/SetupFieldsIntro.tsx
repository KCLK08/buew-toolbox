import { StyleSheet, Text, View } from 'react-native';
import { MaterialCommunityIcons } from '@expo/vector-icons';

import { SingleLineText } from '../../../../components/mobile';
import { colors, shadows, spacing, typography } from '../../../../constants/theme';

type Props = {
  templateName: string;
  fieldCount: number;
  groupCount: number;
};

export function SetupFieldsIntro({ templateName, fieldCount, groupCount }: Props) {
  return (
    <View style={styles.root}>
      <View style={styles.stepRow}>
        <MaterialCommunityIcons name="numeric-2-circle" size={18} color={colors.accent} />
        <Text style={styles.step}>Schritt 2 von 2</Text>
      </View>
      <Text style={styles.title}>Felder konfigurieren</Text>
      <Text style={styles.copy}>
        Passe Anzeigenamen, Pflichtfelder und Standardwerte an. Die Reihenfolge der Gruppen
        legst du in der hervorgehobenen Kachel fest.
      </Text>

      <View style={styles.summary}>
        <Text style={styles.summaryHeading}>Vorlagenübersicht</Text>
        <View style={styles.summaryRow}>
          <Text style={styles.summaryLabel}>PDF</Text>
          <View style={styles.summaryValueWrap}>
            <SingleLineText style={styles.summaryValue}>{templateName}</SingleLineText>
          </View>
        </View>
        <View style={styles.summaryRow}>
          <Text style={styles.summaryLabel}>Felder</Text>
          <Text style={styles.summaryValue}>{fieldCount}</Text>
        </View>
        <View style={styles.summaryRow}>
          <Text style={styles.summaryLabel}>Gruppen</Text>
          <Text style={styles.summaryValue}>{groupCount}</Text>
        </View>
      </View>
    </View>
  );
}

const styles = StyleSheet.create({
  root: {
    gap: spacing.sm,
    paddingBottom: spacing.sm
  },
  stepRow: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: spacing.xxs
  },
  step: {
    ...typography.caption,
    color: colors.accent,
    fontFamily: 'SpaceGrotesk_600SemiBold'
  },
  title: {
    ...typography.title,
    color: colors.ink
  },
  copy: {
    ...typography.body,
    color: colors.muted
  },
  summary: {
    marginTop: spacing.xxs,
    borderRadius: spacing.cardRadius,
    borderWidth: 1,
    borderColor: colors.border,
    backgroundColor: colors.panelElevated,
    padding: spacing.md,
    gap: spacing.sm,
    ...shadows.card
  },
  summaryHeading: {
    ...typography.label,
    color: colors.muted
  },
  summaryRow: {
    flexDirection: 'row',
    justifyContent: 'space-between',
    gap: spacing.sm
  },
  summaryLabel: {
    ...typography.caption,
    color: colors.muted
  },
  summaryValue: {
    ...typography.bodyStrong,
    color: colors.ink,
    textAlign: 'right'
  },
  summaryValueWrap: {
    flex: 1,
    minWidth: 0,
    alignItems: 'flex-end'
  }
});

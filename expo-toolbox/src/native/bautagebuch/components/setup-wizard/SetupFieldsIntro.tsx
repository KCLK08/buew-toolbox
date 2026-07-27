import { StyleSheet, Text, View } from 'react-native';

import { SingleLineText } from '../../../../components/mobile';
import { colors, spacing, typography } from '../../../../constants/theme';

type Props = {
  templateName: string;
  fieldCount: number;
  groupCount: number;
};

export function SetupFieldsIntro({ templateName, fieldCount, groupCount }: Props) {
  return (
    <View style={styles.root}>
      <Text style={styles.step}>Schritt 2 von 2</Text>
      <Text style={styles.title}>Felder konfigurieren</Text>
      <Text style={styles.copy}>
        Die PDF-Vorlage wurde erfolgreich eingerichtet. Jetzt können die Eigenschaften der
        einzelnen Felder angepasst werden.
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
    paddingBottom: spacing.sm,
    borderBottomWidth: 1,
    borderBottomColor: colors.border,
    marginBottom: spacing.xs
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
    marginTop: spacing.xs,
    borderRadius: 12,
    borderWidth: 1,
    borderColor: colors.border,
    backgroundColor: colors.panelElevated,
    padding: spacing.md,
    gap: spacing.sm
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

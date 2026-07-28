import { StyleSheet, Text, View } from 'react-native';
import { MaterialCommunityIcons } from '@expo/vector-icons';

import { colors, shadows, spacing, typography } from '../../../../constants/theme';
import { SetupTemplateRenameControl } from './SetupTemplateRenameControl';

type Props = {
  templateId: string;
  templateName: string;
  fieldCount: number;
  groupCount: number;
  readOnly?: boolean;
  onRenamed: (nextName: string) => void;
};

export function SetupFieldsIntro({
  templateId,
  templateName,
  fieldCount,
  groupCount,
  readOnly = false,
  onRenamed
}: Props) {
  return (
    <View style={styles.root}>
      <View style={styles.stepRow}>
        <MaterialCommunityIcons name="numeric-3-circle" size={18} color={colors.accent} />
        <Text style={styles.step}>Schritt 3 von 3</Text>
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
          <SetupTemplateRenameControl
            templateId={templateId}
            templateName={templateName}
            readOnly={readOnly}
            onRenamed={onRenamed}
          />
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
  }
});

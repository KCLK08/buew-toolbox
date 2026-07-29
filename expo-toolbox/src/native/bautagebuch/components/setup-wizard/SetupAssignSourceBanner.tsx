import { StyleSheet, Text, View } from 'react-native';
import { MaterialCommunityIcons } from '@expo/vector-icons';

import { colors, spacing, typography } from '../../../../constants/theme';
import { countFieldsBySource } from '../../lib/template-field';
import type { DetectedField } from '../../types';

type Props = {
  fields: DetectedField[];
};

export function SetupAssignSourceBanner({ fields }: Props) {
  const stats = countFieldsBySource(fields);
  const hasAuto = stats.acroform > 0;
  const hasManual = stats.manual > 0;

  return (
    <View style={styles.root}>
      <View style={styles.iconWrap}>
        <MaterialCommunityIcons
          name={hasAuto ? 'file-document-check-outline' : 'gesture-tap-button'}
          size={20}
          color={colors.accent2}
        />
      </View>
      <View style={styles.copy}>
        {hasAuto ? (
          <>
            <Text style={styles.title}>
              {stats.acroform} Feld{stats.acroform === 1 ? '' : 'er'} automatisch erkannt
            </Text>
            <Text style={styles.text}>
              Quelle: AcroForm. Bitte überprüfen und ergänzen Sie fehlende Bereiche.
            </Text>
          </>
        ) : (
          <>
            <Text style={styles.title}>Keine Felder automatisch erkannt</Text>
            <Text style={styles.text}>
              Markieren Sie die benötigten Bereiche in der PDF mit „+ Feld hinzufügen“.
            </Text>
          </>
        )}
        {hasManual ? (
          <Text style={styles.manualHint}>
            {stats.manual} manuell erstellte Feld{stats.manual === 1 ? '' : 'er'}
          </Text>
        ) : null}
      </View>
    </View>
  );
}

const styles = StyleSheet.create({
  root: {
    flexDirection: 'row',
    gap: spacing.sm,
    paddingHorizontal: spacing.pageX,
    paddingVertical: spacing.sm,
    backgroundColor: colors.panel,
    borderBottomWidth: 1,
    borderBottomColor: colors.border
  },
  iconWrap: {
    width: 36,
    height: 36,
    borderRadius: 10,
    alignItems: 'center',
    justifyContent: 'center',
    backgroundColor: colors.badgeBg
  },
  copy: {
    flex: 1,
    gap: 2
  },
  title: {
    ...typography.bodyStrong,
    color: colors.ink
  },
  text: {
    ...typography.caption,
    color: colors.muted,
    lineHeight: 18
  },
  manualHint: {
    ...typography.caption,
    color: colors.warning,
    marginTop: spacing.xxs
  }
});

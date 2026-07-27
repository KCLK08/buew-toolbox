import { StyleSheet, Text, View } from 'react-native';
import { MaterialCommunityIcons } from '@expo/vector-icons';

import { Card, PrimaryButton } from '../../../../components/mobile';
import { colors, spacing, typography } from '../../../../constants/theme';

type Props = {
  creating?: boolean;
  disabled?: boolean;
  onStart: () => void;
};

export function NewBTBCard({ creating = false, disabled = false, onStart }: Props) {
  return (
    <Card style={styles.card}>
      <View style={styles.iconWrap}>
        <MaterialCommunityIcons name="notebook-edit-outline" size={32} color={colors.accent} />
      </View>
      <View style={styles.copy}>
        <Text style={styles.title}>Neues Bautagebuch</Text>
        <Text style={styles.description}>Neue Baustellendokumentation erstellen</Text>
      </View>
      <PrimaryButton
        label={creating ? 'Wird gestartet…' : '+ Starten'}
        disabled={creating || disabled}
        onPress={onStart}
        style={styles.button}
      />
    </Card>
  );
}

const styles = StyleSheet.create({
  card: {
    gap: spacing.md,
    backgroundColor: colors.panelElevated,
    borderColor: colors.border,
    paddingVertical: spacing.lg,
    paddingHorizontal: spacing.md
  },
  iconWrap: {
    width: 56,
    height: 56,
    borderRadius: 16,
    backgroundColor: colors.badgeBg,
    alignItems: 'center',
    justifyContent: 'center'
  },
  copy: {
    gap: spacing.xxs
  },
  title: {
    ...typography.subtitle,
    color: colors.ink,
    fontSize: 20
  },
  description: {
    ...typography.body,
    color: colors.muted,
    fontSize: 15
  },
  button: {
    minHeight: spacing.touchMin + 4
  }
});

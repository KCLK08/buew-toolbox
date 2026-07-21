import { colors } from '../constants/theme';
import { Pressable, StyleSheet, Text, View } from 'react-native';

import type { IntegrityReport } from '../types/offline';

type Props = {
  report: IntegrityReport | null;
  error: string | null;
  onDismiss?: () => void;
};

export function OfflineStatusBanner({ report, error, onDismiss }: Props) {
  if (!error && !report?.restoredFromBackup && (!report || report.ok)) {
    return null;
  }

  const message =
    error ||
    (report?.restoredFromBackup
      ? 'Lokale Datenbank wurde aus einem Backup wiederhergestellt.'
      : report?.issues[0]?.message || 'Offline-Hinweis');

  return (
    <View style={styles.banner}>
      <Text style={styles.text}>{message}</Text>
      {onDismiss ? (
        <Pressable onPress={onDismiss} accessibilityRole="button">
          <Text style={styles.action}>OK</Text>
        </Pressable>
      ) : null}
    </View>
  );
}

const styles = StyleSheet.create({
  banner: {
    backgroundColor: 'rgba(211, 84, 60, 0.12)',
    borderColor: colors.border,
    borderWidth: 1,
    borderRadius: 14,
    paddingHorizontal: 14,
    paddingVertical: 12,
    marginBottom: 16,
    flexDirection: 'row',
    alignItems: 'center',
    gap: 12
  },
  text: {
    flex: 1,
    color: colors.accent2,
    fontFamily: 'SpaceGrotesk_400Regular',
    fontSize: 13,
    lineHeight: 18
  },
  action: {
    color: colors.accent,
    fontFamily: 'SpaceGrotesk_700Bold',
    fontSize: 14,
    paddingVertical: 10,
    paddingHorizontal: 4
  }
});

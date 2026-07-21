import { ActivityIndicator, Pressable, StyleSheet, Text, View } from 'react-native';

import { colors } from '../constants/theme';
import type { IntegrityReport } from '../types/offline';

type Props = {
  report: IntegrityReport | null;
  error: string | null;
  restoreBusy?: boolean;
  onAcceptRestore?: () => void;
  onRejectRestore?: () => void;
  onDismiss?: () => void;
};

function formatBackupDate(iso: string): string {
  try {
    return new Date(iso).toLocaleString('de-DE');
  } catch {
    return iso;
  }
}

export function OfflineStatusBanner({
  report,
  error,
  restoreBusy,
  onAcceptRestore,
  onRejectRestore,
  onDismiss
}: Props) {
  if (report?.pendingRestore) {
    return (
      <View style={styles.banner}>
        <Text style={styles.text}>
          Datenproblem erkannt. Backup vom {formatBackupDate(report.pendingRestore.backupDate)} verfügbar.
          Wiederherstellung überschreibt die aktuelle lokale Datenbank.
        </Text>
        <View style={styles.actions}>
          <Pressable
            accessibilityRole="button"
            disabled={restoreBusy}
            onPress={onAcceptRestore}
            style={styles.primaryAction}
          >
            {restoreBusy ? (
              <ActivityIndicator color={colors.white} />
            ) : (
              <Text style={styles.primaryActionText}>Wiederherstellen</Text>
            )}
          </Pressable>
          <Pressable accessibilityRole="button" disabled={restoreBusy} onPress={onRejectRestore}>
            <Text style={styles.action}>Abbrechen</Text>
          </Pressable>
        </View>
      </View>
    );
  }

  if (report?.restoredFromBackup) {
    return (
      <View style={styles.banner}>
        <Text style={styles.text}>Lokale Datenbank wurde aus dem bestätigten Backup wiederhergestellt.</Text>
        {onDismiss ? (
          <Pressable onPress={onDismiss} accessibilityRole="button">
            <Text style={styles.action}>OK</Text>
          </Pressable>
        ) : null}
      </View>
    );
  }

  if (!error && (!report || report.ok) && !(report?.orphanFiles?.length)) {
    return null;
  }

  const orphanHint =
    report?.orphanFiles?.length && report.orphanFiles.length > 0
      ? ` ${report.orphanFiles.length} verwaiste Datei(en) gemeldet (nicht gelöscht).`
      : '';

  const message =
    error ||
    (report?.issues[0]?.message || 'Offline-Hinweis') + orphanHint;

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
    gap: 10
  },
  text: {
    color: colors.accent2,
    fontFamily: 'SpaceGrotesk_400Regular',
    fontSize: 13,
    lineHeight: 18
  },
  actions: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: 12
  },
  primaryAction: {
    backgroundColor: colors.accent,
    borderRadius: 10,
    paddingHorizontal: 12,
    paddingVertical: 8,
    minWidth: 120,
    alignItems: 'center'
  },
  primaryActionText: {
    color: colors.white,
    fontFamily: 'SpaceGrotesk_700Bold',
    fontSize: 13
  },
  action: {
    color: colors.accent,
    fontFamily: 'SpaceGrotesk_700Bold',
    fontSize: 14,
    paddingVertical: 10,
    paddingHorizontal: 4
  }
});

import { Image, Pressable, StyleSheet, Text, View } from 'react-native';

import { colors, spacing, typography } from '../../constants/theme';
import { PrimaryButton } from '../mobile';

type Props = {
  photoUri: string | null;
  onCapture: () => void;
  onRemove: () => void;
  busy?: boolean;
};

export function PhotoCaptureStep({ photoUri, onCapture, onRemove, busy }: Props) {
  return (
    <View style={styles.wrap}>
      {photoUri ? (
        <Image source={{ uri: photoUri }} style={styles.preview} resizeMode="cover" />
      ) : (
        <Pressable style={styles.placeholder} onPress={onCapture} disabled={busy}>
          <Text style={styles.cameraIcon}>📷</Text>
          <Text style={styles.placeholderTitle}>Foto aufnehmen</Text>
          <Text style={styles.placeholderHint}>Tippe für die Kamera</Text>
        </Pressable>
      )}
      <View style={styles.actions}>
        <PrimaryButton
          label={busy ? 'Kamera…' : photoUri ? 'Neues Foto' : 'Kamera öffnen'}
          onPress={onCapture}
          disabled={busy}
        />
        {photoUri ? <PrimaryButton label="Foto entfernen" variant="ghost" onPress={onRemove} /> : null}
      </View>
    </View>
  );
}

const styles = StyleSheet.create({
  wrap: {
    gap: spacing.md
  },
  preview: {
    width: '100%',
    height: 280,
    borderRadius: spacing.cardRadius,
    backgroundColor: colors.border
  },
  placeholder: {
    height: 280,
    borderRadius: spacing.cardRadius,
    backgroundColor: colors.panel,
    borderWidth: 2,
    borderColor: colors.border,
    borderStyle: 'dashed',
    alignItems: 'center',
    justifyContent: 'center',
    gap: spacing.xs
  },
  cameraIcon: {
    fontSize: 48
  },
  placeholderTitle: {
    ...typography.subtitle,
    color: colors.ink
  },
  placeholderHint: {
    ...typography.caption,
    color: colors.muted
  },
  actions: {
    gap: spacing.xs
  }
});

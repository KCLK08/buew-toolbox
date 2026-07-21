import { Image, Pressable, StyleSheet, Text, View } from 'react-native';

import { colors, spacing, typography } from '../../constants/theme';
import { PrimaryButton } from '../mobile';
import { hapticMedium } from '../../lib/haptics';

type Props = {
  photoUri: string | null;
  onCapture: () => void;
  onRemove: () => void;
  busy?: boolean;
};

export function PhotoCaptureStep({ photoUri, onCapture, onRemove, busy }: Props) {
  const handleCapture = () => {
    void hapticMedium();
    onCapture();
  };

  return (
    <View style={styles.wrap}>
      {photoUri ? (
        <Image source={{ uri: photoUri }} style={styles.preview} resizeMode="cover" />
      ) : (
        <Pressable style={styles.placeholder} onPress={handleCapture} disabled={busy}>
          <View style={styles.cameraCircle}>
            <Text style={styles.cameraIcon}>📷</Text>
          </View>
          <Text style={styles.placeholderTitle}>Foto aufnehmen</Text>
          <Text style={styles.placeholderHint}>Tippe für die Kamera</Text>
        </Pressable>
      )}
      <View style={styles.actions}>
        <PrimaryButton
          label={busy ? 'Kamera…' : photoUri ? 'Neues Foto' : 'Kamera öffnen'}
          onPress={handleCapture}
          disabled={busy}
        />
        {photoUri ? <PrimaryButton label="Foto entfernen" variant="ghost" onPress={onRemove} /> : null}
      </View>
    </View>
  );
}

const styles = StyleSheet.create({
  wrap: {
    gap: spacing.lg
  },
  preview: {
    width: '100%',
    height: 320,
    borderRadius: spacing.cardRadius,
    backgroundColor: colors.border
  },
  placeholder: {
    minHeight: 320,
    borderRadius: spacing.cardRadius,
    backgroundColor: colors.panel,
    borderWidth: 2,
    borderColor: colors.border,
    borderStyle: 'dashed',
    alignItems: 'center',
    justifyContent: 'center',
    gap: spacing.sm,
    padding: spacing.xl
  },
  cameraCircle: {
    width: 80,
    height: 80,
    borderRadius: 40,
    backgroundColor: colors.badgeBg,
    alignItems: 'center',
    justifyContent: 'center'
  },
  cameraIcon: {
    fontSize: 36
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

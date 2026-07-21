import { Pressable, StyleSheet, Text, View } from 'react-native';

import { colors, shadows, spacing, typography } from '../../constants/theme';
import { hapticLight } from '../../lib/haptics';

type Props = {
  title: string;
  subtitle?: string;
  onPress: () => void;
};

export function PrimaryActionCard({ title, subtitle, onPress }: Props) {
  return (
    <Pressable
      accessibilityRole="button"
      onPress={() => {
        void hapticLight();
        onPress();
      }}
      style={({ pressed }) => [styles.card, pressed ? styles.pressed : null]}
    >
      <View style={styles.iconCircle}>
        <Text style={styles.icon}>+</Text>
      </View>
      <View style={styles.text}>
        <Text style={styles.title}>{title}</Text>
        {subtitle ? <Text style={styles.subtitle}>{subtitle}</Text> : null}
      </View>
      <Text style={styles.chevron}>›</Text>
    </Pressable>
  );
}

const styles = StyleSheet.create({
  card: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: spacing.md,
    backgroundColor: colors.accent,
    borderRadius: spacing.cardRadius,
    padding: spacing.lg,
    ...shadows.card
  },
  pressed: {
    opacity: 0.94,
    transform: [{ scale: 0.99 }]
  },
  iconCircle: {
    width: 48,
    height: 48,
    borderRadius: 24,
    backgroundColor: 'rgba(255,255,255,0.2)',
    alignItems: 'center',
    justifyContent: 'center'
  },
  icon: {
    fontSize: 28,
    color: colors.white,
    fontFamily: 'SpaceGrotesk_700Bold',
    lineHeight: 32
  },
  text: {
    flex: 1,
    gap: 2
  },
  title: {
    ...typography.subtitle,
    color: colors.white
  },
  subtitle: {
    ...typography.caption,
    color: 'rgba(255,255,255,0.85)'
  },
  chevron: {
    fontSize: 28,
    color: 'rgba(255,255,255,0.7)',
    fontFamily: 'SpaceGrotesk_400Regular'
  }
});

import { useRef } from 'react';
import { Animated, Pressable, StyleSheet, Text, View } from 'react-native';
import { MaterialCommunityIcons } from '@expo/vector-icons';

import { colors, shadows, spacing, typography } from '../constants/theme';
import type { ToolboxTool } from '../constants/tools';

const TOOL_VECTOR_ICONS: Record<ToolboxTool['id'], keyof typeof MaterialCommunityIcons.glyphMap> = {
  sitereport: 'camera-document',
  bautagebuch: 'notebook-edit-outline'
};

type ToolCardProps = {
  tool: ToolboxTool;
  onPress: () => void;
};

export function ToolCard({ tool, onPress }: ToolCardProps) {
  const lift = useRef(new Animated.Value(0)).current;

  const animateTo = (toValue: number) => {
    Animated.spring(lift, {
      toValue,
      useNativeDriver: true,
      friction: 7,
      tension: 140
    }).start();
  };

  return (
    <Pressable
      accessibilityRole="link"
      accessibilityLabel={`${tool.title}. ${tool.description}`}
      onPress={onPress}
      onPressIn={() => animateTo(1)}
      onPressOut={() => animateTo(0)}
      style={({ pressed }) => [styles.pressable, pressed && styles.pressableActive]}
    >
      <Animated.View
        style={[
          styles.card,
          shadows.card,
          {
            transform: [
              {
                translateY: lift.interpolate({
                  inputRange: [0, 1],
                  outputRange: [0, -3]
                })
              }
            ]
          }
        ]}
      >
        <View style={styles.topRow}>
          <View style={styles.iconWrap}>
            <MaterialCommunityIcons
              name={TOOL_VECTOR_ICONS[tool.id]}
              size={32}
              color={colors.accent}
              accessibilityLabel={tool.iconAlt}
            />
          </View>
          <MaterialCommunityIcons name="arrow-top-right" size={22} color={colors.muted} />
        </View>

        <Text style={styles.title}>{tool.title}</Text>
        <Text style={styles.description}>{tool.description}</Text>

        <View style={styles.features}>
          {tool.features.map((feature) => (
            <View key={feature} style={styles.featureRow}>
              <View style={styles.featureDot} />
              <Text style={styles.featureText}>{feature}</Text>
            </View>
          ))}
        </View>
      </Animated.View>
    </Pressable>
  );
}

const styles = StyleSheet.create({
  pressable: {
    width: '100%'
  },
  pressableActive: {
    opacity: 0.98
  },
  card: {
    backgroundColor: colors.panelElevated,
    borderWidth: 1,
    borderColor: colors.border,
    borderRadius: spacing.cardRadius,
    padding: spacing.lg,
    gap: spacing.sm,
    minHeight: 196
  },
  topRow: {
    flexDirection: 'row',
    alignItems: 'flex-start',
    justifyContent: 'space-between'
  },
  iconWrap: {
    width: 56,
    height: 56,
    borderRadius: 16,
    borderWidth: 1,
    borderColor: colors.border,
    backgroundColor: colors.white,
    alignItems: 'center',
    justifyContent: 'center'
  },
  title: {
    ...typography.title,
    color: colors.ink,
    fontSize: 24
  },
  description: {
    ...typography.body,
    color: colors.muted,
    fontSize: 15,
    lineHeight: 22
  },
  features: {
    gap: spacing.xs,
    paddingTop: spacing.xxs
  },
  featureRow: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: spacing.sm,
    minHeight: 28
  },
  featureDot: {
    width: 6,
    height: 6,
    borderRadius: 999,
    backgroundColor: colors.accent
  },
  featureText: {
    ...typography.caption,
    color: colors.ink,
    fontSize: 14
  }
});

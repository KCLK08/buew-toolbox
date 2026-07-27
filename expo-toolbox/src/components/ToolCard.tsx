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
  compact?: boolean;
  onPress: () => void;
};

export function ToolCard({ tool, compact = false, onPress }: ToolCardProps) {
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
      style={({ pressed }) => [styles.pressable, compact ? styles.pressableCompact : null, pressed && styles.pressableActive]}
    >
      <Animated.View
        style={[
          styles.card,
          compact ? styles.cardCompact : null,
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
          <View style={[styles.iconWrap, compact ? styles.iconWrapCompact : null]}>
            <MaterialCommunityIcons
              name={TOOL_VECTOR_ICONS[tool.id]}
              size={compact ? 28 : 32}
              color={colors.accent}
              accessibilityLabel={tool.iconAlt}
            />
          </View>
          <MaterialCommunityIcons name="arrow-top-right" size={compact ? 20 : 22} color={colors.muted} />
        </View>

        <Text style={[styles.title, compact ? styles.titleCompact : null]}>{tool.title}</Text>
        <Text style={[styles.description, compact ? styles.descriptionCompact : null]}>{tool.description}</Text>

        <View style={[styles.features, compact ? styles.featuresCompact : null]}>
          {tool.features.map((feature) => (
            <View key={feature} style={[styles.featureRow, compact ? styles.featureRowCompact : null]}>
              <View style={styles.featureDot} />
              <Text style={[styles.featureText, compact ? styles.featureTextCompact : null]}>{feature}</Text>
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
  pressableCompact: {
    flex: 1,
    minHeight: 0
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
  cardCompact: {
    flex: 1,
    minHeight: 0,
    padding: spacing.md,
    gap: spacing.xs
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
  iconWrapCompact: {
    width: 48,
    height: 48,
    borderRadius: 14
  },
  title: {
    ...typography.title,
    color: colors.ink,
    fontSize: 24
  },
  titleCompact: {
    fontSize: 20,
    lineHeight: 24
  },
  description: {
    ...typography.body,
    color: colors.muted,
    fontSize: 15,
    lineHeight: 22
  },
  descriptionCompact: {
    fontSize: 14,
    lineHeight: 20
  },
  features: {
    gap: spacing.xs,
    paddingTop: spacing.xxs
  },
  featuresCompact: {
    gap: 2,
    paddingTop: 0
  },
  featureRow: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: spacing.sm,
    minHeight: 28
  },
  featureRowCompact: {
    minHeight: 22,
    gap: spacing.xs
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
  },
  featureTextCompact: {
    fontSize: 13,
    lineHeight: 18
  }
});

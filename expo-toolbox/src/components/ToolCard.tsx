import { useRef } from 'react';
import { Animated, Pressable, StyleSheet, Text, View } from 'react-native';
import { MaterialCommunityIcons } from '@expo/vector-icons';

import type { HomeLayoutTier } from './toolbox/homeLayout';
import { colors, shadows, spacing, typography } from '../constants/theme';
import type { ToolboxTool } from '../constants/tools';

const TOOL_VECTOR_ICONS: Record<ToolboxTool['id'], keyof typeof MaterialCommunityIcons.glyphMap> = {
  sitereport: 'camera-document',
  bautagebuch: 'notebook-edit-outline'
};

type ToolCardProps = {
  tool: ToolboxTool;
  tier?: HomeLayoutTier;
  onPress: () => void;
};

export function ToolCard({ tool, tier = 'relaxed', onPress }: ToolCardProps) {
  const compact = tier !== 'relaxed';
  const dense = tier === 'dense';
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
      style={({ pressed }) => [
        styles.pressable,
        compact ? styles.pressableCompact : null,
        pressed && styles.pressableActive
      ]}
    >
      <Animated.View
        style={[
          styles.card,
          compact ? styles.cardCompact : null,
          dense ? styles.cardDense : null,
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
          <View style={[styles.iconWrap, compact ? styles.iconWrapCompact : null, dense ? styles.iconWrapDense : null]}>
            <MaterialCommunityIcons
              name={TOOL_VECTOR_ICONS[tool.id]}
              size={dense ? 24 : compact ? 26 : 30}
              color={colors.accent}
              accessibilityLabel={tool.iconAlt}
            />
          </View>
          <MaterialCommunityIcons name="arrow-top-right" size={dense ? 18 : compact ? 20 : 22} color={colors.muted} />
        </View>

        <Text style={[styles.title, compact ? styles.titleCompact : null, dense ? styles.titleDense : null]}>
          {tool.title}
        </Text>
        <Text
          style={[
            styles.description,
            compact ? styles.descriptionCompact : null,
            dense ? styles.descriptionDense : null
          ]}
          numberOfLines={dense ? 2 : compact ? 3 : undefined}
        >
          {tool.description}
        </Text>

        {!dense ? (
          <View style={[styles.features, compact ? styles.featuresCompact : null]}>
            {tool.features.map((feature) => (
              <View key={feature} style={[styles.featureRow, compact ? styles.featureRowCompact : null]}>
                <View style={styles.featureDot} />
                <Text style={[styles.featureText, compact ? styles.featureTextCompact : null]}>{feature}</Text>
              </View>
            ))}
          </View>
        ) : null}
      </Animated.View>
    </Pressable>
  );
}

const styles = StyleSheet.create({
  pressable: {
    width: '100%',
    flexGrow: 1,
    flexShrink: 1,
    minHeight: 0
  },
  pressableCompact: {
    flex: 1
  },
  pressableActive: {
    opacity: 0.98
  },
  card: {
    flex: 1,
    backgroundColor: colors.panelElevated,
    borderWidth: 1,
    borderColor: colors.border,
    borderRadius: spacing.cardRadius,
    padding: spacing.lg,
    gap: spacing.sm,
    minHeight: 0
  },
  cardCompact: {
    padding: spacing.md,
    gap: spacing.xs
  },
  cardDense: {
    padding: spacing.sm,
    gap: spacing.xxs
  },
  topRow: {
    flexDirection: 'row',
    alignItems: 'flex-start',
    justifyContent: 'space-between'
  },
  iconWrap: {
    width: 52,
    height: 52,
    borderRadius: 15,
    borderWidth: 1,
    borderColor: colors.border,
    backgroundColor: colors.white,
    alignItems: 'center',
    justifyContent: 'center'
  },
  iconWrapCompact: {
    width: 44,
    height: 44,
    borderRadius: 13
  },
  iconWrapDense: {
    width: 40,
    height: 40,
    borderRadius: 12
  },
  title: {
    ...typography.title,
    color: colors.ink,
    fontSize: 22,
    lineHeight: 28
  },
  titleCompact: {
    fontSize: 19,
    lineHeight: 24
  },
  titleDense: {
    fontSize: 17,
    lineHeight: 22
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
  descriptionDense: {
    fontSize: 13,
    lineHeight: 18
  },
  features: {
    gap: spacing.xxs,
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
    minHeight: 24
  },
  featureRowCompact: {
    minHeight: 20,
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
    fontSize: 13,
    flex: 1
  },
  featureTextCompact: {
    fontSize: 12,
    lineHeight: 16
  }
});

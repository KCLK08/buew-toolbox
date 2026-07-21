import { useRef } from 'react';
import {
  Animated,
  Image,
  Pressable,
  StyleSheet,
  Text,
  View
} from 'react-native';

import { colors, spacing } from '../constants/theme';
import type { ToolboxTool } from '../constants/tools';

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
          {
            transform: [
              {
                translateY: lift.interpolate({
                  inputRange: [0, 1],
                  outputRange: [0, -4]
                })
              }
            ]
          }
        ]}
      >
        <Animated.View
          pointerEvents="none"
          style={[
            styles.glow,
            {
              opacity: lift.interpolate({
                inputRange: [0, 1],
                outputRange: [0, 0.55]
              })
            }
          ]}
        />
        <View style={styles.header}>
          <View style={styles.iconWrap}>
            <Image source={tool.icon} style={styles.icon} accessibilityLabel={tool.iconAlt} />
          </View>
          <Text style={styles.title}>{tool.title}</Text>
        </View>
        <Text style={styles.description}>{tool.description}</Text>
      </Animated.View>
    </Pressable>
  );
}

const styles = StyleSheet.create({
  pressable: {
    flexGrow: 1,
    minWidth: 260,
    maxWidth: 520
  },
  pressableActive: {
    opacity: 0.98
  },
  card: {
    backgroundColor: colors.panel,
    borderWidth: 1,
    borderColor: colors.border,
    borderRadius: spacing.cardRadius,
    padding: spacing.cardPadding,
    overflow: 'hidden',
    shadowColor: '#171512',
    shadowOffset: { width: 0, height: 14 },
    shadowRadius: 30,
    shadowOpacity: 0.08,
    elevation: 4,
    minHeight: 132
  },
  glow: {
    position: 'absolute',
    top: -40,
    right: -20,
    width: 180,
    height: 120,
    borderRadius: 90,
    backgroundColor: colors.glow
  },
  header: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: 12,
    marginBottom: 10
  },
  iconWrap: {
    width: spacing.iconSize,
    height: spacing.iconSize,
    borderRadius: spacing.iconRadius,
    borderWidth: 1,
    borderColor: colors.border,
    backgroundColor: colors.white,
    overflow: 'hidden',
    alignItems: 'center',
    justifyContent: 'center'
  },
  icon: {
    width: '100%',
    height: '100%',
    resizeMode: 'cover'
  },
  title: {
    color: colors.ink,
    fontSize: 22,
    fontFamily: 'SpaceGrotesk_700Bold',
    letterSpacing: -0.3
  },
  description: {
    color: colors.muted,
    fontSize: 15,
    lineHeight: 22,
    fontFamily: 'SpaceGrotesk_400Regular'
  }
});

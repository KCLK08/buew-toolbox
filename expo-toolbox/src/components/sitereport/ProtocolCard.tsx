import { Pressable, StyleSheet, Text, View } from 'react-native';

import { colors, spacing, typography } from '../../constants/theme';
import { Card } from '../mobile/Card';

type Props = {
  title: string;
  subtitle?: string;
  meta?: string;
  onPress?: () => void;
  trailing?: React.ReactNode;
  selected?: boolean;
  onSelectToggle?: () => void;
};

export function ProtocolCard({
  title,
  subtitle,
  meta,
  onPress,
  trailing,
  selected,
  onSelectToggle
}: Props) {
  return (
    <Pressable onPress={onPress} disabled={!onPress}>
      <Card style={selected ? { ...styles.card, ...styles.selected } : styles.card}>
        <View style={styles.row}>
          {onSelectToggle ? (
            <Pressable onPress={onSelectToggle} hitSlop={8} style={styles.checkbox}>
              <View style={[styles.checkboxInner, selected ? styles.checkboxOn : null]}>
                {selected ? <Text style={styles.checkmark}>✓</Text> : null}
              </View>
            </Pressable>
          ) : null}
          <View style={styles.body}>
            <Text style={styles.title} numberOfLines={2}>
              {title}
            </Text>
            {subtitle ? (
              <Text style={styles.subtitle} numberOfLines={1}>
                {subtitle}
              </Text>
            ) : null}
            {meta ? <Text style={styles.meta}>{meta}</Text> : null}
          </View>
          {trailing}
        </View>
      </Card>
    </Pressable>
  );
}

const styles = StyleSheet.create({
  card: {
    marginBottom: spacing.sm
  },
  selected: {
    borderColor: colors.accent,
    borderWidth: 2
  },
  row: {
    flexDirection: 'row',
    alignItems: 'flex-start',
    gap: spacing.sm
  },
  checkbox: {
    paddingTop: 2
  },
  checkboxInner: {
    width: 24,
    height: 24,
    borderRadius: 6,
    borderWidth: 2,
    borderColor: colors.borderStrong,
    alignItems: 'center',
    justifyContent: 'center',
    backgroundColor: colors.panelElevated
  },
  checkboxOn: {
    borderColor: colors.accent,
    backgroundColor: colors.accent
  },
  checkmark: {
    color: colors.white,
    fontWeight: '700',
    fontSize: 14
  },
  body: {
    flex: 1,
    gap: 2
  },
  title: {
    ...typography.bodyStrong,
    color: colors.ink
  },
  subtitle: {
    ...typography.body,
    color: colors.muted
  },
  meta: {
    ...typography.caption,
    color: colors.accent2,
    marginTop: 2
  }
});

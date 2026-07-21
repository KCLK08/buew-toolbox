import { ReactNode } from 'react';
import {
  KeyboardAvoidingView,
  Platform,
  Pressable,
  RefreshControl,
  ScrollView,
  StyleSheet,
  Text,
  View,
  type ViewStyle
} from 'react-native';
import { useRouter } from 'expo-router';
import { useSafeAreaInsets } from 'react-native-safe-area-context';

import { colors, spacing, typography } from '../../constants/theme';

type Props = {
  title: string;
  children: ReactNode;
  subtitle?: string;
  scroll?: boolean;
  showBack?: boolean;
  rightAction?: ReactNode;
  footer?: ReactNode;
  contentStyle?: ViewStyle;
  refreshing?: boolean;
  onRefresh?: () => void;
};

export function Screen({
  title,
  subtitle,
  children,
  scroll = true,
  showBack = false,
  rightAction,
  footer,
  contentStyle,
  refreshing,
  onRefresh
}: Props) {
  const insets = useSafeAreaInsets();
  const router = useRouter();

  const body = scroll ? (
    <ScrollView
      keyboardShouldPersistTaps="handled"
      contentContainerStyle={[
        styles.content,
        { paddingBottom: Math.max(spacing.pageBottom, insets.bottom + 24) },
        contentStyle
      ]}
      showsVerticalScrollIndicator={false}
      refreshControl={
        onRefresh ? (
          <RefreshControl refreshing={Boolean(refreshing)} onRefresh={onRefresh} tintColor={colors.accent} />
        ) : undefined
      }
    >
      {children}
    </ScrollView>
  ) : (
    <View style={[styles.content, styles.flex, contentStyle]}>{children}</View>
  );

  return (
    <View style={[styles.root, { paddingTop: insets.top }]}>
      <View style={styles.header}>
        <View style={styles.headerSide}>
          {showBack ? (
            <Pressable
              accessibilityRole="button"
              accessibilityLabel="Zurück"
              hitSlop={8}
              onPress={() => router.back()}
              style={styles.backBtn}
            >
              <Text style={styles.backLabel}>‹ Zurück</Text>
            </Pressable>
          ) : (
            <View style={styles.backPlaceholder} />
          )}
        </View>
        <View style={styles.headerCenter}>
          <Text style={styles.title} numberOfLines={1}>
            {title}
          </Text>
          {subtitle ? (
            <Text style={styles.subtitle} numberOfLines={1}>
              {subtitle}
            </Text>
          ) : null}
        </View>
        <View style={[styles.headerSide, styles.headerRight]}>{rightAction}</View>
      </View>

      <KeyboardAvoidingView
        style={styles.flex}
        behavior={Platform.OS === 'ios' ? 'padding' : undefined}
        keyboardVerticalOffset={8}
      >
        {body}
        {footer ? (
          <View style={[styles.footer, { paddingBottom: Math.max(insets.bottom, spacing.sm) }]}>
            {footer}
          </View>
        ) : null}
      </KeyboardAvoidingView>
    </View>
  );
}

const styles = StyleSheet.create({
  root: {
    flex: 1,
    backgroundColor: colors.bg
  },
  flex: {
    flex: 1
  },
  header: {
    minHeight: spacing.touchMin + 8,
    paddingHorizontal: spacing.pageX,
    paddingBottom: spacing.xs,
    flexDirection: 'row',
    alignItems: 'center',
    borderBottomWidth: 1,
    borderBottomColor: colors.border,
    backgroundColor: colors.panel
  },
  headerSide: {
    width: 88,
    justifyContent: 'center'
  },
  headerRight: {
    alignItems: 'flex-end'
  },
  headerCenter: {
    flex: 1,
    alignItems: 'center'
  },
  title: {
    ...typography.subtitle,
    color: colors.ink
  },
  subtitle: {
    ...typography.caption,
    color: colors.muted
  },
  backBtn: {
    minHeight: spacing.touchMin,
    justifyContent: 'center'
  },
  backPlaceholder: {
    minHeight: spacing.touchMin
  },
  backLabel: {
    ...typography.bodyStrong,
    color: colors.accent
  },
  content: {
    paddingHorizontal: spacing.pageX,
    paddingTop: spacing.md,
    gap: spacing.sm
  },
  footer: {
    borderTopWidth: 1,
    borderTopColor: colors.border,
    backgroundColor: colors.panel,
    paddingHorizontal: spacing.pageX,
    paddingTop: spacing.sm,
    gap: spacing.xs
  }
});

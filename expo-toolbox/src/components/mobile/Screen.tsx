import { ReactNode, useRef } from 'react';
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
import { KeyboardScrollProvider, useKeyboardScroll } from '../../contexts/KeyboardScrollContext';
import { tabBarBottomInset, tabBarReservedHeight } from '../../navigation/appTabBar';

type Props = {
  title: string;
  children: ReactNode;
  subtitle?: string;
  scroll?: boolean;
  showBack?: boolean;
  toolboxBack?: boolean;
  backLabel?: string;
  onBack?: () => void;
  reserveTabBarSpace?: boolean;
  rightAction?: ReactNode;
  footer?: ReactNode;
  overlay?: ReactNode;
  compactFooter?: boolean;
  contentStyle?: ViewStyle;
  refreshing?: boolean;
  onRefresh?: () => void;
};

function ScreenScrollBody({
  children,
  scroll,
  contentStyle,
  refreshing,
  onRefresh,
  insets,
  reserveTabBarSpace
}: {
  children: ReactNode;
  scroll: boolean;
  contentStyle?: ViewStyle;
  refreshing?: boolean;
  onRefresh?: () => void;
  insets: { bottom: number };
  reserveTabBarSpace?: boolean;
}) {
  const keyboardScroll = useKeyboardScroll();
  const scrollRef = useRef<ScrollView>(null);
  const safeBottom = tabBarBottomInset(insets);

  if (!scroll) {
    return <View style={[styles.content, styles.flex, contentStyle]}>{children}</View>;
  }

  return (
    <ScrollView
      ref={(node) => {
        scrollRef.current = node;
        keyboardScroll?.attachScrollView(node);
      }}
      keyboardShouldPersistTaps="handled"
      keyboardDismissMode="interactive"
      automaticallyAdjustKeyboardInsets={Platform.OS === 'ios'}
      onScroll={(event) => {
        keyboardScroll?.reportScrollY(event.nativeEvent.contentOffset.y);
      }}
      scrollEventThrottle={16}
      contentContainerStyle={[
        styles.content,
        {
          paddingBottom: Math.max(
            spacing.pageBottom,
            safeBottom + (reserveTabBarSpace ? tabBarReservedHeight(insets) + spacing.sm : spacing.lg)
          )
        },
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
  );
}

export function Screen({
  title,
  subtitle,
  children,
  scroll = true,
  showBack = false,
  toolboxBack = false,
  backLabel,
  onBack,
  reserveTabBarSpace = false,
  rightAction,
  footer,
  overlay,
  compactFooter = false,
  contentStyle,
  refreshing,
  onRefresh
}: Props) {
  const insets = useSafeAreaInsets();
  const router = useRouter();
  const headerHeight = spacing.touchMin + 8 + spacing.xs;
  const safeBottom = tabBarBottomInset(insets);
  const footerBarHeight = compactFooter ? 52 : spacing.touchMin + 8;
  const footerInset = footer ? footerBarHeight + spacing.sm + safeBottom : safeBottom;
  const showHeaderBack = showBack || toolboxBack;
  const resolvedBackLabel = backLabel || (toolboxBack ? '‹ Toolbox' : '‹ Zurück');
  const handleBack =
    onBack ||
    (toolboxBack
      ? () => {
          if (router.canGoBack()) {
            router.back();
            return;
          }
          router.replace('/');
        }
      : () => router.back());

  return (
    <View style={[styles.root, { paddingTop: insets.top }]}>
      <View style={styles.header}>
        <View style={styles.headerSide}>
          {showHeaderBack ? (
            <Pressable
              accessibilityRole="button"
              accessibilityLabel={toolboxBack ? 'Zurück zur Toolbox' : 'Zurück'}
              hitSlop={8}
              onPress={handleBack}
              style={styles.backBtn}
            >
              <Text style={styles.backLabel}>{resolvedBackLabel}</Text>
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

      <KeyboardScrollProvider footerInset={footerInset}>
        <View style={styles.flex}>
          <View style={styles.bodySlot}>
            <KeyboardAvoidingView
              style={styles.flex}
              behavior={Platform.OS === 'ios' ? 'padding' : 'height'}
              keyboardVerticalOffset={Platform.OS === 'ios' ? insets.top + headerHeight : 0}
            >
              <ScreenScrollBody
                scroll={scroll}
                contentStyle={contentStyle}
                refreshing={refreshing}
                onRefresh={onRefresh}
                insets={insets}
                reserveTabBarSpace={reserveTabBarSpace}
              >
                {children}
              </ScreenScrollBody>
            </KeyboardAvoidingView>
            {overlay ? <View style={styles.overlaySlot}>{overlay}</View> : null}
          </View>
          {footer ? (
            <View
              style={[
                styles.footer,
                compactFooter ? styles.footerCompact : null,
                { paddingBottom: safeBottom + (compactFooter ? spacing.sm : spacing.md) }
              ]}
            >
              {footer}
            </View>
          ) : null}
        </View>
      </KeyboardScrollProvider>
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
  bodySlot: {
    flex: 1,
    position: 'relative'
  },
  overlaySlot: {
    ...StyleSheet.absoluteFillObject,
    zIndex: 20
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
  },
  footerCompact: {
    paddingTop: spacing.xs,
    paddingHorizontal: spacing.sm,
    gap: spacing.xs
  }
});

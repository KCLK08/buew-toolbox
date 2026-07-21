import { colors } from '@buew/shared';
import type { PropsWithChildren, ReactNode } from 'react';
import { StyleSheet, Text, View } from 'react-native';
import { useSafeAreaInsets } from 'react-native-safe-area-context';

import { ToolboxBackground } from '../components/ToolboxBackground';

type AuthScreenProps = PropsWithChildren<{
  title: string;
  subtitle: string;
  footer?: ReactNode;
}>;

export function AuthScreen({ title, subtitle, children, footer }: AuthScreenProps) {
  const insets = useSafeAreaInsets();

  return (
    <ToolboxBackground>
      <View style={[styles.wrap, { paddingTop: insets.top + 24, paddingBottom: insets.bottom + 24 }]}>
        <View style={styles.card}>
          <Text style={styles.brand}>BÜW-Toolbox</Text>
          <Text style={styles.title}>{title}</Text>
          <Text style={styles.subtitle}>{subtitle}</Text>
          {children}
          {footer ? <View style={styles.footer}>{footer}</View> : null}
        </View>
      </View>
    </ToolboxBackground>
  );
}

const styles = StyleSheet.create({
  wrap: {
    flex: 1,
    paddingHorizontal: 20,
    justifyContent: 'center'
  },
  card: {
    backgroundColor: colors.panel,
    borderColor: colors.border,
    borderWidth: 1,
    borderRadius: 18,
    padding: 22,
    gap: 12,
    shadowColor: '#171512',
    shadowOpacity: 0.1,
    shadowRadius: 24,
    shadowOffset: { width: 0, height: 12 },
    elevation: 3
  },
  brand: {
    fontFamily: 'SpaceGrotesk_700Bold',
    color: colors.ink,
    fontSize: 15
  },
  title: {
    fontFamily: 'SpaceGrotesk_700Bold',
    color: colors.ink,
    fontSize: 28
  },
  subtitle: {
    fontFamily: 'SpaceGrotesk_400Regular',
    color: colors.muted,
    fontSize: 15,
    lineHeight: 22,
    marginBottom: 4
  },
  footer: {
    marginTop: 8,
    gap: 10
  }
});

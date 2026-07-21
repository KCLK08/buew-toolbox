import { createContext, useCallback, useContext, useMemo, useRef, useState, type ReactNode } from 'react';
import { Animated, StyleSheet, Text, View } from 'react-native';
import { useSafeAreaInsets } from 'react-native-safe-area-context';

import { colors, spacing, typography } from '../constants/theme';

type ToastContextValue = {
  showToast: (message: string) => void;
};

const ToastContext = createContext<ToastContextValue | null>(null);

export function ToastProvider({ children }: { children: ReactNode }) {
  const insets = useSafeAreaInsets();
  const [message, setMessage] = useState('');
  const opacity = useRef(new Animated.Value(0)).current;
  const timerRef = useRef<ReturnType<typeof setTimeout> | null>(null);

  const showToast = useCallback(
    (nextMessage: string) => {
      if (timerRef.current) clearTimeout(timerRef.current);
      setMessage(nextMessage);
      Animated.sequence([
        Animated.timing(opacity, { toValue: 1, duration: 180, useNativeDriver: true }),
        Animated.delay(2200),
        Animated.timing(opacity, { toValue: 0, duration: 220, useNativeDriver: true })
      ]).start();
      timerRef.current = setTimeout(() => setMessage(''), 2700);
    },
    [opacity]
  );

  const value = useMemo(() => ({ showToast }), [showToast]);

  return (
    <ToastContext.Provider value={value}>
      {children}
      {message ? (
        <Animated.View
          pointerEvents="none"
          style={[styles.toastWrap, { bottom: insets.bottom + 72, opacity }]}
        >
          <View style={styles.toast}>
            <Text style={styles.toastText}>{message}</Text>
          </View>
        </Animated.View>
      ) : null}
    </ToastContext.Provider>
  );
}

export function useToast() {
  const ctx = useContext(ToastContext);
  if (!ctx) {
    throw new Error('useToast must be used within ToastProvider');
  }
  return ctx;
}

const styles = StyleSheet.create({
  toastWrap: {
    position: 'absolute',
    left: spacing.pageX,
    right: spacing.pageX,
    alignItems: 'center',
    zIndex: 999
  },
  toast: {
    backgroundColor: colors.ink,
    borderRadius: 999,
    paddingHorizontal: spacing.lg,
    paddingVertical: spacing.sm,
    maxWidth: '100%'
  },
  toastText: {
    ...typography.bodyStrong,
    color: colors.white,
    textAlign: 'center'
  }
});

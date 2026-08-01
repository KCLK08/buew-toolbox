import type { ReactNode } from 'react';
import { KeyboardAvoidingView, Platform, StyleSheet, type ViewStyle } from 'react-native';
import { useSafeAreaInsets } from 'react-native-safe-area-context';

import { KeyboardScrollProvider } from '../../../../contexts/KeyboardScrollContext';
import { systemBottomInset } from '../../../../navigation/systemInsets';

type Props = {
  children: ReactNode;
  style?: ViewStyle;
};

export function SetupModalKeyboardFrame({ children, style }: Props) {
  const insets = useSafeAreaInsets();
  const footerInset = systemBottomInset(insets);

  return (
    <KeyboardScrollProvider footerInset={footerInset}>
      <KeyboardAvoidingView
        style={[styles.flex, style]}
        behavior={Platform.OS === 'ios' ? 'padding' : 'height'}
      >
        {children}
      </KeyboardAvoidingView>
    </KeyboardScrollProvider>
  );
}

const styles = StyleSheet.create({
  flex: {
    flex: 1
  }
});

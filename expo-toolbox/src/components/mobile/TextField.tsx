import { forwardRef, useEffect, useImperativeHandle, useRef, useState } from 'react';
import { StyleSheet, Text, TextInput, View, type TextInputProps } from 'react-native';

import { useKeyboardScroll } from '../../contexts/KeyboardScrollContext';
import { colors, spacing, typography } from '../../constants/theme';

type Props = TextInputProps & {
  label: string;
  hint?: string;
  autoGrow?: boolean;
};

const MULTILINE_MIN_HEIGHT = 132;
const MULTILINE_MAX_HEIGHT = 320;

export const TextField = forwardRef<TextInput, Props>(function TextField(
  { label, hint, style, multiline, autoGrow, value, onContentSizeChange, onFocus, ...rest },
  ref
) {
  const inputRef = useRef<TextInput>(null);
  const [inputHeight, setInputHeight] = useState<number | undefined>(undefined);
  const keyboardScroll = useKeyboardScroll();

  useImperativeHandle(ref, () => inputRef.current as TextInput);

  useEffect(() => {
    if (!multiline || !autoGrow) {
      setInputHeight(undefined);
    }
  }, [multiline, autoGrow, value]);

  const minHeight = multiline ? (autoGrow ? MULTILINE_MIN_HEIGHT : 96) : spacing.touchMin + 8;

  return (
    <View style={styles.wrap}>
      <Text style={styles.label}>{label}</Text>
      <TextInput
        ref={inputRef}
        placeholderTextColor={colors.muted}
        multiline={multiline}
        textAlignVertical={multiline ? 'top' : 'center'}
        scrollEnabled={!(multiline && autoGrow)}
        value={value}
        onFocus={(event) => {
          keyboardScroll?.scrollInputIntoView(inputRef.current);
          onFocus?.(event);
        }}
        onContentSizeChange={(event) => {
          if (multiline && autoGrow) {
            const nextHeight = Math.min(
              MULTILINE_MAX_HEIGHT,
              Math.max(MULTILINE_MIN_HEIGHT, event.nativeEvent.contentSize.height + spacing.md)
            );
            setInputHeight(nextHeight);
          }
          onContentSizeChange?.(event);
        }}
        style={[
          styles.input,
          multiline ? styles.inputMultiline : null,
          multiline && autoGrow ? { minHeight, height: inputHeight ?? minHeight } : multiline ? { minHeight } : null,
          style
        ]}
        {...rest}
      />
      {hint ? <Text style={styles.hint}>{hint}</Text> : null}
    </View>
  );
});

const styles = StyleSheet.create({
  wrap: {
    gap: spacing.xs,
    marginBottom: spacing.sm
  },
  label: {
    ...typography.label,
    color: colors.ink
  },
  input: {
    minHeight: spacing.touchMin + 8,
    borderWidth: 1,
    borderColor: colors.borderStrong,
    borderRadius: spacing.inputRadius,
    backgroundColor: colors.panelElevated,
    paddingHorizontal: spacing.md,
    paddingVertical: spacing.sm,
    color: colors.ink,
    ...typography.body
  },
  inputMultiline: {
    lineHeight: 22
  },
  hint: {
    ...typography.caption,
    color: colors.muted
  }
});

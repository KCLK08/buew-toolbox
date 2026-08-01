import { forwardRef, useCallback, useRef } from 'react';
import { ScrollView, type ScrollViewProps } from 'react-native';

import { useKeyboardScroll } from '../../../../contexts/KeyboardScrollContext';

type Props = ScrollViewProps & {
  keyboardAware?: boolean;
};

export const SetupScrollView = forwardRef<ScrollView, Props>(function SetupScrollView(
  { onScroll, keyboardAware = true, ...rest },
  ref
) {
  const keyboardScroll = useKeyboardScroll();
  const attachedRef = useRef<ScrollView | null>(null);

  const handleRef = useCallback(
    (node: ScrollView | null) => {
      if (keyboardAware) {
        if (!node) {
          if (attachedRef.current) {
            keyboardScroll?.detachScrollView(attachedRef.current);
            attachedRef.current = null;
          }
        } else {
          if (attachedRef.current && attachedRef.current !== node) {
            keyboardScroll?.detachScrollView(attachedRef.current);
          }
          attachedRef.current = node;
          keyboardScroll?.attachScrollView(node);
        }
      }
      if (typeof ref === 'function') {
        ref(node);
      } else if (ref) {
        ref.current = node;
      }
    },
    [keyboardAware, keyboardScroll, ref]
  );

  return (
    <ScrollView
      ref={handleRef}
      keyboardShouldPersistTaps="handled"
      onScroll={(event) => {
        if (keyboardAware) {
          keyboardScroll?.reportScrollY(event.nativeEvent.contentOffset.y);
        }
        onScroll?.(event);
      }}
      scrollEventThrottle={16}
      {...rest}
    />
  );
});

import { forwardRef, useCallback } from 'react';
import { ScrollView, type ScrollViewProps } from 'react-native';

import { useKeyboardScroll } from '../../../../contexts/KeyboardScrollContext';

type Props = ScrollViewProps;

export const SetupScrollView = forwardRef<ScrollView, Props>(function SetupScrollView(
  { onScroll, ...rest },
  ref
) {
  const keyboardScroll = useKeyboardScroll();

  const handleRef = useCallback(
    (node: ScrollView | null) => {
      keyboardScroll?.attachScrollView(node);
      if (typeof ref === 'function') {
        ref(node);
      } else if (ref) {
        ref.current = node;
      }
    },
    [keyboardScroll, ref]
  );

  return (
    <ScrollView
      ref={handleRef}
      keyboardShouldPersistTaps="handled"
      onScroll={(event) => {
        keyboardScroll?.reportScrollY(event.nativeEvent.contentOffset.y);
        onScroll?.(event);
      }}
      scrollEventThrottle={16}
      {...rest}
    />
  );
});

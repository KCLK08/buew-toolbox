import {
  createContext,
  useCallback,
  useContext,
  useEffect,
  useRef,
  type ReactNode
} from 'react';
import { Dimensions, Keyboard, Platform, ScrollView, type TextInput } from 'react-native';

import {
  resolveKeyboardScrollDelta,
  resolveKeyboardVisibleBounds,
  type KeyboardMetrics
} from '../lib/keyboard-scroll';

type KeyboardScrollContextValue = {
  attachScrollView: (ref: ScrollView | null) => void;
  detachScrollView: (ref: ScrollView | null) => void;
  reportScrollY: (offsetY: number) => void;
  scrollInputIntoView: (input: TextInput | null) => void;
};

const KeyboardScrollContext = createContext<KeyboardScrollContextValue | null>(null);

type ProviderProps = {
  children: ReactNode;
  footerInset?: number;
  topInset?: number;
};

const ANDROID_SCROLL_DELAYS_MS = [0, 80, 180, 320];
const IOS_SCROLL_DELAYS_MS = [0, 60, 140];

export function KeyboardScrollProvider({
  children,
  footerInset = 0,
  topInset = 0
}: ProviderProps) {
  const scrollStackRef = useRef<ScrollView[]>([]);
  const scrollRef = useRef<ScrollView | null>(null);
  const scrollYRef = useRef(0);
  const keyboardMetricsRef = useRef<KeyboardMetrics | null>(null);
  const lastFocusedInputRef = useRef<TextInput | null>(null);

  const getActiveScrollView = useCallback(() => {
    return scrollRef.current ?? scrollStackRef.current[scrollStackRef.current.length - 1] ?? null;
  }, []);

  const performScroll = useCallback(
    (input: TextInput | null) => {
      const scroll = getActiveScrollView();
      if (!scroll || !input) return;

      input.measureInWindow((_fieldX, fieldY, _fieldWidth, fieldHeight) => {
        const windowHeight = Dimensions.get('window').height;
        const { top, bottom } = resolveKeyboardVisibleBounds(
          windowHeight,
          keyboardMetricsRef.current,
          footerInset,
          topInset
        );
        const delta = resolveKeyboardScrollDelta(fieldY, fieldHeight, top, bottom);
        if (delta === 0) return;
        scroll.scrollTo({ y: Math.max(0, scrollYRef.current + delta), animated: true });
      });
    },
    [footerInset, getActiveScrollView, topInset]
  );

  const scrollInputIntoView = useCallback(
    (input: TextInput | null) => {
      if (!input) return;
      lastFocusedInputRef.current = input;
      const delays = Platform.OS === 'android' ? ANDROID_SCROLL_DELAYS_MS : IOS_SCROLL_DELAYS_MS;
      delays.forEach((delay) => {
        setTimeout(() => performScroll(input), delay);
      });
    },
    [performScroll]
  );

  useEffect(() => {
    const showEvent = Platform.OS === 'ios' ? 'keyboardWillShow' : 'keyboardDidShow';
    const hideEvent = Platform.OS === 'ios' ? 'keyboardWillHide' : 'keyboardDidHide';
    const showSub = Keyboard.addListener(showEvent, (event) => {
      keyboardMetricsRef.current = {
        screenY: event.endCoordinates.screenY,
        height: event.endCoordinates.height
      };
      if (lastFocusedInputRef.current) {
        scrollInputIntoView(lastFocusedInputRef.current);
      }
    });
    const hideSub = Keyboard.addListener(hideEvent, () => {
      keyboardMetricsRef.current = null;
      lastFocusedInputRef.current = null;
    });
    return () => {
      showSub.remove();
      hideSub.remove();
    };
  }, [scrollInputIntoView]);

  const attachScrollView = useCallback((ref: ScrollView | null) => {
    if (!ref) return;
    scrollStackRef.current = scrollStackRef.current.filter((item) => item !== ref);
    scrollStackRef.current.push(ref);
    scrollRef.current = ref;
  }, []);

  const detachScrollView = useCallback((ref: ScrollView | null) => {
    if (!ref) return;
    scrollStackRef.current = scrollStackRef.current.filter((item) => item !== ref);
    scrollRef.current = scrollStackRef.current[scrollStackRef.current.length - 1] ?? null;
  }, []);

  const reportScrollY = useCallback((offsetY: number) => {
    scrollYRef.current = offsetY;
  }, []);

  return (
    <KeyboardScrollContext.Provider
      value={{ attachScrollView, detachScrollView, reportScrollY, scrollInputIntoView }}
    >
      {children}
    </KeyboardScrollContext.Provider>
  );
}

export function useKeyboardScroll() {
  return useContext(KeyboardScrollContext);
}

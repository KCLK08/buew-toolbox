import {
  createContext,
  useCallback,
  useContext,
  useEffect,
  useRef,
  type ReactNode
} from 'react';
import { Dimensions, Keyboard, Platform, ScrollView, type TextInput } from 'react-native';

type KeyboardScrollContextValue = {
  attachScrollView: (ref: ScrollView | null) => void;
  reportScrollY: (offsetY: number) => void;
  scrollInputIntoView: (input: TextInput | null) => void;
};

const KeyboardScrollContext = createContext<KeyboardScrollContextValue | null>(null);

type ProviderProps = {
  children: ReactNode;
  footerInset?: number;
};

export function KeyboardScrollProvider({ children, footerInset = 0 }: ProviderProps) {
  const scrollRef = useRef<ScrollView | null>(null);
  const scrollYRef = useRef(0);
  const keyboardHeightRef = useRef(0);
  const lastFocusedInputRef = useRef<TextInput | null>(null);

  const scrollInputIntoView = useCallback(
    (input: TextInput | null) => {
      const scroll = scrollRef.current;
      if (!scroll || !input) return;

      lastFocusedInputRef.current = input;
      const delay = Platform.OS === 'ios' ? 60 : 140;
      setTimeout(() => {
        input.measureInWindow((_fieldX, fieldY, _fieldWidth, fieldHeight) => {
          const windowHeight = Dimensions.get('window').height;
          const keyboardHeight = keyboardHeightRef.current;
          const visibleBottom = windowHeight - keyboardHeight - footerInset;
          const fieldBottom = fieldY + fieldHeight;
          const padding = 28;

          if (fieldBottom > visibleBottom - padding) {
            const delta = fieldBottom - visibleBottom + padding;
            scroll.scrollTo({ y: scrollYRef.current + delta, animated: true });
          }
        });
      }, delay);
    },
    [footerInset]
  );

  useEffect(() => {
    const showEvent = Platform.OS === 'ios' ? 'keyboardWillShow' : 'keyboardDidShow';
    const hideEvent = Platform.OS === 'ios' ? 'keyboardWillHide' : 'keyboardDidHide';
    const showSub = Keyboard.addListener(showEvent, (event) => {
      keyboardHeightRef.current = event.endCoordinates.height;
      if (lastFocusedInputRef.current) {
        scrollInputIntoView(lastFocusedInputRef.current);
      }
    });
    const hideSub = Keyboard.addListener(hideEvent, () => {
      keyboardHeightRef.current = 0;
      lastFocusedInputRef.current = null;
    });
    return () => {
      showSub.remove();
      hideSub.remove();
    };
  }, [scrollInputIntoView]);

  const attachScrollView = useCallback((ref: ScrollView | null) => {
    scrollRef.current = ref;
  }, []);

  const reportScrollY = useCallback((offsetY: number) => {
    scrollYRef.current = offsetY;
  }, []);

  return (
    <KeyboardScrollContext.Provider value={{ attachScrollView, reportScrollY, scrollInputIntoView }}>
      {children}
    </KeyboardScrollContext.Provider>
  );
}

export function useKeyboardScroll() {
  return useContext(KeyboardScrollContext);
}

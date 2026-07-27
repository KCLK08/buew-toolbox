import { ScrollView, StyleSheet, Text, type TextStyle } from 'react-native';

type Props = {
  children: string;
  style?: TextStyle;
  centered?: boolean;
};

/** Shows full text on one line with horizontal scroll — no ellipsis, no wrap. */
export function SingleLineText({ children, style, centered = false }: Props) {
  return (
    <ScrollView
      horizontal
      nestedScrollEnabled
      showsHorizontalScrollIndicator={false}
      contentContainerStyle={centered ? styles.centered : undefined}
    >
      <Text style={[styles.text, style]}>{children}</Text>
    </ScrollView>
  );
}

const styles = StyleSheet.create({
  text: {
    flexShrink: 0
  },
  centered: {
    flexGrow: 1,
    justifyContent: 'center'
  }
});

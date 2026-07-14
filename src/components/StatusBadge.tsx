import { StyleSheet, Text, View } from 'react-native';

import { colors } from '@/theme/colors';

export function StatusBadge({ state }: { state: 'done' | 'progress' | 'todo' }) {
  const label = state === 'done' ? 'Fertig' : state === 'progress' ? 'In Arbeit' : 'Offen';
  const backgroundColor = state === 'done' ? colors.done : state === 'progress' ? colors.progress : colors.todo;

  return (
    <View style={[styles.badge, { backgroundColor }]}>
      <Text style={styles.text}>{label}</Text>
    </View>
  );
}

const styles = StyleSheet.create({
  badge: {
    borderRadius: 999,
    paddingHorizontal: 10,
    paddingVertical: 4,
  },
  text: {
    color: '#fff',
    fontSize: 12,
    fontWeight: '600',
  },
});

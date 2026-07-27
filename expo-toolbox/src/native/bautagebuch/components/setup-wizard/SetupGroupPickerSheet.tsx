import { Modal, Pressable, ScrollView, StyleSheet, Text, View } from 'react-native';
import { useSafeAreaInsets } from 'react-native-safe-area-context';

import { colors, spacing, typography } from '../../../../constants/theme';
import { systemBottomInset } from '../../../../navigation/systemInsets';

type GroupItem = {
  sectionId: string;
  label: string;
  count: number;
};

type Props = {
  visible: boolean;
  groups: GroupItem[];
  activeSectionId: string;
  onSelect: (sectionId: string) => void;
  onClose: () => void;
};

export function SetupGroupPickerSheet({
  visible,
  groups,
  activeSectionId,
  onSelect,
  onClose
}: Props) {
  const insets = useSafeAreaInsets();

  return (
    <Modal visible={visible} animationType="slide" transparent onRequestClose={onClose}>
      <Pressable style={styles.backdrop} onPress={onClose}>
        <Pressable
          style={[styles.sheet, { paddingBottom: systemBottomInset(insets) + spacing.sm }]}
          onPress={(event) => event.stopPropagation()}
        >
          <View style={styles.handle} />
          <Text style={styles.title}>Gruppen</Text>
          <ScrollView style={styles.list} keyboardShouldPersistTaps="handled">
            {groups.map((group) => {
              const active = group.sectionId === activeSectionId;
              return (
                <Pressable
                  key={group.sectionId}
                  style={[styles.row, active ? styles.rowActive : null]}
                  onPress={() => {
                    onSelect(group.sectionId);
                    onClose();
                  }}
                >
                  <Text style={styles.rowCheck}>{active ? '✓' : ' '}</Text>
                  <View style={styles.rowCopy}>
                    <Text style={[styles.rowLabel, active ? styles.rowLabelActive : null]}>
                      {group.label}
                    </Text>
                    <Text style={styles.rowMeta}>
                      {group.count} {group.count === 1 ? 'Feld' : 'Felder'}
                    </Text>
                  </View>
                </Pressable>
              );
            })}
          </ScrollView>
        </Pressable>
      </Pressable>
    </Modal>
  );
}

const styles = StyleSheet.create({
  backdrop: {
    flex: 1,
    justifyContent: 'flex-end',
    backgroundColor: 'rgba(26, 25, 22, 0.45)'
  },
  sheet: {
    maxHeight: '72%',
    borderTopLeftRadius: 18,
    borderTopRightRadius: 18,
    backgroundColor: colors.panel,
    borderWidth: 1,
    borderColor: colors.border,
    paddingTop: spacing.sm,
    paddingHorizontal: spacing.pageX
  },
  handle: {
    alignSelf: 'center',
    width: 40,
    height: 4,
    borderRadius: 999,
    backgroundColor: colors.border,
    marginBottom: spacing.sm
  },
  title: {
    ...typography.subtitle,
    color: colors.ink,
    marginBottom: spacing.sm
  },
  list: {
    flexGrow: 0
  },
  row: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: spacing.sm,
    minHeight: spacing.touchMin + 4,
    paddingVertical: spacing.sm,
    paddingHorizontal: spacing.sm,
    borderRadius: 12,
    marginBottom: spacing.xs
  },
  rowActive: {
    backgroundColor: colors.badgeBg
  },
  rowCheck: {
    width: 20,
    ...typography.bodyStrong,
    color: colors.accent,
    textAlign: 'center'
  },
  rowCopy: {
    flex: 1,
    gap: 2
  },
  rowLabel: {
    ...typography.bodyStrong,
    color: colors.ink
  },
  rowLabelActive: {
    color: colors.accent2
  },
  rowMeta: {
    ...typography.caption,
    color: colors.muted
  }
});

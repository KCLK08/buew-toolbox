import { useEffect, useRef, useState } from 'react';
import { Modal, Pressable, ScrollView, StyleSheet, Text, View } from 'react-native';
import { useSafeAreaInsets } from 'react-native-safe-area-context';

import { PrimaryButton, TextField } from '../../../../components/mobile';
import { colors, spacing, typography } from '../../../../constants/theme';
import { systemBottomInset } from '../../../../navigation/systemInsets';

type Props = {
  visible: boolean;
  initialName?: string;
  initialDescription?: string;
  readOnly?: boolean;
  onClose: () => void;
  onSave: (input: { name: string; description?: string }) => void;
};

export function SetupStructureGroupModal({
  visible,
  initialName = '',
  initialDescription = '',
  readOnly = false,
  onClose,
  onSave
}: Props) {
  const insets = useSafeAreaInsets();
  const [name, setName] = useState(initialName);
  const [description, setDescription] = useState(initialDescription);
  const wasVisibleRef = useRef(false);

  useEffect(() => {
    if (visible && !wasVisibleRef.current) {
      setName(initialName);
      setDescription(initialDescription);
    }
    wasVisibleRef.current = visible;
  }, [visible, initialName, initialDescription]);

  const submit = () => {
    onSave({
      name: name.trim(),
      description: description.trim() || undefined
    });
  };

  return (
    <Modal visible={visible} animationType="slide" onRequestClose={onClose}>
      <View style={[styles.root, { paddingTop: insets.top }]}>
        <View style={styles.header}>
          <Pressable accessibilityRole="button" onPress={onClose}>
            <Text style={styles.cancel}>Abbrechen</Text>
          </Pressable>
          <Text style={styles.title}>{initialName ? 'Gruppe bearbeiten' : 'Gruppe hinzufügen'}</Text>
          <View style={styles.headerSpacer} />
        </View>

        <ScrollView
          style={styles.body}
          contentContainerStyle={[
            styles.bodyContent,
            { paddingBottom: systemBottomInset(insets) + spacing.xl }
          ]}
          keyboardShouldPersistTaps="handled"
        >
          <TextField
            label="Name"
            value={name}
            onChangeText={setName}
            placeholder="z. B. Allgemeine Angaben"
            editable={!readOnly}
            autoCapitalize="sentences"
          />
          <TextField
            label="Beschreibung (optional)"
            value={description}
            onChangeText={setDescription}
            placeholder="z. B. Projektbezogene Informationen"
            editable={!readOnly}
            multiline
          />

          <View style={styles.typeCard}>
            <Text style={styles.typeLabel}>Typ</Text>
            <Text style={styles.typeValue}>Formulargruppe</Text>
          </View>
        </ScrollView>

        {!readOnly ? (
          <View style={[styles.footer, { paddingBottom: systemBottomInset(insets) + spacing.sm }]}>
            <PrimaryButton label="Speichern" onPress={submit} disabled={!name.trim()} />
          </View>
        ) : null}
      </View>
    </Modal>
  );
}

const styles = StyleSheet.create({
  root: {
    flex: 1,
    backgroundColor: colors.bg
  },
  header: {
    flexDirection: 'row',
    alignItems: 'center',
    justifyContent: 'space-between',
    paddingHorizontal: spacing.pageX,
    paddingVertical: spacing.sm,
    borderBottomWidth: 1,
    borderBottomColor: colors.border,
    backgroundColor: colors.panel
  },
  cancel: {
    ...typography.bodyStrong,
    color: colors.accent,
    minWidth: 88
  },
  title: {
    ...typography.subtitle,
    color: colors.ink
  },
  headerSpacer: {
    minWidth: 88
  },
  body: {
    flex: 1
  },
  bodyContent: {
    padding: spacing.pageX,
    gap: spacing.md
  },
  typeCard: {
    borderRadius: spacing.cardRadius,
    borderWidth: 1,
    borderColor: colors.border,
    backgroundColor: colors.panelElevated,
    padding: spacing.md,
    gap: spacing.xxs
  },
  typeLabel: {
    ...typography.caption,
    color: colors.muted
  },
  typeValue: {
    ...typography.bodyStrong,
    color: colors.ink
  },
  footer: {
    paddingHorizontal: spacing.pageX,
    paddingTop: spacing.sm,
    borderTopWidth: 1,
    borderTopColor: colors.border,
    backgroundColor: colors.panel
  }
});

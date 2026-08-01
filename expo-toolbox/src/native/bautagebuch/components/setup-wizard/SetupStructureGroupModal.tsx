import { useEffect, useRef, useState } from 'react';
import { Modal, Pressable, StyleSheet, Text, View } from 'react-native';
import { MaterialCommunityIcons } from '@expo/vector-icons';
import { useSafeAreaInsets } from 'react-native-safe-area-context';

import { PrimaryButton, TextField } from '../../../../components/mobile';
import { colors, spacing, typography } from '../../../../constants/theme';
import { systemBottomInset } from '../../../../navigation/systemInsets';
import { SetupModalKeyboardFrame } from './SetupModalKeyboardFrame';
import { SetupScrollView } from './SetupScrollView';

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
  const isEditing = Boolean(initialName);

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
      <SetupModalKeyboardFrame>
        <View style={[styles.root, { paddingTop: insets.top }]}>
        <View style={styles.header}>
          <Pressable accessibilityRole="button" style={styles.headerBtn} onPress={onClose}>
            <Text style={styles.cancel}>Abbrechen</Text>
          </Pressable>
          <Text style={styles.title}>{isEditing ? 'Gruppe bearbeiten' : 'Gruppe hinzufügen'}</Text>
          <View style={styles.headerBtn} />
        </View>

        <SetupScrollView
          style={styles.body}
          contentContainerStyle={[
            styles.bodyContent,
            { paddingBottom: systemBottomInset(insets) + spacing.xl }
          ]}
        >
          <View style={styles.hero}>
            <View style={styles.heroIconWrap}>
              <MaterialCommunityIcons name="file-document-outline" size={28} color={colors.accent2} />
            </View>
            <Text style={styles.heroTitle}>Formulargruppe</Text>
            <Text style={styles.heroCopy}>
              Ein Bereich für zusammengehörige Felder wie Projektinfos oder Wetter.
            </Text>
          </View>

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
        </SetupScrollView>

        {!readOnly ? (
          <View style={[styles.footer, { paddingBottom: systemBottomInset(insets) + spacing.sm }]}>
            <PrimaryButton label="Speichern" onPress={submit} disabled={!name.trim()} />
          </View>
        ) : null}
        </View>
      </SetupModalKeyboardFrame>
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
  headerBtn: {
    minWidth: 88
  },
  cancel: {
    ...typography.bodyStrong,
    color: colors.accent
  },
  title: {
    ...typography.subtitle,
    color: colors.ink
  },
  body: {
    flex: 1
  },
  bodyContent: {
    padding: spacing.pageX,
    gap: spacing.md
  },
  hero: {
    alignItems: 'center',
    gap: spacing.xs,
    paddingVertical: spacing.sm
  },
  heroIconWrap: {
    width: 56,
    height: 56,
    borderRadius: 16,
    alignItems: 'center',
    justifyContent: 'center',
    backgroundColor: 'rgba(36, 50, 64, 0.1)'
  },
  heroTitle: {
    ...typography.subtitle,
    color: colors.ink
  },
  heroCopy: {
    ...typography.body,
    color: colors.muted,
    textAlign: 'center',
    lineHeight: 22
  },
  footer: {
    paddingHorizontal: spacing.pageX,
    paddingTop: spacing.sm,
    borderTopWidth: 1,
    borderTopColor: colors.border,
    backgroundColor: colors.panel
  }
});

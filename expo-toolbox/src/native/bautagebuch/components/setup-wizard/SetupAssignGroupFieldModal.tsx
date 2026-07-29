import { useEffect, useRef, useState } from 'react';
import { Modal, Pressable, ScrollView, StyleSheet, Text, View } from 'react-native';
import { MaterialCommunityIcons } from '@expo/vector-icons';
import { useSafeAreaInsets } from 'react-native-safe-area-context';

import { PrimaryButton, TextField } from '../../../../components/mobile';
import { colors, spacing, typography } from '../../../../constants/theme';
import { systemBottomInset } from '../../../../navigation/systemInsets';

type Props = {
  visible: boolean;
  groupName: string;
  initialFieldName: string;
  readOnly?: boolean;
  onClose: () => void;
  onConfirm: (fieldName: string) => void;
};

export function SetupAssignGroupFieldModal({
  visible,
  groupName,
  initialFieldName,
  readOnly = false,
  onClose,
  onConfirm
}: Props) {
  const insets = useSafeAreaInsets();
  const [fieldName, setFieldName] = useState(initialFieldName);
  const wasVisibleRef = useRef(false);

  useEffect(() => {
    if (visible && !wasVisibleRef.current) {
      setFieldName(initialFieldName);
    }
    wasVisibleRef.current = visible;
  }, [visible, initialFieldName]);

  const submit = () => {
    onConfirm(fieldName.trim() || initialFieldName);
  };

  return (
    <Modal visible={visible} animationType="slide" onRequestClose={onClose}>
      <View style={[styles.root, { paddingTop: insets.top }]}>
        <View style={styles.header}>
          <Pressable accessibilityRole="button" style={styles.headerBtn} onPress={onClose}>
            <Text style={styles.cancel}>Abbrechen</Text>
          </Pressable>
          <Text style={styles.title}>Feld erstellen</Text>
          <View style={styles.headerBtn} />
        </View>

        <ScrollView
          style={styles.body}
          contentContainerStyle={[
            styles.bodyContent,
            { paddingBottom: systemBottomInset(insets) + spacing.xl }
          ]}
          keyboardShouldPersistTaps="handled"
        >
          <View style={styles.hero}>
            <View style={styles.heroIconWrap}>
              <MaterialCommunityIcons name="file-document-outline" size={28} color={colors.accent2} />
            </View>
            <Text style={styles.heroTitle}>Gruppe</Text>
            <Text style={styles.groupName}>{groupName}</Text>
          </View>

          <TextField
            label="Feldname"
            value={fieldName}
            onChangeText={setFieldName}
            placeholder="Name des Feldes"
            editable={!readOnly}
            autoCapitalize="sentences"
            autoFocus
          />
        </ScrollView>

        {!readOnly ? (
          <View style={[styles.footer, { paddingBottom: systemBottomInset(insets) + spacing.sm }]}>
            <PrimaryButton label="Übernehmen" onPress={submit} disabled={!fieldName.trim()} />
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
    gap: spacing.xxs,
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
    ...typography.caption,
    color: colors.muted
  },
  groupName: {
    ...typography.subtitle,
    color: colors.ink,
    textAlign: 'center'
  },
  footer: {
    paddingHorizontal: spacing.pageX,
    paddingTop: spacing.sm,
    borderTopWidth: 1,
    borderTopColor: colors.border,
    backgroundColor: colors.panel
  }
});

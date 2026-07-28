import { useEffect, useState } from 'react';
import { Alert, Modal, Pressable, StyleSheet, Text, View } from 'react-native';

import { PrimaryButton, SingleLineText, TextField } from '../../../../components/mobile';
import { colors, spacing, typography } from '../../../../constants/theme';
import { renameTemplate } from '../../services/templateService';

type Props = {
  templateId: string;
  templateName: string;
  readOnly?: boolean;
  onRenamed: (nextName: string) => void;
  variant?: 'summary' | 'title';
};

export function SetupTemplateRenameControl({
  templateId,
  templateName,
  readOnly = false,
  onRenamed,
  variant = 'summary'
}: Props) {
  const [open, setOpen] = useState(false);
  const [draft, setDraft] = useState(templateName);
  const [saving, setSaving] = useState(false);

  useEffect(() => {
    setDraft(templateName);
  }, [templateName]);

  const openEditor = () => {
    setDraft(templateName);
    setOpen(true);
  };

  const submit = async () => {
    const trimmed = draft.trim();
    if (!trimmed) {
      Alert.alert('Umbenennen', 'Bitte einen Namen eingeben.');
      return;
    }
    setSaving(true);
    try {
      await renameTemplate(templateId, trimmed);
      onRenamed(trimmed);
      setOpen(false);
    } catch (err) {
      Alert.alert('Umbenennen', err instanceof Error ? err.message : 'Umbenennen fehlgeschlagen.');
    } finally {
      setSaving(false);
    }
  };

  if (variant === 'title') {
    return (
      <>
        <View style={styles.titleRow}>
          <Text style={styles.titleName}>{templateName}</Text>
          {!readOnly ? (
            <Pressable accessibilityRole="button" style={styles.linkBtn} onPress={openEditor}>
              <Text style={styles.linkLabel}>Umbenennen</Text>
            </Pressable>
          ) : null}
        </View>
        <RenameModal
          visible={open}
          draft={draft}
          saving={saving}
          onChangeDraft={setDraft}
          onClose={() => setOpen(false)}
          onSubmit={() => void submit()}
        />
      </>
    );
  }

  return (
    <>
      <View style={styles.summaryValueWrap}>
        <SingleLineText style={styles.summaryValue}>{templateName}</SingleLineText>
        {!readOnly ? (
          <Pressable accessibilityRole="button" style={styles.linkBtn} onPress={openEditor}>
            <Text style={styles.linkLabel}>Umbenennen</Text>
          </Pressable>
        ) : null}
      </View>
      <RenameModal
        visible={open}
        draft={draft}
        saving={saving}
        onChangeDraft={setDraft}
        onClose={() => setOpen(false)}
        onSubmit={() => void submit()}
      />
    </>
  );
}

function RenameModal({
  visible,
  draft,
  saving,
  onChangeDraft,
  onClose,
  onSubmit
}: {
  visible: boolean;
  draft: string;
  saving: boolean;
  onChangeDraft: (value: string) => void;
  onClose: () => void;
  onSubmit: () => void;
}) {
  return (
    <Modal visible={visible} transparent animationType="fade">
      <View style={styles.modalBackdrop}>
        <View style={styles.modalCard}>
          <Text style={styles.modalTitle}>Vorlage umbenennen</Text>
          <TextField label="Name" value={draft} onChangeText={onChangeDraft} autoCapitalize="sentences" />
          <View style={styles.modalActions}>
            <PrimaryButton label="Abbrechen" variant="ghost" onPress={onClose} disabled={saving} />
            <PrimaryButton label={saving ? 'Speichern…' : 'Speichern'} onPress={onSubmit} disabled={saving} />
          </View>
        </View>
      </View>
    </Modal>
  );
}

const styles = StyleSheet.create({
  titleRow: {
    gap: spacing.xxs
  },
  titleName: {
    ...typography.title,
    color: colors.ink
  },
  summaryValueWrap: {
    flex: 1,
    minWidth: 0,
    alignItems: 'flex-end',
    gap: 2
  },
  summaryValue: {
    ...typography.bodyStrong,
    color: colors.ink,
    textAlign: 'right'
  },
  linkBtn: {
    paddingVertical: 2
  },
  linkLabel: {
    ...typography.caption,
    color: colors.accent,
    fontFamily: 'SpaceGrotesk_600SemiBold',
    textAlign: 'right'
  },
  modalBackdrop: {
    flex: 1,
    backgroundColor: colors.overlay,
    justifyContent: 'center',
    padding: spacing.xl
  },
  modalCard: {
    backgroundColor: colors.panelElevated,
    borderRadius: spacing.cardRadius,
    padding: spacing.md,
    gap: spacing.sm,
    borderWidth: 1,
    borderColor: colors.border
  },
  modalTitle: {
    ...typography.subtitle,
    color: colors.ink
  },
  modalActions: {
    flexDirection: 'row',
    justifyContent: 'flex-end',
    gap: spacing.xs
  }
});

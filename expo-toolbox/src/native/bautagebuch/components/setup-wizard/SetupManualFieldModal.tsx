import { useEffect, useMemo, useState } from 'react';
import { Modal, Pressable, ScrollView, StyleSheet, Text, TextInput, View } from 'react-native';
import { MaterialCommunityIcons } from '@expo/vector-icons';
import { useSafeAreaInsets } from 'react-native-safe-area-context';

import { PrimaryButton } from '../../../../components/mobile';
import { colors, spacing, typography } from '../../../../constants/theme';
import { SETUP_FIELD_TYPE_OPTIONS } from '../../lib/setup-field-settings';
import { getStructureItems } from '../../lib/setup-structure';
import type { FieldRect, SetupFieldType, SetupStructureItem } from '../../types';

type Props = {
  visible: boolean;
  page: number;
  rect: FieldRect;
  setupModel: Record<string, unknown>;
  readOnly?: boolean;
  onClose: () => void;
  onConfirm: (input: {
    name: string;
    type: SetupFieldType;
    target: { kind: 'group'; id: string } | { kind: 'table'; id: string } | null;
  }) => void;
};

export function SetupManualFieldModal({
  visible,
  page,
  rect,
  setupModel,
  readOnly = false,
  onClose,
  onConfirm
}: Props) {
  const insets = useSafeAreaInsets();
  const [name, setName] = useState('');
  const [type, setType] = useState<SetupFieldType>('text');
  const [targetId, setTargetId] = useState<string | null>(null);

  const structure = useMemo(() => getStructureItems(setupModel), [setupModel]);

  useEffect(() => {
    if (!visible) return;
    setName('');
    setType('text');
    setTargetId(structure[0]?.id || null);
  }, [visible, structure]);

  const selectTarget = (item: SetupStructureItem) => {
    setTargetId(item.id);
  };

  const handleConfirm = () => {
    const trimmed = name.trim();
    if (!trimmed) return;
    const target = targetId
      ? structure.find((item) => item.id === targetId)
      : null;
    onConfirm({
      name: trimmed,
      type,
      target: target
        ? target.type === 'group'
          ? { kind: 'group', id: target.id }
          : { kind: 'table', id: target.id }
        : null
    });
  };

  return (
    <Modal visible={visible} animationType="slide" transparent onRequestClose={onClose}>
      <View style={styles.backdrop}>
        <View style={[styles.sheet, { paddingBottom: insets.bottom + spacing.md }]}>
          <Text style={styles.title}>Neues Feld erstellen</Text>
          <Text style={styles.subtitle}>
            Seite {page} · Bereich {Math.round(rect.width)}×{Math.round(rect.height)} pt
          </Text>

          <Text style={styles.label}>Feldname</Text>
          <TextInput
            value={name}
            onChangeText={setName}
            placeholder="z. B. Unterschrift Bauüberwachung"
            style={styles.input}
            editable={!readOnly}
          />

          <Text style={styles.label}>Feldtyp</Text>
          <ScrollView horizontal showsHorizontalScrollIndicator={false} style={styles.typeRow}>
            {SETUP_FIELD_TYPE_OPTIONS.filter((option) => option.value !== 'table').map((option) => {
              const active = option.value === type;
              return (
                <Pressable
                  key={option.value}
                  style={[styles.typeChip, active ? styles.typeChipActive : null]}
                  onPress={() => setType(option.value)}
                  disabled={readOnly}
                >
                  <Text style={active ? styles.typeChipLabelActive : styles.typeChipLabel}>
                    {option.label}
                  </Text>
                </Pressable>
              );
            })}
          </ScrollView>

          <Text style={styles.label}>Gruppe oder Tabelle (optional)</Text>
          <View style={styles.targetList}>
            {structure.map((item) => {
              const active = targetId === item.id;
              return (
                <Pressable
                  key={item.id}
                  style={[styles.targetRow, active ? styles.targetRowActive : null]}
                  onPress={() => selectTarget(item)}
                  disabled={readOnly}
                >
                  <MaterialCommunityIcons
                    name={item.type === 'group' ? 'folder-outline' : 'table'}
                    size={18}
                    color={active ? colors.accent : colors.muted}
                  />
                  <Text style={active ? styles.targetLabelActive : styles.targetLabel}>{item.name}</Text>
                </Pressable>
              );
            })}
          </View>

          <View style={styles.actions}>
            <PrimaryButton label="Abbrechen" variant="ghost" onPress={onClose} />
            <PrimaryButton
              label="Feld speichern"
              disabled={readOnly || !name.trim()}
              onPress={handleConfirm}
            />
          </View>
        </View>
      </View>
    </Modal>
  );
}

const styles = StyleSheet.create({
  backdrop: {
    flex: 1,
    justifyContent: 'flex-end',
    backgroundColor: 'rgba(26, 25, 22, 0.35)'
  },
  sheet: {
    backgroundColor: colors.panel,
    borderTopLeftRadius: 18,
    borderTopRightRadius: 18,
    paddingHorizontal: spacing.pageX,
    paddingTop: spacing.md,
    gap: spacing.sm,
    maxHeight: '88%'
  },
  title: {
    ...typography.subtitle,
    color: colors.ink
  },
  subtitle: {
    ...typography.caption,
    color: colors.muted
  },
  label: {
    ...typography.label,
    color: colors.muted,
    marginTop: spacing.xxs
  },
  input: {
    ...typography.body,
    color: colors.ink,
    borderWidth: 1,
    borderColor: colors.border,
    borderRadius: 12,
    paddingHorizontal: spacing.md,
    minHeight: spacing.touchMin,
    backgroundColor: colors.panelElevated
  },
  typeRow: {
    flexGrow: 0
  },
  typeChip: {
    paddingHorizontal: spacing.sm,
    paddingVertical: spacing.xs,
    borderRadius: 999,
    borderWidth: 1,
    borderColor: colors.border,
    marginRight: spacing.xs,
    backgroundColor: colors.panelElevated
  },
  typeChipActive: {
    borderColor: colors.accent,
    backgroundColor: colors.badgeBg
  },
  typeChipLabel: {
    ...typography.caption,
    color: colors.muted
  },
  typeChipLabelActive: {
    ...typography.caption,
    color: colors.accent2,
    fontFamily: 'SpaceGrotesk_600SemiBold'
  },
  targetList: {
    gap: spacing.xxs,
    maxHeight: 160
  },
  targetRow: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: spacing.sm,
    minHeight: spacing.touchMin,
    paddingHorizontal: spacing.sm,
    borderRadius: 10,
    borderWidth: 1,
    borderColor: colors.border
  },
  targetRowActive: {
    borderColor: colors.accent,
    backgroundColor: colors.badgeBg
  },
  targetLabel: {
    ...typography.body,
    color: colors.ink
  },
  targetLabelActive: {
    ...typography.bodyStrong,
    color: colors.accent2
  },
  actions: {
    flexDirection: 'row',
    justifyContent: 'flex-end',
    gap: spacing.sm,
    marginTop: spacing.sm
  }
});

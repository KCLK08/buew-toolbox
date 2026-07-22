import { useMemo, useState } from 'react';
import { Pressable, ScrollView, StyleSheet, Text, TextInput, View } from 'react-native';

import { PrimaryButton } from '../../../../components/mobile';
import { colors, shadows, spacing, typography } from '../../../../constants/theme';
import type { OverlayPlacement } from '../../lib/setup-mapping';
import type { SetupWizardGroup } from '../../types';

type Props = {
  groups: SetupWizardGroup[];
  placement: OverlayPlacement;
  onSelectGroup: (sectionId: string) => void;
  onCreateGroup: (label: string) => void;
  disabled?: boolean;
};

export function GroupOverlayCards({
  groups,
  placement,
  onSelectGroup,
  onCreateGroup,
  disabled = false
}: Props) {
  const [creating, setCreating] = useState(false);
  const [newGroupLabel, setNewGroupLabel] = useState('');

  const containerStyle = useMemo(() => {
    if (placement === 'top') return styles.overlayTop;
    if (placement === 'left') return styles.overlayLeft;
    if (placement === 'right') return styles.overlayRight;
    return styles.overlayBottom;
  }, [placement]);

  const handleCreate = () => {
    const label = newGroupLabel.trim();
    if (!label) return;
    onCreateGroup(label);
    setNewGroupLabel('');
    setCreating(false);
  };

  return (
    <View style={[styles.host, containerStyle]} pointerEvents="box-none">
      <View style={[styles.panel, shadows.card]} pointerEvents="auto">
        <Text style={styles.heading}>Gruppe wählen</Text>
        <ScrollView
          horizontal={placement === 'bottom' || placement === 'top'}
          showsHorizontalScrollIndicator={false}
          contentContainerStyle={styles.cardRow}
          keyboardShouldPersistTaps="handled"
        >
          {groups.map((group) => (
            <Pressable
              key={group.sectionId}
              disabled={disabled}
              style={({ pressed }) => [styles.card, pressed ? styles.cardPressed : null]}
              onPress={() => onSelectGroup(group.sectionId)}
            >
              <Text style={styles.cardLabel}>{group.label}</Text>
            </Pressable>
          ))}
          <Pressable
            style={({ pressed }) => [styles.card, styles.createCard, pressed ? styles.cardPressed : null]}
            onPress={() => setCreating((value) => !value)}
          >
            <Text style={styles.createLabel}>+ Neue Gruppe</Text>
          </Pressable>
        </ScrollView>

        {creating ? (
          <View style={styles.createRow}>
            <TextInput
              value={newGroupLabel}
              onChangeText={setNewGroupLabel}
              placeholder="Gruppenname"
              placeholderTextColor={colors.muted}
              style={styles.input}
              autoFocus
            />
            <PrimaryButton label="Hinzufügen" onPress={handleCreate} />
          </View>
        ) : null}
      </View>
    </View>
  );
}

const styles = StyleSheet.create({
  host: {
    position: 'absolute',
    zIndex: 20
  },
  overlayTop: {
    top: spacing.sm,
    left: spacing.pageX,
    right: spacing.pageX
  },
  overlayBottom: {
    bottom: spacing.sm,
    left: spacing.pageX,
    right: spacing.pageX
  },
  overlayLeft: {
    left: spacing.pageX,
    top: '28%',
    width: 180
  },
  overlayRight: {
    right: spacing.pageX,
    top: '28%',
    width: 180
  },
  panel: {
    backgroundColor: colors.panelElevated,
    borderRadius: spacing.cardRadius,
    borderWidth: 1,
    borderColor: colors.border,
    padding: spacing.sm,
    gap: spacing.sm
  },
  heading: {
    ...typography.label,
    color: colors.muted
  },
  cardRow: {
    gap: spacing.sm,
    paddingRight: spacing.xs
  },
  card: {
    minWidth: 132,
    minHeight: spacing.touchMin,
    borderRadius: 14,
    backgroundColor: colors.panel,
    borderWidth: 1,
    borderColor: colors.border,
    paddingHorizontal: spacing.md,
    paddingVertical: spacing.sm,
    justifyContent: 'center'
  },
  cardPressed: {
    backgroundColor: colors.badgeBg,
    borderColor: colors.accent
  },
  cardLabel: {
    ...typography.bodyStrong,
    color: colors.ink
  },
  createCard: {
    borderStyle: 'dashed',
    borderColor: colors.accent
  },
  createLabel: {
    ...typography.bodyStrong,
    color: colors.accent
  },
  createRow: {
    gap: spacing.sm
  },
  input: {
    ...typography.body,
    color: colors.ink,
    borderWidth: 1,
    borderColor: colors.border,
    borderRadius: spacing.inputRadius,
    paddingHorizontal: spacing.md,
    paddingVertical: spacing.sm,
    backgroundColor: colors.panel
  }
});

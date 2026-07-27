import { useMemo, useState } from 'react';
import { Pressable, ScrollView, StyleSheet, Text, TextInput, View } from 'react-native';

import { PrimaryButton } from '../../../../components/mobile';
import { colors, shadows, spacing, typography } from '../../../../constants/theme';
import type { OverlayPlacement } from '../../lib/setup-mapping';
import type { SetupWizardGroup } from '../../types';

type Props = {
  groups: SetupWizardGroup[];
  placement: OverlayPlacement;
  selectedGroupId?: string | null;
  onSelectGroup: (sectionId: string) => void;
  onCreateGroup: (label: string) => void;
  disabled?: boolean;
};

export function GroupOverlayCards({
  groups,
  placement,
  selectedGroupId = null,
  onSelectGroup,
  onCreateGroup,
  disabled = false
}: Props) {
  const [creating, setCreating] = useState(false);
  const [newGroupLabel, setNewGroupLabel] = useState('');
  const hasGroups = groups.length > 0;

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

  const openCreateDialog = () => {
    setCreating(true);
  };

  if (!hasGroups) {
    return (
      <View style={[styles.host, styles.overlayBottom]} pointerEvents="box-none">
        <View style={[styles.panel, styles.panelEmpty, shadows.card]} pointerEvents="auto">
          <Text style={styles.emptyTitle}>Gruppen erforderlich</Text>
          <Text style={styles.emptyCopy}>
            Für die Zuordnung müssen zuerst Gruppen erstellt werden.
          </Text>
          {creating ? (
            <View style={styles.createRow}>
              <Text style={styles.createLabel}>Name der Gruppe</Text>
              <TextInput
                value={newGroupLabel}
                onChangeText={setNewGroupLabel}
                placeholder="z. B. Kopfdaten"
                placeholderTextColor={colors.muted}
                style={styles.input}
                autoFocus
                returnKeyType="done"
                onSubmitEditing={handleCreate}
              />
              <PrimaryButton label="Speichern" onPress={handleCreate} />
            </View>
          ) : (
            <PrimaryButton label="+ Neue Gruppe" onPress={openCreateDialog} />
          )}
        </View>
      </View>
    );
  }

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
          {groups.map((group) => {
            const selected = selectedGroupId === group.sectionId;
            return (
              <Pressable
                key={group.sectionId}
                disabled={disabled}
                style={({ pressed }) => [
                  styles.card,
                  selected ? styles.cardSelected : null,
                  pressed ? styles.cardPressed : null
                ]}
                onPress={() => onSelectGroup(group.sectionId)}
              >
                <Text style={[styles.cardLabel, selected ? styles.cardLabelSelected : null]}>
                  {group.label}
                </Text>
              </Pressable>
            );
          })}
          <Pressable
            style={({ pressed }) => [styles.card, styles.createCard, pressed ? styles.cardPressed : null]}
            onPress={openCreateDialog}
          >
            <Text style={styles.createLabel}>+ Neue Gruppe</Text>
          </Pressable>
        </ScrollView>

        {creating ? (
          <View style={styles.createRow}>
            <Text style={styles.createLabel}>Name der Gruppe</Text>
            <TextInput
              value={newGroupLabel}
              onChangeText={setNewGroupLabel}
              placeholder="Gruppenname"
              placeholderTextColor={colors.muted}
              style={styles.input}
              autoFocus
              returnKeyType="done"
              onSubmitEditing={handleCreate}
            />
            <PrimaryButton label="Speichern" onPress={handleCreate} />
          </View>
        ) : null}
      </View>
    </View>
  );
}

const styles = StyleSheet.create({
  host: {
    position: 'absolute',
    zIndex: 20,
    maxWidth: '100%'
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
    top: '22%',
    maxWidth: 200
  },
  overlayRight: {
    right: spacing.pageX,
    top: '22%',
    maxWidth: 200
  },
  panel: {
    backgroundColor: colors.panelElevated,
    borderRadius: spacing.cardRadius,
    borderWidth: 1,
    borderColor: colors.border,
    padding: spacing.md,
    gap: spacing.sm
  },
  panelEmpty: {
    gap: spacing.md
  },
  emptyTitle: {
    ...typography.bodyStrong,
    color: colors.ink
  },
  emptyCopy: {
    ...typography.body,
    color: colors.muted
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
    minWidth: 140,
    minHeight: spacing.touchMin + 4,
    borderRadius: 14,
    backgroundColor: colors.panel,
    borderWidth: 1,
    borderColor: colors.border,
    paddingHorizontal: spacing.md,
    paddingVertical: spacing.sm,
    justifyContent: 'center'
  },
  cardSelected: {
    backgroundColor: colors.badgeBg,
    borderColor: colors.accent,
    borderWidth: 2
  },
  cardPressed: {
    backgroundColor: colors.badgeBg,
    borderColor: colors.accent
  },
  cardLabel: {
    ...typography.bodyStrong,
    color: colors.ink,
    textAlign: 'center'
  },
  cardLabelSelected: {
    color: colors.accent2
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
    minHeight: spacing.touchMin,
    paddingVertical: spacing.sm,
    backgroundColor: colors.panel
  }
});

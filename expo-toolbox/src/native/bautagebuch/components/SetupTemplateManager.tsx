import { ActivityIndicator, Pressable, ScrollView, StyleSheet, Text, View } from 'react-native';

import { PrimaryButton, StatusBadge } from '../../../components/mobile';
import { colors, spacing, typography } from '../../../constants/theme';
import type { BautagebuchTemplate } from '../types';

type Props = {
  templates: BautagebuchTemplate[];
  activeTemplateId: string;
  editingTemplateId: string;
  importing?: boolean;
  onSelectEdit: (templateId: string) => void;
  onSetActive: (templateId: string) => void;
  onImport: () => void;
};

export function SetupTemplateManager({
  templates,
  activeTemplateId,
  editingTemplateId,
  importing = false,
  onSelectEdit,
  onSetActive,
  onImport
}: Props) {
  return (
    <View style={styles.root}>
      <View style={styles.headerRow}>
        <Text style={styles.title}>Vorlagen</Text>
        <PrimaryButton
          label={importing ? 'Import…' : '+ PDF hinzufügen'}
          variant="secondary"
          disabled={importing}
          onPress={onImport}
        />
      </View>
      <Text style={styles.hint}>
        Die aktive Vorlage wird für neue BTBs verwendet. Tippe eine Vorlage an, um sie im Setup zu bearbeiten.
      </Text>

      <ScrollView horizontal showsHorizontalScrollIndicator={false} contentContainerStyle={styles.cardRow}>
        {templates.map((template) => {
          const isActive = template.templateId === activeTemplateId;
          const isEditing = template.templateId === editingTemplateId;
          return (
            <Pressable
              key={template.templateId}
              style={[styles.card, isEditing ? styles.cardEditing : null]}
              onPress={() => onSelectEdit(template.templateId)}
            >
              <View style={styles.cardTop}>
                <Text style={styles.cardTitle} numberOfLines={2}>
                  {template.templateName}
                </Text>
                {isActive ? <StatusBadge label="Aktiv" tone="success" /> : null}
              </View>
              <Text style={styles.cardMeta} numberOfLines={1}>
                {template.fileName}
              </Text>
              <Text style={styles.cardMeta}>
                {template.status === 'ready' ? 'Startbereit' : 'Setup offen'} · {template.pageCount} Seite(n)
              </Text>
              {!isActive ? (
                <Pressable
                  style={styles.activateBtn}
                  onPress={(event) => {
                    event.stopPropagation?.();
                    onSetActive(template.templateId);
                  }}
                >
                  <Text style={styles.activateLabel}>Als aktiv setzen</Text>
                </Pressable>
              ) : (
                <Text style={styles.activeNote}>Wird für neue BTBs genutzt</Text>
              )}
            </Pressable>
          );
        })}

        {importing ? (
          <View style={[styles.card, styles.importCard]}>
            <ActivityIndicator color={colors.accent} />
            <Text style={styles.cardMeta}>PDF wird eingelesen…</Text>
          </View>
        ) : null}
      </ScrollView>
    </View>
  );
}

const styles = StyleSheet.create({
  root: {
    gap: spacing.xs
  },
  headerRow: {
    flexDirection: 'row',
    alignItems: 'center',
    justifyContent: 'space-between',
    gap: spacing.sm
  },
  title: {
    ...typography.bodyStrong,
    color: colors.ink
  },
  hint: {
    ...typography.caption,
    color: colors.muted
  },
  cardRow: {
    gap: spacing.sm,
    paddingRight: spacing.sm
  },
  card: {
    width: 220,
    gap: spacing.xs,
    padding: spacing.sm,
    borderRadius: 12,
    borderWidth: 1,
    borderColor: colors.border,
    backgroundColor: colors.panel
  },
  cardEditing: {
    borderColor: colors.accent,
    backgroundColor: colors.badgeBg
  },
  cardTop: {
    flexDirection: 'row',
    alignItems: 'flex-start',
    justifyContent: 'space-between',
    gap: spacing.xs
  },
  cardTitle: {
    ...typography.bodyStrong,
    color: colors.ink,
    flex: 1
  },
  cardMeta: {
    ...typography.caption,
    color: colors.muted
  },
  activateBtn: {
    marginTop: spacing.xxs,
    alignSelf: 'flex-start',
    paddingVertical: 6,
    paddingHorizontal: spacing.xs,
    borderRadius: 8,
    backgroundColor: colors.panelElevated,
    borderWidth: 1,
    borderColor: colors.border
  },
  activateLabel: {
    ...typography.caption,
    color: colors.accent,
    fontFamily: 'SpaceGrotesk_600SemiBold'
  },
  activeNote: {
    ...typography.caption,
    color: colors.success
  },
  importCard: {
    alignItems: 'center',
    justifyContent: 'center'
  }
});

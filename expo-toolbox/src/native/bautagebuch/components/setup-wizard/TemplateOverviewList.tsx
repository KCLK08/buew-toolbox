import { ActivityIndicator, Pressable, StyleSheet, Text, View } from 'react-native';
import { MaterialCommunityIcons } from '@expo/vector-icons';

import { PrimaryButton, SingleLineText, StatusBadge } from '../../../../components/mobile';
import { colors, shadows, spacing, typography } from '../../../../constants/theme';
import { templateDisplayStatus } from '../../lib/setup-mapping';
import type { BautagebuchTemplate } from '../../types';

type Props = {
  templates: BautagebuchTemplate[];
  activeTemplateId: string;
  importing?: boolean;
  onImport: () => void;
  onOpen: (templateId: string) => void;
  onContinueSetup: (templateId: string) => void;
  onActivate: (templateId: string) => void;
  onRename?: (templateId: string) => void;
  onArchive?: (templateId: string) => void;
  onDelete?: (templateId: string) => void;
  canDelete?: (template: BautagebuchTemplate) => boolean;
};

function actionForTemplate(
  template: BautagebuchTemplate,
  isActive: boolean
): { label: string; variant: 'primary' | 'secondary' | 'ghost'; action: 'open' | 'activate' | 'continue' | 'view' } {
  if (template.status === 'archived') {
    return { label: 'Nur ansehen', variant: 'ghost', action: 'view' };
  }
  if (isActive && template.status === 'ready') {
    return { label: 'Öffnen', variant: 'secondary', action: 'open' };
  }
  if (template.status === 'ready') {
    return { label: 'Aktivieren', variant: 'primary', action: 'activate' };
  }
  if (template.status === 'draft' || template.status === 'in_progress') {
    return { label: 'Setup fortsetzen', variant: 'primary', action: 'continue' };
  }
  return { label: 'Öffnen', variant: 'secondary', action: 'open' };
}

export function TemplateOverviewList({
  templates,
  activeTemplateId,
  importing = false,
  onImport,
  onOpen,
  onContinueSetup,
  onActivate,
  onRename,
  onArchive,
  onDelete,
  canDelete
}: Props) {
  return (
    <View style={styles.root}>
      <View style={styles.heroCard}>
        <View style={styles.heroIconWrap}>
          <MaterialCommunityIcons name="file-document-edit-outline" size={28} color={colors.accent} />
        </View>
        <View style={styles.heroCopy}>
          <Text style={styles.heroTitle}>PDF-Vorlagen</Text>
          <Text style={styles.heroText}>
            Importiere Vorlagen, richte Felder ein und lege fest, welche Vorlage für neue
            Bautagebücher verwendet wird.
          </Text>
        </View>
        <PrimaryButton
          label={importing ? 'Import läuft…' : 'PDF importieren'}
          variant="primary"
          disabled={importing}
          onPress={onImport}
        />
      </View>

      {importing ? (
        <View style={styles.importBanner}>
          <ActivityIndicator color={colors.accent} />
          <Text style={styles.importText}>PDF wird analysiert und Felder werden erkannt…</Text>
        </View>
      ) : null}

      <Text style={styles.sectionTitle}>
        {templates.length} Vorlage{templates.length === 1 ? '' : 'n'}
      </Text>

      <View style={styles.cardList}>
        {templates.map((template) => {
          const isActive = template.templateId === activeTemplateId;
          const status = templateDisplayStatus(template.status, isActive);
          const action = actionForTemplate(template, isActive);

          const handleAction = () => {
            if (action.action === 'activate') onActivate(template.templateId);
            else if (action.action === 'continue') onContinueSetup(template.templateId);
            else onOpen(template.templateId);
          };

          return (
            <View
              key={template.templateId}
              style={[styles.templateCard, isActive ? styles.templateCardActive : null, shadows.card]}
            >
              <View style={styles.cardTop}>
                <View style={styles.cardTitleBlock}>
                  <SingleLineText style={styles.templateName}>{template.templateName}</SingleLineText>
                  <SingleLineText style={styles.fileName}>{template.fileName}</SingleLineText>
                </View>
                <StatusBadge label={status.label} tone={status.tone} />
              </View>

              <View style={styles.metaRow}>
                <View style={styles.metaItem}>
                  <MaterialCommunityIcons name="file-pdf-box" size={16} color={colors.muted} />
                  <Text style={styles.metaText}>{template.pageCount} Seite(n)</Text>
                </View>
                {isActive ? (
                  <View style={styles.metaItem}>
                    <MaterialCommunityIcons name="check-circle" size={16} color={colors.success} />
                    <Text style={[styles.metaText, styles.metaTextActive]}>Aktiv für neue BTBs</Text>
                  </View>
                ) : null}
              </View>

              <View style={styles.actionRow}>
                <PrimaryButton label={action.label} variant={action.variant} onPress={handleAction} />
                {onRename ? (
                  <Pressable style={styles.linkBtn} onPress={() => onRename(template.templateId)}>
                    <Text style={styles.linkLabel}>Umbenennen</Text>
                  </Pressable>
                ) : null}
                {onArchive && template.status !== 'archived' && !isActive ? (
                  <Pressable style={styles.linkBtn} onPress={() => onArchive(template.templateId)}>
                    <Text style={styles.archiveLabel}>Archivieren</Text>
                  </Pressable>
                ) : null}
                {onDelete && canDelete?.(template) ? (
                  <Pressable style={styles.linkBtn} onPress={() => onDelete(template.templateId)}>
                    <Text style={styles.deleteLabel}>Löschen</Text>
                  </Pressable>
                ) : null}
              </View>
            </View>
          );
        })}
      </View>
    </View>
  );
}

const styles = StyleSheet.create({
  root: {
    gap: spacing.md,
    paddingHorizontal: spacing.pageX,
    paddingBottom: spacing.xl
  },
  heroCard: {
    gap: spacing.md,
    padding: spacing.md,
    borderRadius: spacing.cardRadius,
    backgroundColor: colors.panelElevated,
    borderWidth: 1,
    borderColor: colors.border,
    ...shadows.card
  },
  heroIconWrap: {
    width: 48,
    height: 48,
    borderRadius: spacing.iconRadius,
    alignItems: 'center',
    justifyContent: 'center',
    backgroundColor: colors.badgeBg
  },
  heroCopy: {
    gap: spacing.xxs
  },
  heroTitle: {
    ...typography.subtitle,
    color: colors.ink
  },
  heroText: {
    ...typography.body,
    color: colors.muted
  },
  importBanner: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: spacing.sm,
    padding: spacing.md,
    borderRadius: 12,
    backgroundColor: colors.badgeBg,
    borderWidth: 1,
    borderColor: colors.border
  },
  importText: {
    ...typography.caption,
    color: colors.ink,
    flex: 1
  },
  sectionTitle: {
    ...typography.label,
    color: colors.muted,
    marginTop: spacing.xxs
  },
  cardList: {
    gap: spacing.cardGap
  },
  templateCard: {
    gap: spacing.sm,
    padding: spacing.cardPadding,
    borderRadius: spacing.cardRadius,
    backgroundColor: colors.panelElevated,
    borderWidth: 1,
    borderColor: colors.border
  },
  templateCardActive: {
    borderColor: colors.accent,
    backgroundColor: colors.panel
  },
  cardTop: {
    flexDirection: 'row',
    alignItems: 'flex-start',
    justifyContent: 'space-between',
    gap: spacing.sm
  },
  cardTitleBlock: {
    flex: 1,
    gap: 2,
    minWidth: 0
  },
  templateName: {
    ...typography.bodyStrong,
    color: colors.ink,
    fontSize: 17
  },
  fileName: {
    ...typography.caption,
    color: colors.muted
  },
  metaRow: {
    flexDirection: 'row',
    flexWrap: 'wrap',
    gap: spacing.sm
  },
  metaItem: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: spacing.xxs
  },
  metaText: {
    ...typography.caption,
    color: colors.muted
  },
  metaTextActive: {
    color: colors.success,
    fontFamily: 'SpaceGrotesk_600SemiBold'
  },
  actionRow: {
    flexDirection: 'row',
    flexWrap: 'wrap',
    alignItems: 'center',
    gap: spacing.sm,
    paddingTop: spacing.xxs
  },
  linkBtn: {
    paddingVertical: spacing.xxs
  },
  linkLabel: {
    ...typography.caption,
    color: colors.accent,
    fontFamily: 'SpaceGrotesk_600SemiBold'
  },
  archiveLabel: {
    ...typography.caption,
    color: colors.muted
  },
  deleteLabel: {
    ...typography.caption,
    color: colors.danger,
    fontFamily: 'SpaceGrotesk_600SemiBold'
  }
});

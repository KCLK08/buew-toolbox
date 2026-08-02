import { useState } from 'react';
import { ActivityIndicator, LayoutAnimation, Platform, Pressable, StyleSheet, Text, UIManager, View } from 'react-native';
import { MaterialCommunityIcons } from '@expo/vector-icons';

import { PrimaryButton } from '../../../../components/mobile';
import { colors, spacing, typography } from '../../../../constants/theme';
import type { BautagebuchTemplate } from '../../types';
import { TemplateOverviewCard } from './TemplateOverviewCard';

if (Platform.OS === 'android' && UIManager.setLayoutAnimationEnabledExperimental) {
  UIManager.setLayoutAnimationEnabledExperimental(true);
}

type Props = {
  templates: BautagebuchTemplate[];
  activeTemplateId: string;
  importing?: boolean;
  onImport: () => void;
  onOpen: (templateId: string) => void;
  onContinueSetup: (templateId: string) => void;
  onActivate: (templateId: string) => void;
  onArchive?: (templateId: string) => void;
  onDelete?: (templateId: string) => void;
  canArchive?: (template: BautagebuchTemplate) => boolean;
  canDelete?: (template: BautagebuchTemplate) => boolean;
};

function actionForTemplate(
  template: BautagebuchTemplate,
  isActive: boolean
): { label: string; variant: 'primary' | 'secondary' | 'ghost'; action: 'open' | 'activate' | 'continue' | 'view' } {
  if (template.status === 'archived') {
    return { label: 'Nur ansehen', variant: 'ghost', action: 'view' };
  }
  if (template.status === 'draft' || template.status === 'in_progress') {
    return { label: 'Setup fortsetzen', variant: 'primary', action: 'continue' };
  }
  if (template.status === 'ready') {
    return { label: 'Öffnen', variant: isActive ? 'secondary' : 'primary', action: 'open' };
  }
  return { label: 'Öffnen', variant: 'secondary', action: 'open' };
}

function TemplateCards({
  templates,
  activeTemplateId,
  onOpen,
  onContinueSetup,
  onActivate,
  onArchive,
  onDelete,
  canArchive,
  canDelete
}: Omit<Props, 'importing' | 'onImport'>) {
  return (
    <View style={styles.cardList}>
      {templates.map((template) => {
        const isActive = template.templateId === activeTemplateId;
        const action = actionForTemplate(template, isActive);
        const handleAction = () => {
          if (action.action === 'activate') onActivate(template.templateId);
          else if (action.action === 'continue') onContinueSetup(template.templateId);
          else onOpen(template.templateId);
        };

        return (
          <TemplateOverviewCard
            key={template.templateId}
            template={template}
            isActive={isActive}
            action={action}
            onAction={handleAction}
            onActivate={
              template.status === 'ready' && !isActive
                ? () => onActivate(template.templateId)
                : undefined
            }
            onArchive={onArchive ? () => onArchive(template.templateId) : undefined}
            onDelete={onDelete ? () => onDelete(template.templateId) : undefined}
            canArchive={canArchive?.(template) ?? false}
            canDelete={canDelete?.(template) ?? false}
          />
        );
      })}
    </View>
  );
}

export function TemplateOverviewList({
  templates,
  activeTemplateId,
  importing = false,
  onImport,
  onOpen,
  onContinueSetup,
  onActivate,
  onArchive,
  onDelete,
  canArchive,
  canDelete
}: Props) {
  const [archivedExpanded, setArchivedExpanded] = useState(false);
  const activeTemplates = templates.filter((template) => template.status !== 'archived');
  const archivedTemplates = templates.filter((template) => template.status === 'archived');

  const toggleArchived = () => {
    LayoutAnimation.configureNext(LayoutAnimation.Presets.easeInEaseOut);
    setArchivedExpanded((current) => !current);
  };

  const sharedCardProps = {
    activeTemplateId,
    onOpen,
    onContinueSetup,
    onActivate,
    onArchive,
    onDelete,
    canArchive,
    canDelete
  };

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
        {activeTemplates.length} Vorlage{activeTemplates.length === 1 ? '' : 'n'}
      </Text>

      {activeTemplates.length > 0 ? (
        <TemplateCards templates={activeTemplates} {...sharedCardProps} />
      ) : (
        <Text style={styles.emptyCopy}>Keine aktiven Vorlagen vorhanden.</Text>
      )}

      {archivedTemplates.length > 0 ? (
        <View style={styles.archivedSection}>
          <Pressable
            accessibilityRole="button"
            style={styles.archivedHeader}
            onPress={toggleArchived}
          >
            <View style={styles.archivedHeaderCopy}>
              <MaterialCommunityIcons name="archive-outline" size={18} color={colors.muted} />
              <Text style={styles.archivedTitle}>
                Archivierte Vorlagen ({archivedTemplates.length})
              </Text>
            </View>
            <MaterialCommunityIcons
              name={archivedExpanded ? 'chevron-up' : 'chevron-down'}
              size={22}
              color={colors.muted}
            />
          </Pressable>
          {archivedExpanded ? (
            <TemplateCards templates={archivedTemplates} {...sharedCardProps} />
          ) : null}
        </View>
      ) : null}
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
    borderColor: colors.border
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
  emptyCopy: {
    ...typography.body,
    color: colors.muted
  },
  cardList: {
    gap: spacing.cardGap
  },
  archivedSection: {
    gap: spacing.sm,
    marginTop: spacing.sm,
    paddingTop: spacing.sm,
    borderTopWidth: 1,
    borderTopColor: colors.border
  },
  archivedHeader: {
    flexDirection: 'row',
    alignItems: 'center',
    justifyContent: 'space-between',
    minHeight: 44,
    paddingHorizontal: spacing.xs
  },
  archivedHeaderCopy: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: spacing.xs
  },
  archivedTitle: {
    ...typography.bodyStrong,
    color: colors.ink
  }
});

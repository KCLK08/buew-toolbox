import { ActivityIndicator, Pressable, StyleSheet, Text, View } from 'react-native';

import { PrimaryButton, StatusBadge } from '../../../../components/mobile';
import { colors, spacing, typography } from '../../../../constants/theme';
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
  onArchive?: (templateId: string) => void;
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
  onArchive
}: Props) {
  return (
    <View style={styles.root}>
      <View style={styles.headerRow}>
        <Text style={styles.title}>Vorlagen</Text>
        <PrimaryButton
          label={importing ? 'Import…' : '+ PDF importieren'}
          variant="secondary"
          disabled={importing}
          onPress={onImport}
        />
      </View>

      <View style={styles.tableHead}>
        <Text style={[styles.headCell, styles.nameCol]}>Template</Text>
        <Text style={[styles.headCell, styles.statusCol]}>Status</Text>
        <Text style={[styles.headCell, styles.actionCol]}>Aktionen</Text>
      </View>

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
          <View key={template.templateId} style={styles.row}>
            <View style={styles.nameCol}>
              <Text style={styles.templateName}>{template.templateName}</Text>
              <Text style={styles.fileName} numberOfLines={1}>
                {template.fileName}
              </Text>
            </View>
            <View style={styles.statusCol}>
              <StatusBadge label={status.label} tone={status.tone} />
            </View>
            <View style={styles.actionCol}>
              <PrimaryButton label={action.label} variant={action.variant} onPress={handleAction} />
              {onArchive && template.status !== 'archived' && !isActive ? (
                <Pressable onPress={() => onArchive(template.templateId)}>
                  <Text style={styles.archiveLink}>Archivieren</Text>
                </Pressable>
              ) : null}
            </View>
          </View>
        );
      })}

      {importing ? (
        <View style={styles.importRow}>
          <ActivityIndicator color={colors.accent} />
          <Text style={styles.importText}>PDF wird analysiert…</Text>
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
  headerRow: {
    flexDirection: 'row',
    alignItems: 'center',
    justifyContent: 'space-between',
    gap: spacing.sm
  },
  title: {
    ...typography.title,
    color: colors.ink
  },
  tableHead: {
    flexDirection: 'row',
    gap: spacing.sm,
    paddingBottom: spacing.xs,
    borderBottomWidth: 1,
    borderBottomColor: colors.border
  },
  headCell: {
    ...typography.label,
    color: colors.muted
  },
  row: {
    flexDirection: 'row',
    gap: spacing.sm,
    alignItems: 'center',
    paddingVertical: spacing.md,
    borderBottomWidth: 1,
    borderBottomColor: colors.border
  },
  nameCol: {
    flex: 1.4,
    gap: 2
  },
  statusCol: {
    flex: 0.8
  },
  actionCol: {
    flex: 1,
    gap: spacing.xs,
    alignItems: 'flex-start'
  },
  templateName: {
    ...typography.bodyStrong,
    color: colors.ink
  },
  fileName: {
    ...typography.caption,
    color: colors.muted
  },
  archiveLink: {
    ...typography.caption,
    color: colors.muted,
    textDecorationLine: 'underline'
  },
  importRow: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: spacing.sm,
    paddingVertical: spacing.md
  },
  importText: {
    ...typography.caption,
    color: colors.muted
  }
});

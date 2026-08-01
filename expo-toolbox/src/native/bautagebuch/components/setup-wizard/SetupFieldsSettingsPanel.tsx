import { Pressable, StyleSheet, Text, View } from 'react-native';
import { MaterialCommunityIcons } from '@expo/vector-icons';

import { PrimaryButton } from '../../../../components/mobile';
import { colors, shadows, spacing, typography } from '../../../../constants/theme';
import { hapticSelection } from '../../../../lib/haptics';
import {
  normalizeSetupFieldType,
  resolveFieldFromTarget,
  resolveHybridFieldSource,
  resolveTargetDisplayName,
  resolveTargetGroupLabel,
  setupFieldTypeLabel
} from '../../lib/setup-field-settings';
import { type MappingField } from '../../lib/setup-mapping';
import { fieldSourceLabel, fieldSourceTone } from '../../lib/template-field';
import type { DetectedField, FieldSettingsTarget } from '../../types';
import { SetupFieldSettingsForm } from './SetupFieldSettingsForm';
import { SetupFieldStepNavigator } from './SetupFieldStepNavigator';
import { SetupScrollView } from './SetupScrollView';
import type { SetupFieldType } from '../../types';

type Props = {
  setupModel: Record<string, unknown>;
  target: FieldSettingsTarget | null;
  currentTargetIndex: number;
  totalTargets: number;
  fieldNumber: number;
  detectedFields: DetectedField[];
  mappingFields?: MappingField[];
  draftLabels?: Record<string, string>;
  readOnly?: boolean;
  onChange: (next: Record<string, unknown>) => void;
  onFieldLabelChange?: (fieldId: string, label: string) => void;
  onOpenTypePicker: () => void;
  onSaveAndNext: () => void;
  onSelectTarget: (index: number) => void;
  showSaveButton?: boolean;
  bottomInset?: number;
};

export function SetupFieldsSettingsPanel({
  setupModel,
  target,
  currentTargetIndex,
  totalTargets,
  fieldNumber,
  detectedFields,
  mappingFields = [],
  draftLabels = {},
  readOnly = false,
  onChange,
  onFieldLabelChange,
  onOpenTypePicker,
  onSaveAndNext,
  onSelectTarget,
  showSaveButton = true,
  bottomInset = 0
}: Props) {
  if (!target) {
    return (
      <View style={styles.empty}>
        <Text style={styles.emptyTitle}>Alle Felder konfiguriert</Text>
        <Text style={styles.emptyCopy}>
          Du kannst die Einstellungen jederzeit über „Alle Felder“ erneut öffnen.
        </Text>
      </View>
    );
  }

  const field = resolveFieldFromTarget(setupModel, target);
  const displayName = resolveTargetDisplayName(
    setupModel,
    target,
    field,
    mappingFields,
    draftLabels
  );
  const groupLabel = resolveTargetGroupLabel(setupModel, target);
  const hybridSource = resolveHybridFieldSource(field, detectedFields);
  const sourceLabel = hybridSource ? fieldSourceLabel(hybridSource) : null;
  const sourceTone = hybridSource ? fieldSourceTone(hybridSource) : null;
  const fieldType: SetupFieldType =
    target.kind === 'table-meta' ? 'table' : normalizeSetupFieldType(field, detectedFields);

  return (
    <SetupScrollView
      style={styles.scroll}
      contentContainerStyle={[styles.content, { paddingBottom: spacing.xl + bottomInset }]}
      showsVerticalScrollIndicator={false}
    >
      <SetupFieldStepNavigator
        fieldNumber={fieldNumber}
        total={totalTargets}
        canGoPrevious={currentTargetIndex > 0}
        canGoNext={currentTargetIndex < totalTargets - 1}
        onPrevious={() => onSelectTarget(currentTargetIndex - 1)}
        onNext={() => onSelectTarget(currentTargetIndex + 1)}
      />

      {displayName ? <Text style={styles.fieldName}>{displayName}</Text> : null}

      <View style={styles.section}>
        {sourceLabel ? (
          <View style={styles.infoCard}>
            <Text style={styles.infoLabel}>Quelle</Text>
            <Text
              style={[
                styles.infoValue,
                sourceTone === 'success'
                  ? styles.sourceAuto
                  : sourceTone === 'warning'
                    ? styles.sourceManual
                    : null
              ]}
            >
              {sourceLabel}
            </Text>
          </View>
        ) : null}
        <View style={styles.infoCard}>
          <Text style={styles.infoLabel}>{target.kind === 'table-meta' ? 'Tabelle' : 'Gruppe'}</Text>
          <Text style={styles.infoValue}>{groupLabel}</Text>
        </View>
      </View>

      {target.kind !== 'table-meta' ? (
        <View style={styles.section}>
          <Text style={styles.sectionTitle}>Feldtyp</Text>
          <Pressable
            accessibilityRole="button"
            style={[styles.typePicker, shadows.card]}
            disabled={readOnly}
            onPress={() => {
              void hapticSelection();
              onOpenTypePicker();
            }}
          >
            <View style={styles.typeCopy}>
              <Text style={styles.typeHint}>Aktuell</Text>
              <Text style={styles.typeValue}>{setupFieldTypeLabel(fieldType)}</Text>
            </View>
            <MaterialCommunityIcons name="chevron-down" size={22} color={colors.muted} />
          </Pressable>
        </View>
      ) : null}

      <View style={styles.section}>
        <Text style={styles.sectionTitle}>
          {target.kind === 'table-meta' ? 'Tabelleneinstellungen' : 'Feldeinstellungen'}
        </Text>
        <SetupFieldSettingsForm
          setupModel={setupModel}
          target={target}
          field={field}
          detectedFields={detectedFields}
          mappingFields={mappingFields}
          draftLabels={draftLabels}
          readOnly={readOnly}
          onChange={onChange}
          onFieldLabelChange={onFieldLabelChange}
        />
      </View>

      {showSaveButton && !readOnly ? (
        <PrimaryButton label="Speichern & weiter" onPress={onSaveAndNext} />
      ) : null}
    </SetupScrollView>
  );
}

const styles = StyleSheet.create({
  scroll: {
    flex: 1
  },
  content: {
    padding: spacing.pageX,
    paddingBottom: spacing.xl,
    gap: spacing.lg
  },
  fieldName: {
    ...typography.title,
    color: colors.ink
  },
  section: {
    gap: spacing.sm
  },
  sectionTitle: {
    ...typography.subtitle,
    color: colors.ink
  },
  infoCard: {
    gap: 2,
    padding: spacing.md,
    borderRadius: 12,
    borderWidth: 1,
    borderColor: colors.border,
    backgroundColor: colors.panelElevated
  },
  infoLabel: {
    ...typography.caption,
    color: colors.muted
  },
  infoValue: {
    ...typography.bodyStrong,
    color: colors.ink
  },
  sourceAuto: {
    color: colors.success
  },
  sourceManual: {
    color: colors.warning
  },
  typePicker: {
    flexDirection: 'row',
    alignItems: 'center',
    justifyContent: 'space-between',
    minHeight: spacing.touchMin + 8,
    paddingHorizontal: spacing.md,
    paddingVertical: spacing.sm,
    borderRadius: 14,
    borderWidth: 1,
    borderColor: colors.border,
    backgroundColor: colors.panel
  },
  typeCopy: {
    gap: 2
  },
  typeHint: {
    ...typography.caption,
    color: colors.muted
  },
  typeValue: {
    ...typography.bodyStrong,
    color: colors.accent2
  },
  empty: {
    flex: 1,
    alignItems: 'center',
    justifyContent: 'center',
    padding: spacing.pageX,
    gap: spacing.sm
  },
  emptyTitle: {
    ...typography.subtitle,
    color: colors.ink,
    textAlign: 'center'
  },
  emptyCopy: {
    ...typography.body,
    color: colors.muted,
    textAlign: 'center',
    lineHeight: 22
  }
});

import { useEffect, useMemo, useRef, useState } from 'react';
import { Pressable, StyleSheet, Text, View } from 'react-native';
import { useSafeAreaInsets } from 'react-native-safe-area-context';

import { PrimaryButton } from '../../../../components/mobile';
import { colors, spacing, typography } from '../../../../constants/theme';
import { systemBottomInset } from '../../../../navigation/systemInsets';
import {
  assignFieldToGroup,
  assignFieldToTableCell,
  addWizardGroup,
  checkMappingTransition,
  deferField,
  getMappingCompletionSummary,
  getMappingProgress,
  getNextUnassignedIndex,
  getWizardState,
  isMappingComplete,
  resolveCurrentMappingIndex,
  resolveOverlayPlacement,
  type MappingField
} from '../../lib/setup-mapping';
import type { SetupWizardTableAssignment } from '../../types';
import type { DetectedField } from '../../types';
import { SetupPdfFieldPreview } from '../SetupPdfFieldPreview';
import { GroupOverlayCards } from './GroupOverlayCards';
import { TableMappingOverlay } from './TableMappingOverlay';
import { SetupMappingCompletion } from './SetupMappingCompletion';
import { SetupMappingValidation } from './SetupMappingValidation';
import { SetupProgressHeader } from './SetupProgressHeader';
import { SetupTemplateRenameControl } from './SetupTemplateRenameControl';

type Props = {
  templateId: string;
  templateName: string;
  pdfPath: string | null;
  detectedFields: DetectedField[];
  mappingFields: MappingField[];
  setupModel: Record<string, unknown>;
  readOnly?: boolean;
  onChange: (next: Record<string, unknown>) => void;
  onComplete: (nextModel: Record<string, unknown>) => void;
  onFinishLater: () => void;
  onTemplateRenamed: (nextName: string) => void;
};

export function SetupMappingStep({
  templateId,
  templateName,
  pdfPath,
  detectedFields,
  mappingFields,
  setupModel,
  readOnly = false,
  onChange,
  onComplete,
  onFinishLater,
  onTemplateRenamed
}: Props) {
  const insets = useSafeAreaInsets();
  const indexSyncedRef = useRef(false);
  const wizard = useMemo(() => getWizardState(setupModel), [setupModel]);
  const [mappingMode, setMappingMode] = useState<'group' | 'table'>('group');
  const [selectedGroupId, setSelectedGroupId] = useState<string | null>(null);
  const [selectedTableAssignment, setSelectedTableAssignment] =
    useState<SetupWizardTableAssignment | null>(null);
  const [activeTableId, setActiveTableId] = useState<string | null>(wizard.tables[0]?.tableId || null);
  const [showCompletion, setShowCompletion] = useState(false);
  const [showValidation, setShowValidation] = useState(false);

  const currentIndex = resolveCurrentMappingIndex(mappingFields, wizard);
  const currentField = mappingFields[currentIndex] || null;
  const progress = useMemo(
    () => getMappingProgress(mappingFields, wizard),
    [mappingFields, wizard]
  );
  const placement = resolveOverlayPlacement(currentField?.rect || null);
  const mappingDone = isMappingComplete(mappingFields, wizard);
  const hasGroups = wizard.groups.length > 0;
  const fieldNumber = currentField ? currentIndex + 1 : 0;
  const completionSummary = useMemo(
    () => getMappingCompletionSummary(setupModel, mappingFields),
    [setupModel, mappingFields]
  );
  const transitionCheck = useMemo(
    () => checkMappingTransition(setupModel, mappingFields),
    [setupModel, mappingFields]
  );

  useEffect(() => {
    if (indexSyncedRef.current || mappingFields.length === 0) return;
    indexSyncedRef.current = true;
    const resolved = resolveCurrentMappingIndex(mappingFields, wizard);
    if (resolved !== wizard.currentFieldIndex) {
      onChange({
        ...setupModel,
        wizard: {
          ...wizard,
          currentFieldIndex: resolved
        }
      });
    }
  }, [mappingFields, wizard, setupModel, onChange]);

  useEffect(() => {
    if (mappingDone) {
      setShowCompletion(true);
    }
  }, [mappingDone]);

  const advanceAfterChange = (next: Record<string, unknown>) => {
    const nextWizard = getWizardState(next);
    setSelectedGroupId(null);

    if (isMappingComplete(mappingFields, nextWizard)) {
      onChange(next);
      setShowCompletion(true);
      return;
    }

    const nextIndex = getNextUnassignedIndex(mappingFields, nextWizard, 0);
    onChange({
      ...next,
      wizard: {
        ...nextWizard,
        currentFieldIndex: nextIndex >= 0 ? nextIndex : nextWizard.currentFieldIndex
      }
    });
  };

  const assignGroup = (sectionId: string) => {
    if (!currentField || !hasGroups) return;
    setSelectedGroupId(sectionId);
    setSelectedTableAssignment(null);
    const next = assignFieldToGroup(setupModel, currentField.fieldId, sectionId);
    advanceAfterChange(next);
  };

  const assignTableCell = (assignment: SetupWizardTableAssignment) => {
    if (!currentField) return;
    setSelectedTableAssignment(assignment);
    setSelectedGroupId(null);
    const next = assignFieldToTableCell(setupModel, currentField.fieldId, assignment);
    advanceAfterChange(next);
  };

  const createGroup = (label: string) => {
    const result = addWizardGroup(setupModel, label);
    setSelectedGroupId(result.group.sectionId);
    onChange(result.setupModel);
  };

  const goBack = () => {
    setSelectedGroupId(null);
    onChange({
      ...setupModel,
      wizard: {
        ...wizard,
        currentFieldIndex: Math.max(0, currentIndex - 1)
      }
    });
  };

  const goForward = () => {
    setSelectedGroupId(null);
    onChange({
      ...setupModel,
      wizard: {
        ...wizard,
        currentFieldIndex: Math.min(mappingFields.length - 1, currentIndex + 1)
      }
    });
  };

  const skipField = () => {
    if (!currentField) return;
    const next = deferField(setupModel, currentField.fieldId);
    advanceAfterChange(next);
  };

  const proceedToFields = () => {
    if (transitionCheck.issues.length > 0) {
      setShowValidation(true);
      return;
    }
    onComplete(setupModel);
  };

  const footerBottom = systemBottomInset(insets);

  if (showCompletion && mappingDone) {
    return (
      <View style={styles.root}>
        <SetupMappingCompletion
          templateId={templateId}
          templateName={templateName}
          readOnly={readOnly}
          summary={completionSummary}
          onConfigureFields={proceedToFields}
          onFinishLater={onFinishLater}
          onTemplateRenamed={onTemplateRenamed}
        />
        {showValidation ? (
          <SetupMappingValidation
            check={transitionCheck}
            onBackToMapping={() => {
              setShowValidation(false);
              setShowCompletion(false);
            }}
            onContinueAnyway={() => {
              setShowValidation(false);
              onComplete(setupModel);
            }}
          />
        ) : null}
      </View>
    );
  }

  return (
    <View style={styles.root}>
      <View style={styles.templateBar}>
        <SetupTemplateRenameControl
          templateId={templateId}
          templateName={templateName}
          readOnly={readOnly}
          onRenamed={onTemplateRenamed}
          variant="title"
        />
      </View>
      <SetupProgressHeader
        progress={progress}
        fieldNumber={fieldNumber}
        fieldName={currentField?.fieldName}
        labelCandidate={currentField?.labelCandidate}
      />

      <View style={styles.modeRow}>
        <Pressable
          style={[styles.modeChip, mappingMode === 'group' ? styles.modeChipActive : null]}
          onPress={() => setMappingMode('group')}
        >
          <Text style={mappingMode === 'group' ? styles.modeChipLabelActive : styles.modeChipLabel}>Gruppe</Text>
        </Pressable>
        <Pressable
          style={[styles.modeChip, mappingMode === 'table' ? styles.modeChipActive : null]}
          onPress={() => setMappingMode('table')}
        >
          <Text style={mappingMode === 'table' ? styles.modeChipLabelActive : styles.modeChipLabel}>Tabelle</Text>
        </Pressable>
      </View>

      <View style={styles.canvas}>
        <SetupPdfFieldPreview
          pdfPath={pdfPath}
          detectedFields={detectedFields}
          activeFieldId={currentField?.fieldId || null}
          activeFieldLabel={currentField?.labelCandidate || null}
          activeFieldPage={currentField?.page || 1}
          variant="mapping"
        />

        {currentField && mappingMode === 'group' ? (
          <GroupOverlayCards
            groups={wizard.groups}
            placement={placement}
            selectedGroupId={selectedGroupId}
            onSelectGroup={assignGroup}
            onCreateGroup={createGroup}
            disabled={!hasGroups}
          />
        ) : null}
        {currentField && mappingMode === 'table' ? (
          <TableMappingOverlay
            wizard={wizard}
            currentField={currentField}
            placement={placement}
            selectedAssignment={selectedTableAssignment}
            activeTableId={activeTableId}
            setupModel={setupModel}
            onChange={onChange}
            onAssignCell={assignTableCell}
            onSelectTable={setActiveTableId}
          />
        ) : null}
      </View>

      <View style={[styles.footer, { paddingBottom: footerBottom }]}>
        <View style={styles.navRow}>
          <Pressable
            style={[styles.navBtn, currentIndex <= 0 ? styles.navBtnDisabled : null]}
            onPress={goBack}
            disabled={currentIndex <= 0}
          >
            <Text style={[styles.navLabel, currentIndex <= 0 ? styles.navLabelDisabled : null]}>
              Zurück
            </Text>
          </Pressable>
          <Pressable style={styles.navBtnPrimary} onPress={skipField} disabled={!currentField}>
            <Text style={styles.navLabelPrimary}>Überspringen</Text>
          </Pressable>
          <Pressable
            style={[
              styles.navBtn,
              currentIndex >= mappingFields.length - 1 ? styles.navBtnDisabled : null
            ]}
            onPress={goForward}
            disabled={currentIndex >= mappingFields.length - 1}
          >
            <Text
              style={[
                styles.navLabel,
                currentIndex >= mappingFields.length - 1 ? styles.navLabelDisabled : null
              ]}
            >
              Weiter
            </Text>
          </Pressable>
        </View>

        {mappingDone ? (
          <PrimaryButton label="Felder konfigurieren" onPress={proceedToFields} />
        ) : (
          <PrimaryButton label="Später fortsetzen" variant="ghost" onPress={onFinishLater} />
        )}
      </View>
    </View>
  );
}

const styles = StyleSheet.create({
  root: {
    flex: 1
  },
  templateBar: {
    paddingHorizontal: spacing.pageX,
    paddingTop: spacing.sm,
    paddingBottom: spacing.xxs,
    backgroundColor: colors.panel,
    borderBottomWidth: 1,
    borderBottomColor: colors.border
  },
  modeRow: {
    flexDirection: 'row',
    gap: spacing.xs,
    paddingHorizontal: spacing.pageX,
    paddingVertical: spacing.xs,
    backgroundColor: colors.panel,
    borderBottomWidth: 1,
    borderBottomColor: colors.border
  },
  modeChip: {
    minHeight: 34,
    paddingHorizontal: spacing.md,
    borderRadius: 999,
    borderWidth: 1,
    borderColor: colors.border,
    justifyContent: 'center',
    backgroundColor: colors.panelElevated
  },
  modeChipActive: {
    borderColor: colors.accent,
    backgroundColor: colors.badgeBg
  },
  modeChipLabel: {
    ...typography.caption,
    color: colors.ink
  },
  modeChipLabelActive: {
    ...typography.caption,
    color: colors.accent2,
    fontFamily: 'SpaceGrotesk_600SemiBold'
  },
  canvas: {
    flex: 1,
    position: 'relative',
    minHeight: 0
  },
  footer: {
    gap: spacing.sm,
    paddingHorizontal: spacing.pageX,
    paddingTop: spacing.md,
    borderTopWidth: 1,
    borderTopColor: colors.border,
    backgroundColor: colors.panel
  },
  navRow: {
    flexDirection: 'row',
    justifyContent: 'space-between',
    gap: spacing.sm
  },
  navBtn: {
    flex: 1,
    minHeight: spacing.touchMin + 4,
    borderRadius: 12,
    borderWidth: 1,
    borderColor: colors.border,
    backgroundColor: colors.panel,
    justifyContent: 'center',
    alignItems: 'center',
    paddingHorizontal: spacing.sm
  },
  navBtnPrimary: {
    flex: 1.4,
    minHeight: spacing.touchMin + 4,
    borderRadius: 12,
    backgroundColor: colors.accent,
    justifyContent: 'center',
    alignItems: 'center',
    paddingHorizontal: spacing.sm
  },
  navBtnDisabled: {
    opacity: 0.45
  },
  navLabel: {
    ...typography.bodyStrong,
    color: colors.accent2
  },
  navLabelPrimary: {
    ...typography.bodyStrong,
    color: colors.panel
  },
  navLabelDisabled: {
    color: colors.muted
  }
});

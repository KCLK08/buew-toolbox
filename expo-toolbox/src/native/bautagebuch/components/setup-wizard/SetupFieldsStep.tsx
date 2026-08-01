import { useEffect, useMemo, useRef, useState } from 'react';
import { LayoutAnimation, Platform, StyleSheet, UIManager, View } from 'react-native';

import { PrimaryButton } from '../../../../components/mobile';
import { colors, spacing } from '../../../../constants/theme';
import { debounce } from '../../../../lib/debounce';
import { hapticSuccess } from '../../../../lib/haptics';
import {
  advanceFieldSettingsWalkthrough,
  applyFieldTypeChange,
  isFieldSettingsWalkthroughComplete,
  listFieldSettingsTargets,
  normalizeSetupFieldType,
  resolveCurrentFieldSettingsIndex,
  resolveFieldDisplayOrder,
  resolveFieldFromTarget,
  resolveFieldPreviewPage,
  resolveTargetDisplayName,
  updateFieldSettingsTarget
} from '../../lib/setup-field-settings';
import { getWizardState, sortMappingFields, withWizardState } from '../../lib/setup-mapping';
import type { DetectedField, SetupFieldType } from '../../types';
import { SetupPdfFieldPreview } from '../SetupPdfFieldPreview';
import { SetupFieldTypePickerSheet } from './SetupFieldTypePickerSheet';
import { SetupFieldsHeader, type SetupFieldsViewTab } from './SetupFieldsHeader';
import { SetupFieldsOverview } from './SetupFieldsOverview';
import { SetupFieldsSettingsPanel } from './SetupFieldsSettingsPanel';

if (Platform.OS === 'android' && UIManager.setLayoutAnimationEnabledExperimental) {
  UIManager.setLayoutAnimationEnabledExperimental(true);
}

type Props = {
  pdfPath: string | null;
  detectedFields: DetectedField[];
  setupModel: Record<string, unknown>;
  readOnly?: boolean;
  showWizardNav?: boolean;
  onChange: (next: Record<string, unknown>) => void;
  onFinish: () => void;
  onBack: () => void;
  onUpdateField?: (
    fieldId: string,
    patch: { labelCandidate?: string; type?: string }
  ) => Promise<void>;
};

export function SetupFieldsStep({
  pdfPath,
  detectedFields,
  setupModel,
  readOnly = false,
  showWizardNav = false,
  onChange,
  onFinish,
  onBack,
  onUpdateField
}: Props) {
  const [activeTab, setActiveTab] = useState<SetupFieldsViewTab>('settings');
  const [typePickerOpen, setTypePickerOpen] = useState(false);
  const [overviewOpen, setOverviewOpen] = useState(false);
  const [draftLabels, setDraftLabels] = useState<Record<string, string>>({});
  const persistLabelRef = useRef(
    debounce((fieldId: string, label: string) => {
      void onUpdateField?.(fieldId, { labelCandidate: label.trim() });
    }, 450)
  );

  useEffect(() => () => persistLabelRef.current.flush(), [onUpdateField]);

  const targets = useMemo(() => listFieldSettingsTargets(setupModel), [setupModel]);
  const mappingFields = useMemo(() => sortMappingFields(detectedFields), [detectedFields]);
  const wizard = useMemo(() => getWizardState(setupModel), [setupModel]);
  const currentIndex = resolveCurrentFieldSettingsIndex(targets, wizard);
  const currentTarget = targets[currentIndex] || null;
  const currentField = currentTarget ? resolveFieldFromTarget(setupModel, currentTarget) : null;
  const walkthroughDone = isFieldSettingsWalkthroughComplete(targets, wizard);
  const fieldDisplayOrder = resolveFieldDisplayOrder(mappingFields, currentField?.fieldId);
  const fieldNumber = fieldDisplayOrder ?? (currentTarget ? currentIndex + 1 : 0);
  const activeFieldType = currentField
    ? normalizeSetupFieldType(currentField, detectedFields)
    : 'text';
  const activeFieldPage = resolveFieldPreviewPage(currentField, mappingFields);

  useEffect(() => {
    if (targets.length === 0) return;
    const stored = Math.max(0, Number(wizard.currentFieldSettingsIndex || 0));
    const clamped = Math.min(stored, targets.length - 1);
    if (clamped !== stored) {
      onChange(withWizardState(setupModel, { currentFieldSettingsIndex: clamped }));
    }
  }, [targets.length, wizard.currentFieldSettingsIndex, setupModel, onChange]);

  const switchTab = (tab: SetupFieldsViewTab) => {
    LayoutAnimation.configureNext(LayoutAnimation.Presets.easeInEaseOut);
    setActiveTab(tab);
  };

  const handleTypeChange = (nextType: SetupFieldType) => {
    if (!currentField || !currentTarget || currentTarget.kind === 'table-meta') return;
    const patch = applyFieldTypeChange(currentField, nextType, detectedFields);
    onChange(updateFieldSettingsTarget(setupModel, currentTarget, patch));
    void (async () => {
      await onUpdateField?.(currentField.fieldId, { type: nextType });
    })();
  };

  const handleFieldLabelChange = (fieldId: string, label: string) => {
    if (!currentTarget || currentTarget.kind === 'table-meta') return;
    setDraftLabels((current) => ({ ...current, [fieldId]: label }));
    onChange(updateFieldSettingsTarget(setupModel, currentTarget, { label: label.trim() }));
    persistLabelRef.current(fieldId, label);
  };

  const handleSaveAndNext = () => {
    if (!currentTarget) return;
    void hapticSuccess();
    const next = advanceFieldSettingsWalkthrough(setupModel, targets, currentTarget);
    onChange(next);
  };

  const selectTargetAtIndex = (index: number) => {
    if (index < 0 || index >= targets.length) return;
    onChange(withWizardState(setupModel, { currentFieldSettingsIndex: index }));
  };

  const selectFieldById = (fieldId: string) => {
    const index = targets.findIndex(
      (target) => target.kind !== 'table-meta' && target.fieldId === fieldId
    );
    if (index >= 0) {
      selectTargetAtIndex(index);
      switchTab('settings');
    }
  };

  return (
    <View style={styles.root}>
      <SetupFieldsHeader
        activeTab={activeTab}
        onTabChange={switchTab}
        onBack={onBack}
        applyTopInset={!showWizardNav}
        onOpenOverview={() => setOverviewOpen(true)}
      />

      <View style={styles.body}>
        {activeTab === 'pdf' ? (
          <View style={styles.tabPane}>
            <SetupPdfFieldPreview
              pdfPath={pdfPath}
              detectedFields={detectedFields}
              mappingFields={mappingFields}
              overlayFramesOnly
              activeFieldId={currentField?.fieldId || null}
              activeFieldLabel={
                currentTarget
                  ? resolveTargetDisplayName(
                      setupModel,
                      currentTarget,
                      currentField,
                      mappingFields,
                      draftLabels
                    )
                  : null
              }
              activeFieldPage={activeFieldPage}
              variant="assign"
              onFieldSelect={selectFieldById}
            />
          </View>
        ) : null}
        <View style={[styles.tabPane, activeTab !== 'settings' ? styles.tabPaneHidden : null]}>
          <SetupFieldsSettingsPanel
            setupModel={setupModel}
            target={currentTarget}
            currentTargetIndex={currentIndex}
            totalTargets={targets.length}
            fieldNumber={fieldNumber}
            detectedFields={detectedFields}
            mappingFields={mappingFields}
            draftLabels={draftLabels}
            readOnly={readOnly}
            onChange={onChange}
            onFieldLabelChange={handleFieldLabelChange}
            onOpenTypePicker={() => setTypePickerOpen(true)}
            onSaveAndNext={handleSaveAndNext}
            onSelectTarget={selectTargetAtIndex}
            showSaveButton={Boolean(currentTarget) && !walkthroughDone}
            bottomInset={spacing.sm}
          />
        </View>
      </View>

      {walkthroughDone && !readOnly ? (
        <View style={[styles.footer, { paddingBottom: spacing.xs }]}>
          <PrimaryButton compact label="Vorlage abschließen" onPress={onFinish} />
        </View>
      ) : null}

      <SetupFieldTypePickerSheet
        visible={typePickerOpen}
        value={activeFieldType}
        onClose={() => setTypePickerOpen(false)}
        onSelect={handleTypeChange}
      />

      <SetupFieldsOverview
        visible={overviewOpen}
        setupModel={setupModel}
        detectedFields={detectedFields}
        mappingFields={mappingFields}
        targets={targets}
        currentTargetKey={currentTarget?.key || null}
        onClose={() => setOverviewOpen(false)}
        onSelectTarget={selectTargetAtIndex}
      />
    </View>
  );
}

const styles = StyleSheet.create({
  root: {
    flex: 1,
    backgroundColor: colors.bg
  },
  body: {
    flex: 1,
    minHeight: 0
  },
  tabPane: {
    ...StyleSheet.absoluteFillObject
  },
  tabPaneHidden: {
    opacity: 0,
    pointerEvents: 'none'
  },
  footer: {
    borderTopWidth: 1,
    borderTopColor: colors.border,
    backgroundColor: colors.panel,
    paddingHorizontal: spacing.pageX,
    paddingTop: spacing.xs
  }
});

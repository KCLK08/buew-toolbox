import { useEffect, useMemo, useRef, useState } from 'react';
import { LayoutAnimation, Platform, StyleSheet, UIManager, View } from 'react-native';
import { useSafeAreaInsets } from 'react-native-safe-area-context';

import { PrimaryButton } from '../../../../components/mobile';
import { colors, spacing } from '../../../../constants/theme';
import { hapticSuccess } from '../../../../lib/haptics';
import { systemBottomInset } from '../../../../navigation/systemInsets';
import {
  advanceFieldSettingsWalkthrough,
  applyFieldTypeChange,
  getFieldSettingsProgress,
  isFieldSettingsWalkthroughComplete,
  listFieldSettingsTargets,
  normalizeSetupFieldType,
  resolveCurrentFieldSettingsIndex,
  resolveDetectedFieldTypeLabel,
  resolveFieldFromTarget,
  resolveTargetDisplayName,
  updateFieldSettingsTarget
} from '../../lib/setup-field-settings';
import { getWizardState, withWizardState } from '../../lib/setup-mapping';
import type { DetectedField, SetupFieldType } from '../../types';
import { SetupPdfFieldPreview } from '../SetupPdfFieldPreview';
import { SetupFieldTypePickerSheet } from './SetupFieldTypePickerSheet';
import { SetupFieldsHeader, type SetupFieldsViewTab } from './SetupFieldsHeader';
import { SetupFieldsOverview } from './SetupFieldsOverview';
import { SetupFieldsProgressBar } from './SetupFieldsProgressBar';
import { SetupFieldsSettingsPanel } from './SetupFieldsSettingsPanel';

if (Platform.OS === 'android' && UIManager.setLayoutAnimationEnabledExperimental) {
  UIManager.setLayoutAnimationEnabledExperimental(true);
}

type Props = {
  pdfPath: string | null;
  detectedFields: DetectedField[];
  setupModel: Record<string, unknown>;
  readOnly?: boolean;
  onChange: (next: Record<string, unknown>) => void;
  onFinish: () => void;
  onBack: () => void;
};

export function SetupFieldsStep({
  pdfPath,
  detectedFields,
  setupModel,
  readOnly = false,
  onChange,
  onFinish,
  onBack
}: Props) {
  const insets = useSafeAreaInsets();
  const indexSyncedRef = useRef(false);
  const [activeTab, setActiveTab] = useState<SetupFieldsViewTab>('settings');
  const [typePickerOpen, setTypePickerOpen] = useState(false);
  const [overviewOpen, setOverviewOpen] = useState(false);

  const targets = useMemo(() => listFieldSettingsTargets(setupModel), [setupModel]);
  const wizard = useMemo(() => getWizardState(setupModel), [setupModel]);
  const currentIndex = resolveCurrentFieldSettingsIndex(targets, wizard);
  const currentTarget = targets[currentIndex] || null;
  const currentField = currentTarget ? resolveFieldFromTarget(setupModel, currentTarget) : null;
  const progress = useMemo(() => getFieldSettingsProgress(targets, wizard), [targets, wizard]);
  const walkthroughDone = isFieldSettingsWalkthroughComplete(targets, wizard);
  const fieldNumber = currentTarget ? currentIndex + 1 : 0;
  const displayName = currentTarget
    ? resolveTargetDisplayName(setupModel, currentTarget, currentField)
    : null;
  const detectedTypeLabel =
    currentTarget && currentTarget.kind !== 'table-meta'
      ? resolveDetectedFieldTypeLabel(currentField, detectedFields)
      : null;
  const activeFieldType = currentField
    ? normalizeSetupFieldType(currentField, detectedFields)
    : 'text';

  useEffect(() => {
    if (indexSyncedRef.current || targets.length === 0) return;
    indexSyncedRef.current = true;
    const resolved = resolveCurrentFieldSettingsIndex(targets, wizard);
    if (resolved !== wizard.currentFieldSettingsIndex) {
      onChange(
        withWizardState(setupModel, {
          currentFieldSettingsIndex: resolved
        })
      );
    }
  }, [targets, wizard, setupModel, onChange]);

  const switchTab = (tab: SetupFieldsViewTab) => {
    LayoutAnimation.configureNext(LayoutAnimation.Presets.easeInEaseOut);
    setActiveTab(tab);
  };

  const handleTypeChange = (nextType: SetupFieldType) => {
    if (!currentField || !currentTarget || currentTarget.kind === 'table-meta') return;
    const patch = applyFieldTypeChange(currentField, nextType, detectedFields);
    onChange(updateFieldSettingsTarget(setupModel, currentTarget, patch));
  };

  const handleSaveAndNext = () => {
    if (!currentTarget) return;
    void hapticSuccess();
    const next = advanceFieldSettingsWalkthrough(setupModel, targets, currentTarget);
    onChange(next);
    switchTab('pdf');
  };

  const selectTargetAtIndex = (index: number) => {
    onChange(withWizardState(setupModel, { currentFieldSettingsIndex: index }));
    switchTab('pdf');
  };

  return (
    <View style={styles.root}>
      <SetupFieldsHeader
        activeTab={activeTab}
        onTabChange={switchTab}
        onBack={onBack}
        onOpenOverview={() => setOverviewOpen(true)}
      />

      <SetupFieldsProgressBar
        progress={progress}
        fieldNumber={fieldNumber}
        fieldLabel={displayName}
        detectedTypeLabel={detectedTypeLabel}
      />

      <View style={styles.body}>
        <View style={[styles.tabPane, activeTab !== 'pdf' ? styles.tabPaneHidden : null]}>
          <SetupPdfFieldPreview
            pdfPath={pdfPath}
            detectedFields={detectedFields}
            activeFieldId={currentField?.fieldId || null}
            activeFieldLabel={displayName}
            activeFieldPage={currentField?.page || 1}
            variant="assign"
          />
        </View>
        <View style={[styles.tabPane, activeTab !== 'settings' ? styles.tabPaneHidden : null]}>
          <SetupFieldsSettingsPanel
            setupModel={setupModel}
            target={currentTarget}
            detectedFields={detectedFields}
            readOnly={readOnly}
            onChange={onChange}
            onOpenTypePicker={() => setTypePickerOpen(true)}
            onSaveAndNext={handleSaveAndNext}
            showSaveButton={Boolean(currentTarget) && !walkthroughDone}
          />
        </View>
      </View>

      {walkthroughDone && !readOnly ? (
        <View style={[styles.footer, { paddingBottom: systemBottomInset(insets) + spacing.xs }]}>
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

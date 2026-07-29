import { useEffect, useMemo, useRef, useState } from 'react';
import { LayoutAnimation, Platform, StyleSheet, UIManager, View } from 'react-native';
import { useSafeAreaInsets } from 'react-native-safe-area-context';

import { PrimaryButton } from '../../../../components/mobile';
import { colors, spacing } from '../../../../constants/theme';
import { hapticSuccess } from '../../../../lib/haptics';
import { systemBottomInset } from '../../../../navigation/systemInsets';
import {
  assignFieldToGroup,
  assignFieldToTableColumn,
  getAssignedFieldIds,
  getMappingProgress,
  getNextUnassignedIndex,
  getWizardState,
  isMappingComplete,
  resolveCurrentMappingIndex,
  resolveFieldDisplayLabel,
  type MappingField
} from '../../lib/setup-mapping';
import type { DetectedField, SetupStructureItem } from '../../types';
import { SetupPdfFieldPreview } from '../SetupPdfFieldPreview';
import { SetupAssignFieldOverview } from './SetupAssignFieldOverview';
import { SetupAssignGroupFieldModal } from './SetupAssignGroupFieldModal';
import { SetupAssignHeader, type SetupAssignViewTab } from './SetupAssignHeader';
import { SetupAssignProgressBar } from './SetupAssignProgressBar';
import { SetupAssignStructurePanel } from './SetupAssignStructurePanel';
import { SetupAssignTableColumnModal } from './SetupAssignTableColumnModal';

if (Platform.OS === 'android' && UIManager.setLayoutAnimationEnabledExperimental) {
  UIManager.setLayoutAnimationEnabledExperimental(true);
}

type PendingTarget =
  | { kind: 'group'; item: SetupStructureItem }
  | { kind: 'table'; item: SetupStructureItem }
  | null;

type Props = {
  pdfPath: string | null;
  detectedFields: DetectedField[];
  mappingFields: MappingField[];
  setupModel: Record<string, unknown>;
  readOnly?: boolean;
  onChange: (next: Record<string, unknown>) => void;
  onComplete: (next: Record<string, unknown>) => void;
  onBack: () => void;
};

export function SetupAssignStep({
  pdfPath,
  detectedFields,
  mappingFields,
  setupModel,
  readOnly = false,
  onChange,
  onComplete,
  onBack
}: Props) {
  const insets = useSafeAreaInsets();
  const indexSyncedRef = useRef(false);
  const [activeTab, setActiveTab] = useState<SetupAssignViewTab>('pdf');
  const [pendingTarget, setPendingTarget] = useState<PendingTarget>(null);
  const [showFieldOverview, setShowFieldOverview] = useState(false);

  const wizard = useMemo(() => getWizardState(setupModel), [setupModel]);
  const currentIndex = resolveCurrentMappingIndex(mappingFields, wizard);
  const currentField = mappingFields[currentIndex] || null;
  const progress = useMemo(() => getMappingProgress(mappingFields, wizard), [mappingFields, wizard]);
  const assignedFieldIds = useMemo(() => getAssignedFieldIds(wizard), [wizard]);
  const mappingDone = isMappingComplete(mappingFields, wizard);
  const fieldNumber = currentField ? currentIndex + 1 : 0;
  const currentLabel = currentField ? resolveFieldDisplayLabel(currentField, wizard) : null;

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

  const switchTab = (tab: SetupAssignViewTab) => {
    LayoutAnimation.configureNext(LayoutAnimation.Presets.easeInEaseOut);
    setActiveTab(tab);
  };

  const advanceAfterAssign = (next: Record<string, unknown>) => {
    const nextWizard = getWizardState(next);
    if (isMappingComplete(mappingFields, nextWizard)) {
      onChange(next);
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

  const confirmGroupField = (fieldName: string) => {
    if (!currentField || pendingTarget?.kind !== 'group') return;
    const next = assignFieldToGroup(
      setupModel,
      currentField.fieldId,
      pendingTarget.item.id,
      fieldName
    );
    setPendingTarget(null);
    void hapticSuccess();
    advanceAfterAssign(next);
  };

  const confirmTableColumn = (input: {
    columnId?: string;
    newColumnName?: string;
    fieldLabel: string;
  }) => {
    if (!currentField || pendingTarget?.kind !== 'table') return;
    const next = assignFieldToTableColumn(
      setupModel,
      currentField.fieldId,
      pendingTarget.item.id,
      input
    );
    setPendingTarget(null);
    void hapticSuccess();
    advanceAfterAssign(next);
  };

  const selectFieldAtIndex = (index: number) => {
    onChange({
      ...setupModel,
      wizard: {
        ...wizard,
        currentFieldIndex: index
      }
    });
    switchTab('pdf');
  };

  const pendingGroup =
    pendingTarget?.kind === 'group' && pendingTarget.item.type === 'group'
      ? pendingTarget.item
      : null;
  const pendingTable =
    pendingTarget?.kind === 'table' && pendingTarget.item.type === 'table'
      ? pendingTarget.item
      : null;

  return (
    <View style={styles.root}>
      <SetupAssignHeader
        activeTab={activeTab}
        onTabChange={switchTab}
        onBack={onBack}
        onOpenFields={() => setShowFieldOverview(true)}
      />

      <SetupAssignProgressBar
        progress={progress}
        fieldNumber={fieldNumber}
        fieldLabel={currentLabel}
      />

      <View style={styles.body}>
        <View style={[styles.tabPane, activeTab !== 'pdf' ? styles.tabPaneHidden : null]}>
          <SetupPdfFieldPreview
            pdfPath={pdfPath}
            detectedFields={detectedFields}
            activeFieldId={currentField?.fieldId || null}
            activeFieldLabel={currentLabel}
            activeFieldPage={currentField?.page || 1}
            assignedFieldIds={assignedFieldIds}
            variant="assign"
          />
        </View>
        <View style={[styles.tabPane, activeTab !== 'assign' ? styles.tabPaneHidden : null]}>
          <SetupAssignStructurePanel
            setupModel={setupModel}
            mappingFields={mappingFields}
            currentField={currentField}
            readOnly={readOnly}
            onSelectGroup={(item) => setPendingTarget({ kind: 'group', item })}
            onSelectTable={(item) => setPendingTarget({ kind: 'table', item })}
          />
        </View>
      </View>

      {mappingDone ? (
        <View style={[styles.footer, { paddingBottom: systemBottomInset(insets) + spacing.xs }]}>
          <PrimaryButton compact label="Weiter zu Schritt 3" onPress={() => onComplete(setupModel)} />
        </View>
      ) : null}

      <SetupAssignGroupFieldModal
        visible={Boolean(pendingGroup)}
        groupName={pendingGroup?.name || ''}
        initialFieldName={currentLabel || ''}
        readOnly={readOnly}
        onClose={() => setPendingTarget(null)}
        onConfirm={confirmGroupField}
      />

      <SetupAssignTableColumnModal
        visible={Boolean(pendingTable)}
        tableName={pendingTable?.name || ''}
        columns={pendingTable?.columns || []}
        suggestedFieldName={currentLabel || 'Feld'}
        readOnly={readOnly}
        onClose={() => setPendingTarget(null)}
        onConfirm={confirmTableColumn}
      />

      <SetupAssignFieldOverview
        visible={showFieldOverview}
        mappingFields={mappingFields}
        setupModel={setupModel}
        currentFieldId={currentField?.fieldId || null}
        onClose={() => setShowFieldOverview(false)}
        onSelectField={selectFieldAtIndex}
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

import { useEffect, useMemo, useRef, useState } from 'react';
import { LayoutAnimation, Platform, Pressable, StyleSheet, Text, UIManager, View } from 'react-native';
import { MaterialCommunityIcons } from '@expo/vector-icons';
import { useSafeAreaInsets } from 'react-native-safe-area-context';

import { PrimaryButton } from '../../../../components/mobile';
import { colors, spacing, typography } from '../../../../constants/theme';
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
import { createManualFieldInput } from '../../lib/template-field';
import { getStructureItems } from '../../lib/setup-structure';
import type { DetectedField, FieldRect, SetupStructureItem } from '../../types';
import { SetupPdfFieldPreview } from '../SetupPdfFieldPreview';
import { SetupAssignFieldOverview } from './SetupAssignFieldOverview';
import { SetupAssignGroupFieldModal } from './SetupAssignGroupFieldModal';
import { SetupAssignHeader, type SetupAssignViewTab } from './SetupAssignHeader';
import { SetupAssignProgressBar } from './SetupAssignProgressBar';
import { SetupAssignSourceBanner } from './SetupAssignSourceBanner';
import { SetupAssignStructurePanel } from './SetupAssignStructurePanel';
import { SetupAssignTableColumnModal } from './SetupAssignTableColumnModal';
import { SetupManualFieldModal } from './SetupManualFieldModal';

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
  onFieldsChanged?: () => void;
  onCreateManualField?: (
    field: ReturnType<typeof createManualFieldInput>,
    target: { kind: 'group'; id: string } | { kind: 'table'; id: string } | null
  ) => Promise<DetectedField | null>;
};

export function SetupAssignStep({
  pdfPath,
  detectedFields,
  mappingFields,
  setupModel,
  readOnly = false,
  onChange,
  onComplete,
  onBack,
  onFieldsChanged,
  onCreateManualField
}: Props) {
  const insets = useSafeAreaInsets();
  const indexSyncedRef = useRef(false);
  const [activeTab, setActiveTab] = useState<SetupAssignViewTab>('pdf');
  const [pendingTarget, setPendingTarget] = useState<PendingTarget>(null);
  const [showFieldOverview, setShowFieldOverview] = useState(false);
  const [drawMode, setDrawMode] = useState(false);
  const [pendingDraw, setPendingDraw] = useState<{ page: number; rect: FieldRect } | null>(null);

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

      <SetupAssignSourceBanner fields={detectedFields} />

      <SetupAssignProgressBar
        progress={progress}
        fieldNumber={fieldNumber}
        fieldLabel={currentLabel}
      />

      <View style={styles.body}>
        <View style={[styles.tabPane, activeTab !== 'pdf' ? styles.tabPaneHidden : null]}>
          {!readOnly ? (
            <View style={styles.pdfToolbar}>
              <Pressable
                accessibilityRole="button"
                style={[styles.addFieldBtn, drawMode ? styles.addFieldBtnActive : null]}
                onPress={() => setDrawMode((value) => !value)}
              >
                <MaterialCommunityIcons
                  name={drawMode ? 'gesture-tap' : 'plus'}
                  size={18}
                  color={drawMode ? colors.white : colors.accent}
                />
                <Text style={drawMode ? styles.addFieldLabelActive : styles.addFieldLabel}>
                  {drawMode ? 'Bereich markieren…' : '+ Feld hinzufügen'}
                </Text>
              </Pressable>
            </View>
          ) : null}
          <SetupPdfFieldPreview
            pdfPath={pdfPath}
            detectedFields={detectedFields}
            activeFieldId={currentField?.fieldId || null}
            activeFieldLabel={currentLabel}
            activeFieldPage={currentField?.page || 1}
            assignedFieldIds={assignedFieldIds}
            variant="assign"
            drawMode={drawMode}
            onFieldDrawn={(payload) => {
              setDrawMode(false);
              setPendingDraw(payload);
            }}
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

      <SetupManualFieldModal
        visible={Boolean(pendingDraw)}
        page={pendingDraw?.page || 1}
        rect={pendingDraw?.rect || { x: 0, y: 0, width: 0, height: 0 }}
        setupModel={setupModel}
        readOnly={readOnly}
        onClose={() => setPendingDraw(null)}
        onConfirm={(input) => {
          if (!pendingDraw || !onCreateManualField) {
            setPendingDraw(null);
            return;
          }
          void (async () => {
            const draft = createManualFieldInput({
              name: input.name,
              type: input.type,
              page: pendingDraw.page,
              rect: pendingDraw.rect
            });
            const saved = await onCreateManualField(draft, input.target);
            setPendingDraw(null);
            if (!saved) return;
            void hapticSuccess();
            onFieldsChanged?.();
            if (input.target?.kind === 'group') {
              const next = assignFieldToGroup(setupModel, saved.fieldId, input.target.id, input.name);
              onChange(next);
            } else if (input.target?.kind === 'table') {
              const table = getStructureItems(setupModel).find(
                (item) => item.id === input.target?.id && item.type === 'table'
              );
              const firstColumn =
                table && table.type === 'table' ? table.columns[0]?.id : undefined;
              if (firstColumn) {
                const next = assignFieldToTableColumn(setupModel, saved.fieldId, input.target.id, {
                  columnId: firstColumn,
                  fieldLabel: input.name
                });
                onChange(next);
              }
            }
          })();
        }}
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
  },
  pdfToolbar: {
    paddingHorizontal: spacing.pageX,
    paddingVertical: spacing.xs,
    backgroundColor: colors.panel,
    borderBottomWidth: 1,
    borderBottomColor: colors.border
  },
  addFieldBtn: {
    flexDirection: 'row',
    alignItems: 'center',
    justifyContent: 'center',
    gap: spacing.xs,
    minHeight: spacing.touchMin,
    borderRadius: 999,
    borderWidth: 1,
    borderColor: colors.accent,
    backgroundColor: colors.panelElevated,
    paddingHorizontal: spacing.md
  },
  addFieldBtnActive: {
    backgroundColor: colors.accent,
    borderColor: colors.accent
  },
  addFieldLabel: {
    ...typography.caption,
    color: colors.accent,
    fontFamily: 'SpaceGrotesk_600SemiBold'
  },
  addFieldLabelActive: {
    ...typography.caption,
    color: colors.white,
    fontFamily: 'SpaceGrotesk_600SemiBold'
  }
});

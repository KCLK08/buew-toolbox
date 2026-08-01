import { useEffect, useMemo, useRef, useState } from 'react';
import { LayoutAnimation, Platform, Pressable, StyleSheet, Text, UIManager, View, Alert } from 'react-native';
import { MaterialCommunityIcons } from '@expo/vector-icons';

import { PrimaryButton } from '../../../../components/mobile';
import { colors, spacing, typography } from '../../../../constants/theme';
import { debounce } from '../../../../lib/debounce';
import { hapticSuccess } from '../../../../lib/haptics';
import {
  assignFieldToGroup,
  assignFieldToTableColumn,
  getAssignedFieldIds,
  getMappingProgress,
  getNextUnassignedIndex,
  getWizardState,
  isMappingComplete,
  removeFieldFromWizard,
  resolveCurrentMappingIndex,
  resolveFieldDisplayLabel,
  type MappingField
} from '../../lib/setup-mapping';
import { createManualFieldInput } from '../../lib/template-field';
import { getStructureItems } from '../../lib/setup-structure';
import {
  geometryDraftFromField,
  isManualMappingField
} from '../../lib/setup-manual-field';
import type { DetectedField, FieldRect, SetupFieldType, SetupStructureItem } from '../../types';
import { SetupPdfFieldPreview } from '../SetupPdfFieldPreview';
import { SetupAssignFieldListPanel } from './SetupAssignFieldListPanel';
import { SetupAssignFieldOverview } from './SetupAssignFieldOverview';
import { SetupAssignHeader, type SetupAssignViewTab } from './SetupAssignHeader';
import { SetupAssignSourceBanner } from './SetupAssignSourceBanner';
import { SetupAssignTableColumnModal } from './SetupAssignTableColumnModal';
import { SetupManualFieldModal } from './SetupManualFieldModal';
import { SetupManualFieldActionMenu } from './SetupManualFieldActionMenu';
import { SetupDrawConfirmPanel, SETUP_DRAFT_CONFIRM_RESERVE_PX } from './SetupDrawConfirmPanel';

if (Platform.OS === 'android' && UIManager.setLayoutAnimationEnabledExperimental) {
  UIManager.setLayoutAnimationEnabledExperimental(true);
}

type Props = {
  pdfPath: string | null;
  detectedFields: DetectedField[];
  mappingFields: MappingField[];
  setupModel: Record<string, unknown>;
  readOnly?: boolean;
  showWizardNav?: boolean;
  onChange: (next: Record<string, unknown>) => void;
  onComplete: (next: Record<string, unknown>) => void;
  onBack: () => void;
  onFieldsChanged?: () => void;
  onUpdateField?: (
    fieldId: string,
    patch: {
      labelCandidate?: string;
      type?: string;
      geometry?: { page: number; rect: FieldRect } | null;
      page?: number;
    }
  ) => Promise<void>;
  onDeleteField?: (fieldId: string) => Promise<void>;
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
  showWizardNav = false,
  onChange,
  onComplete,
  onBack,
  onFieldsChanged,
  onUpdateField,
  onDeleteField,
  onCreateManualField
}: Props) {
  const indexSyncedRef = useRef(false);
  const [activeTab, setActiveTab] = useState<SetupAssignViewTab>('pdf');
  const [showFieldOverview, setShowFieldOverview] = useState(false);
  const [drawMode, setDrawMode] = useState(false);
  const [draftDraw, setDraftDraw] = useState<{ page: number; rect: FieldRect } | null>(null);
  const [draftRectEditable, setDraftRectEditable] = useState(false);
  const [pendingDraw, setPendingDraw] = useState<{ page: number; rect: FieldRect } | null>(null);
  const [tableAssignTarget, setTableAssignTarget] = useState<SetupStructureItem | null>(null);
  const [positionEditFieldId, setPositionEditFieldId] = useState<string | null>(null);
  const [draftLabels, setDraftLabels] = useState<Record<string, string>>({});
  const persistFieldNameRef = useRef(
    debounce((fieldId: string, name: string) => {
      void onUpdateField?.(fieldId, { labelCandidate: name.trim() });
    }, 450)
  );

  useEffect(() => () => persistFieldNameRef.current.flush(), [onUpdateField]);

  const wizard = useMemo(() => getWizardState(setupModel), [setupModel]);
  const currentIndex = resolveCurrentMappingIndex(mappingFields, wizard);
  const currentField = mappingFields[currentIndex] || null;
  const progress = useMemo(() => getMappingProgress(mappingFields, wizard), [mappingFields, wizard]);
  const assignedFieldIds = useMemo(() => getAssignedFieldIds(wizard), [wizard]);
  const mappingDone = isMappingComplete(mappingFields, wizard);
  const fieldNumber = currentField?.displayOrder ?? 0;
  const resolveLabel = (field: MappingField) =>
    resolveFieldDisplayLabel(field, wizard, draftLabels);
  const currentLabel = currentField ? resolveLabel(currentField) : null;

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
    if (tab !== 'pdf') {
      cancelPositionEdit();
    }
    setDrawMode(false);
    if (!positionEditFieldId) {
      setDraftDraw(null);
      setDraftRectEditable(false);
    }
  };

  const cancelPositionEdit = () => {
    setPositionEditFieldId(null);
    setDraftDraw(null);
    setDraftRectEditable(false);
  };

  const cancelDraftDraw = () => {
    if (positionEditFieldId) {
      cancelPositionEdit();
      return;
    }
    setDraftDraw(null);
    setDraftRectEditable(false);
    setDrawMode(false);
  };

  const confirmDraftDraw = () => {
    if (!draftDraw) return;
    if (positionEditFieldId) {
      void (async () => {
        await onUpdateField?.(positionEditFieldId, {
          geometry: { page: draftDraw.page, rect: draftDraw.rect },
          page: draftDraw.page
        });
        cancelPositionEdit();
        onFieldsChanged?.();
        void hapticSuccess();
      })();
      return;
    }
    setPendingDraw(draftDraw);
    setDraftDraw(null);
    setDraftRectEditable(false);
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

  const assignToGroup = (item: SetupStructureItem) => {
    if (!currentField || readOnly || item.type !== 'group') return;
    const trimmed = resolveLabel(currentField);
    const next = assignFieldToGroup(setupModel, currentField.fieldId, item.id, trimmed);
    void hapticSuccess();
    void (async () => {
      if (trimmed) {
        await onUpdateField?.(currentField.fieldId, { labelCandidate: trimmed });
      }
      advanceAfterAssign(next);
    })();
  };

  const assignToTable = (item: SetupStructureItem) => {
    if (!currentField || readOnly || item.type !== 'table') return;
    if (item.columns.length === 0) {
      Alert.alert(
        'Keine Spalten',
        'Lege zuerst in Schritt 1 Spalten für diese Tabelle an.'
      );
      return;
    }
    setTableAssignTarget(item);
  };

  const confirmTableAssignment = (input: {
    columnId?: string;
    newColumnName?: string;
    fieldLabel: string;
  }) => {
    if (!currentField || !tableAssignTarget || tableAssignTarget.type !== 'table') return;
    const trimmed = resolveLabel(currentField);
    const next = assignFieldToTableColumn(setupModel, currentField.fieldId, tableAssignTarget.id, {
      columnId: input.columnId,
      newColumnName: input.newColumnName,
      fieldLabel: trimmed || input.fieldLabel
    });
    setTableAssignTarget(null);
    void hapticSuccess();
    void (async () => {
      if (trimmed) {
        await onUpdateField?.(currentField.fieldId, { labelCandidate: trimmed });
      }
      advanceAfterAssign(next);
    })();
  };

  const selectFieldAtIndex = (index: number) => {
    onChange({
      ...setupModel,
      wizard: {
        ...wizard,
        currentFieldIndex: index
      }
    });
  };

  const handleFieldNameChange = (fieldId: string, name: string) => {
    setDraftLabels((current) => ({ ...current, [fieldId]: name }));
    persistFieldNameRef.current(fieldId, name);
  };

  const handleFieldTypeChange = (fieldId: string, type: SetupFieldType) => {
    void (async () => {
      await onUpdateField?.(fieldId, { type });
      onFieldsChanged?.();
    })();
  };

  const handleDeleteField = (fieldId: string) => {
    void (async () => {
      await onDeleteField?.(fieldId);
      const next = removeFieldFromWizard(setupModel, fieldId, mappingFields.length - 1);
      onChange(next);
      onFieldsChanged?.();
      void hapticSuccess();
    })();
  };

  const handleConfirmDeleteField = (fieldId: string) => {
    if (readOnly) return;
    Alert.alert('Feld löschen', 'Möchten Sie dieses Feld wirklich löschen?', [
      { text: 'Abbrechen', style: 'cancel' },
      { text: 'Löschen', style: 'destructive', onPress: () => handleDeleteField(fieldId) }
    ]);
  };

  const startPositionEdit = (fieldId: string) => {
    if (readOnly) return;
    const field = mappingFields.find((entry) => entry.fieldId === fieldId);
    const draft = field ? geometryDraftFromField(field) : null;
    if (!draft) {
      Alert.alert('Position', 'Für dieses Feld ist keine Position gespeichert.');
      return;
    }
    const index = mappingFields.findIndex((entry) => entry.fieldId === fieldId);
    if (index >= 0) {
      selectFieldAtIndex(index);
    }
    setPositionEditFieldId(fieldId);
    setDraftDraw(draft);
    setDraftRectEditable(false);
    setDrawMode(false);
    setPendingDraw(null);
    LayoutAnimation.configureNext(LayoutAnimation.Presets.easeInEaseOut);
    setActiveTab('pdf');
  };

  const selectFieldById = (fieldId: string) => {
    const index = mappingFields.findIndex((field) => field.fieldId === fieldId);
    if (index >= 0) {
      cancelPositionEdit();
      selectFieldAtIndex(index);
      switchTab('fields');
    }
  };

  const isManualCurrent = isManualMappingField(currentField);
  const showPdfManualActions =
    activeTab === 'pdf' &&
    isManualCurrent &&
    !readOnly &&
    !drawMode &&
    !draftDraw &&
    !positionEditFieldId;

  return (
    <View style={styles.root}>
      <SetupAssignHeader
        activeTab={activeTab}
        onTabChange={switchTab}
        onBack={onBack}
        applyTopInset={!showWizardNav}
        onOpenFields={() => {
          switchTab('fields');
        }}
      />

      <SetupAssignSourceBanner fields={detectedFields} />

      <View style={styles.body}>
        <View style={[styles.tabPane, activeTab !== 'pdf' ? styles.tabPaneHidden : null]}>
          {!readOnly && !draftDraw && !positionEditFieldId ? (
            <View style={styles.pdfToolbar}>
              <Pressable
                accessibilityRole="button"
                style={[styles.addFieldBtn, drawMode ? styles.addFieldBtnActive : null]}
                onPress={() => {
                  if (drawMode) {
                    setDrawMode(false);
                    return;
                  }
                  cancelDraftDraw();
                  setDrawMode(true);
                }}
              >
                <MaterialCommunityIcons
                  name={drawMode ? 'gesture-tap' : 'plus'}
                  size={18}
                  color={drawMode ? colors.white : colors.accent}
                />
                <Text style={drawMode ? styles.addFieldLabelActive : styles.addFieldLabel}>
                  {drawMode ? 'Bereich markieren…' : '+ Feld markieren'}
                </Text>
              </Pressable>
            </View>
          ) : null}
          <View style={styles.pdfStage}>
            <SetupPdfFieldPreview
              pdfPath={pdfPath}
              detectedFields={detectedFields}
              mappingFields={mappingFields}
              resolveFieldLabel={resolveLabel}
              activeFieldId={currentField?.fieldId || null}
              activeFieldLabel={currentLabel}
              activeFieldPage={currentField?.page || 1}
              assignedFieldIds={assignedFieldIds}
              overlayFramesOnly
              variant="assign"
              drawMode={drawMode && !draftDraw}
              draftRect={draftDraw}
              draftRectEditable={draftRectEditable}
              draftConfirmReservePx={draftDraw ? SETUP_DRAFT_CONFIRM_RESERVE_PX : 0}
              onFieldSelect={selectFieldById}
              onFieldDrawDraft={(payload) => {
                setDrawMode(false);
                setDraftDraw({ page: payload.page, rect: payload.rect });
                setDraftRectEditable(false);
              }}
              onFieldDraftUpdated={(payload) => {
                setDraftDraw({ page: payload.page, rect: payload.rect });
              }}
            />
            {draftDraw ? (
              <View style={styles.draftConfirmOverlay} pointerEvents="box-none">
                <SetupDrawConfirmPanel
                  rectEditEnabled={draftRectEditable}
                  bottomInset={spacing.xxs}
                  confirmLabel={positionEditFieldId ? 'Position speichern' : 'OK'}
                  onEnableEdit={() => setDraftRectEditable(true)}
                  onConfirm={confirmDraftDraw}
                  onCancel={cancelDraftDraw}
                />
              </View>
            ) : null}
            {showPdfManualActions && currentField ? (
              <View style={styles.pdfActionOverlay} pointerEvents="box-none">
                <View style={styles.pdfActionPanel}>
                  <SetupManualFieldActionMenu
                    compact
                    onEdit={() => switchTab('fields')}
                    onEditPosition={() => startPositionEdit(currentField.fieldId)}
                    onDelete={() => handleConfirmDeleteField(currentField.fieldId)}
                  />
                </View>
              </View>
            ) : null}
          </View>
        </View>
        <View style={[styles.tabPane, activeTab !== 'fields' ? styles.tabPaneHidden : null]}>
          <SetupAssignFieldListPanel
            mappingFields={mappingFields}
            setupModel={setupModel}
            currentField={currentField}
            currentFieldIndex={currentIndex}
            unassignedCount={progress.open}
            draftLabels={draftLabels}
            readOnly={readOnly}
            onSelectField={selectFieldAtIndex}
            onShowInPdf={() => switchTab('pdf')}
            onAssignGroup={assignToGroup}
            onAssignTable={assignToTable}
            onChangeFieldName={handleFieldNameChange}
            onChangeFieldType={handleFieldTypeChange}
            onEditFieldPosition={startPositionEdit}
            onConfirmDeleteField={handleConfirmDeleteField}
            bottomInset={spacing.sm}
          />
        </View>
      </View>

      {mappingDone ? (
        <View style={[styles.footer, { paddingBottom: spacing.xs }]}>
          <PrimaryButton compact label="Weiter zu Schritt 3" onPress={() => onComplete(setupModel)} />
        </View>
      ) : null}

      <SetupAssignFieldOverview
        visible={showFieldOverview}
        mappingFields={mappingFields}
        setupModel={setupModel}
        currentFieldId={currentField?.fieldId || null}
        onClose={() => setShowFieldOverview(false)}
        onSelectField={(index) => {
          selectFieldAtIndex(index);
          switchTab('pdf');
        }}
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
              } else {
                onChange(setupModel);
              }
            }
            switchTab('fields');
          })();
        }}
      />

      <SetupAssignTableColumnModal
        visible={Boolean(tableAssignTarget)}
        tableName={tableAssignTarget?.type === 'table' ? tableAssignTarget.name : ''}
        columns={tableAssignTarget?.type === 'table' ? tableAssignTarget.columns : []}
        suggestedFieldName={currentLabel || ''}
        readOnly={readOnly}
        allowCreateColumn={false}
        onClose={() => setTableAssignTarget(null)}
        onConfirm={confirmTableAssignment}
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
  pdfStage: {
    flex: 1,
    minHeight: 0,
    position: 'relative'
  },
  draftConfirmOverlay: {
    position: 'absolute',
    left: 0,
    right: 0,
    bottom: 0
  },
  pdfActionOverlay: {
    position: 'absolute',
    left: 0,
    right: 0,
    top: spacing.xs,
    alignItems: 'center'
  },
  pdfActionPanel: {
    maxWidth: '100%',
    paddingHorizontal: spacing.pageX,
    paddingVertical: spacing.xxs,
    borderRadius: 12,
    borderWidth: 1,
    borderColor: colors.border,
    backgroundColor: 'rgba(255, 255, 255, 0.96)'
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
    paddingVertical: spacing.xxs,
    backgroundColor: colors.panel,
    borderBottomWidth: 1,
    borderBottomColor: colors.border
  },
  addFieldBtn: {
    flexDirection: 'row',
    alignItems: 'center',
    justifyContent: 'center',
    gap: spacing.xs,
    minHeight: 40,
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

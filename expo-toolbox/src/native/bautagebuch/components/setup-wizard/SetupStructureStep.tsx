import { useEffect, useMemo, useRef, useState } from 'react';
import { Alert, LayoutAnimation, Platform, Pressable, StyleSheet, Text, UIManager, View } from 'react-native';
import { MaterialCommunityIcons } from '@expo/vector-icons';
import { useSafeAreaInsets } from 'react-native-safe-area-context';

import { PrimaryButton } from '../../../../components/mobile';
import { colors, spacing, typography } from '../../../../constants/theme';
import { hapticLight, hapticSuccess } from '../../../../lib/haptics';
import { systemBottomInset } from '../../../../navigation/systemInsets';
import { countAssignedFieldsForStructureItem } from '../../lib/setup-mapping';
import {
  addStructureGroup,
  addStructureTable,
  completeStructureStep,
  deleteStructureItem,
  deleteStructureItemWithFields,
  getStructureItems,
  moveStructureItem,
  updateStructureGroup,
  updateStructureTable,
  validateStructureStep
} from '../../lib/setup-structure';
import type { SetupStructureItem } from '../../types';
import { PdfPreviewPanel } from '../PdfPreviewPanel';
import { SetupStructureGroupModal } from './SetupStructureGroupModal';
import { SetupStructureHeader, type SetupStructureViewTab } from './SetupStructureHeader';
import { SetupStructureList } from './SetupStructureList';
import { SetupStructureTableModal } from './SetupStructureTableModal';

if (Platform.OS === 'android' && UIManager.setLayoutAnimationEnabledExperimental) {
  UIManager.setLayoutAnimationEnabledExperimental(true);
}

type EditorState =
  | { kind: 'group'; mode: 'create' | 'edit'; item?: SetupStructureItem }
  | { kind: 'table'; mode: 'create' | 'edit'; item?: SetupStructureItem }
  | null;

type Props = {
  pdfPath: string | null;
  setupModel: Record<string, unknown>;
  readOnly?: boolean;
  editMode?: boolean;
  onChange: (next: Record<string, unknown>) => void;
  onComplete: (next: Record<string, unknown>) => void;
  onBack: () => void;
  onNavigateToAssign?: () => void;
};

export function SetupStructureStep({
  pdfPath,
  setupModel,
  readOnly = false,
  editMode = false,
  onChange,
  onComplete,
  onBack,
  onNavigateToAssign
}: Props) {
  const insets = useSafeAreaInsets();
  const hintTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);
  const [activeTab, setActiveTab] = useState<SetupStructureViewTab>('structure');
  const [editor, setEditor] = useState<EditorState>(null);
  const [continueHint, setContinueHint] = useState<string | null>(null);

  const structure = useMemo(() => getStructureItems(setupModel), [setupModel]);

  useEffect(() => {
    return () => {
      if (hintTimerRef.current) clearTimeout(hintTimerRef.current);
    };
  }, []);

  const switchTab = (tab: SetupStructureViewTab) => {
    LayoutAnimation.configureNext(LayoutAnimation.Presets.easeInEaseOut);
    setActiveTab(tab);
  };

  const showContinueHint = (message: string) => {
    void hapticLight();
    setContinueHint(message);
    if (hintTimerRef.current) clearTimeout(hintTimerRef.current);
    hintTimerRef.current = setTimeout(() => setContinueHint(null), 3200);
  };

  const openGroupEditor = (item?: SetupStructureItem) => {
    setEditor({ kind: 'group', mode: item ? 'edit' : 'create', item });
  };

  const openTableEditor = (item?: SetupStructureItem) => {
    setEditor({ kind: 'table', mode: item ? 'edit' : 'create', item });
  };

  const saveGroup = (input: { name: string; description?: string }) => {
    if (editor?.kind !== 'group') return;
    const next =
      editor.mode === 'edit' && editor.item
        ? updateStructureGroup(setupModel, editor.item.id, input)
        : addStructureGroup(setupModel, input);
    onChange(next);
    setEditor(null);
    setContinueHint(null);
  };

  const saveTable = (input: { name: string; columns: Array<{ id?: string; name: string }> }) => {
    if (editor?.kind !== 'table') return;
    const next =
      editor.mode === 'edit' && editor.item
        ? updateStructureTable(setupModel, editor.item.id, input)
        : addStructureTable(setupModel, input);
    onChange(next);
    setEditor(null);
    setContinueHint(null);
  };

  const handleDelete = (item: SetupStructureItem) => {
    const assignedCount = countAssignedFieldsForStructureItem(setupModel, item);
    if (assignedCount === 0) {
      onChange(deleteStructureItem(setupModel, item.id));
      return;
    }

    const kindLabel = item.type === 'group' ? 'Gruppe' : 'Tabelle';
    Alert.alert(
      `${kindLabel} löschen`,
      `Diese ${kindLabel.toLowerCase()} enthält bereits ${assignedCount} Felder.\n\nWas möchtest du tun?`,
      [
        { text: 'Abbrechen', style: 'cancel' },
        {
          text: 'Felder neu zuordnen',
          onPress: () => onNavigateToAssign?.()
        },
        {
          text: 'Inkl. Felder löschen',
          style: 'destructive',
          onPress: () => {
            void hapticLight();
            onChange(deleteStructureItemWithFields(setupModel, item.id));
          }
        }
      ]
    );
  };

  const handleMove = (id: string, direction: -1 | 1) => {
    LayoutAnimation.configureNext(LayoutAnimation.Presets.easeInEaseOut);
    onChange(moveStructureItem(setupModel, id, direction));
  };

  const handleContinue = () => {
    const issue = validateStructureStep(structure);
    if (issue) {
      showContinueHint(issue);
      return;
    }
    try {
      const next = completeStructureStep(setupModel);
      void hapticSuccess();
      onComplete(next);
    } catch (err) {
      showContinueHint(err instanceof Error ? err.message : 'Struktur konnte nicht gespeichert werden.');
    }
  };

  const groupEditorVisible = editor?.kind === 'group';
  const tableEditorVisible = editor?.kind === 'table';
  const editingGroup = editor?.kind === 'group' && editor.item?.type === 'group' ? editor.item : null;
  const editingTable = editor?.kind === 'table' && editor.item?.type === 'table' ? editor.item : null;

  return (
    <View style={styles.root}>
      <SetupStructureHeader
        activeTab={activeTab}
        onTabChange={switchTab}
        onBack={onBack}
        applyTopInset={!editMode}
      />

      <View style={styles.body}>
        <View style={[styles.tabPane, activeTab !== 'pdf' ? styles.tabPaneHidden : null]}>
          <PdfPreviewPanel pdfPath={pdfPath} fullscreen />
        </View>
        <View style={[styles.tabPane, activeTab !== 'structure' ? styles.tabPaneHidden : null]}>
          <SetupStructureList
            items={structure}
            readOnly={readOnly}
            onAddGroup={() => openGroupEditor()}
            onAddTable={() => openTableEditor()}
            onEdit={(item) => {
              if (item.type === 'group') openGroupEditor(item);
              else openTableEditor(item);
            }}
            onDelete={handleDelete}
            onDeletePress={handleDelete}
            onMove={handleMove}
          />
        </View>
      </View>

      {activeTab === 'structure' && !readOnly ? (
        <View style={[styles.footerWrap, { paddingBottom: systemBottomInset(insets) + spacing.xs }]}>
          {continueHint ? (
            <View style={styles.hintBubble}>
              <MaterialCommunityIcons name="information-outline" size={16} color={colors.accent2} />
              <Text style={styles.hintText}>{continueHint}</Text>
            </View>
          ) : null}
          <View style={styles.footer}>
            <PrimaryButton compact label="Später" variant="ghost" onPress={onBack} style={styles.footerBtn} />
            <PrimaryButton compact label="Weiter" onPress={handleContinue} style={styles.footerBtnPrimary} />
          </View>
        </View>
      ) : null}

      <SetupStructureGroupModal
        visible={groupEditorVisible}
        initialName={editingGroup?.name || ''}
        initialDescription={editingGroup?.description || ''}
        readOnly={readOnly}
        onClose={() => setEditor(null)}
        onSave={saveGroup}
      />

      <SetupStructureTableModal
        visible={tableEditorVisible}
        initialName={editingTable?.name || ''}
        initialColumns={editingTable?.columns.map((column) => ({ id: column.id, name: column.name })) || []}
        readOnly={readOnly}
        onClose={() => setEditor(null)}
        onSave={saveTable}
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
  footerWrap: {
    position: 'relative',
    borderTopWidth: 1,
    borderTopColor: colors.border,
    backgroundColor: colors.panel,
    paddingHorizontal: spacing.pageX,
    paddingTop: spacing.xs
  },
  hintBubble: {
    position: 'absolute',
    left: spacing.pageX,
    right: spacing.pageX,
    bottom: '100%',
    marginBottom: spacing.xxs,
    flexDirection: 'row',
    alignItems: 'flex-start',
    gap: spacing.xxs,
    paddingHorizontal: spacing.sm,
    paddingVertical: spacing.xs,
    borderRadius: 10,
    backgroundColor: colors.accent2,
    shadowColor: '#1A1916',
    shadowOpacity: 0.15,
    shadowRadius: 8,
    shadowOffset: { width: 0, height: 2 },
    elevation: 4
  },
  hintText: {
    ...typography.caption,
    color: colors.white,
    flex: 1,
    lineHeight: 18
  },
  footer: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: spacing.xs
  },
  footerBtn: {
    flex: 1
  },
  footerBtnPrimary: {
    flex: 1.4
  }
});

import { useMemo, useState } from 'react';
import { LayoutAnimation, Platform, StyleSheet, Text, UIManager, View } from 'react-native';
import { MaterialCommunityIcons } from '@expo/vector-icons';
import { useSafeAreaInsets } from 'react-native-safe-area-context';

import { PrimaryButton } from '../../../../components/mobile';
import { colors, shadows, spacing, typography } from '../../../../constants/theme';
import { hapticSuccess } from '../../../../lib/haptics';
import { systemBottomInset } from '../../../../navigation/systemInsets';
import {
  addStructureGroup,
  addStructureTable,
  completeStructureStep,
  deleteStructureItem,
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
  templateName?: string;
  pdfPath: string | null;
  setupModel: Record<string, unknown>;
  readOnly?: boolean;
  onChange: (next: Record<string, unknown>) => void;
  onComplete: (next: Record<string, unknown>) => void;
  onBack: () => void;
};

export function SetupStructureStep({
  templateName,
  pdfPath,
  setupModel,
  readOnly = false,
  onChange,
  onComplete,
  onBack
}: Props) {
  const insets = useSafeAreaInsets();
  const [activeTab, setActiveTab] = useState<SetupStructureViewTab>('structure');
  const [editor, setEditor] = useState<EditorState>(null);
  const [error, setError] = useState<string | null>(null);

  const structure = useMemo(() => getStructureItems(setupModel), [setupModel]);

  const switchTab = (tab: SetupStructureViewTab) => {
    LayoutAnimation.configureNext(LayoutAnimation.Presets.easeInEaseOut);
    setActiveTab(tab);
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
    setError(null);
  };

  const saveTable = (input: { name: string; columns: Array<{ id?: string; name: string }> }) => {
    if (editor?.kind !== 'table') return;
    const next =
      editor.mode === 'edit' && editor.item
        ? updateStructureTable(setupModel, editor.item.id, input)
        : addStructureTable(setupModel, input);
    onChange(next);
    setEditor(null);
    setError(null);
  };

  const handleDelete = (item: SetupStructureItem) => {
    onChange(deleteStructureItem(setupModel, item.id));
  };

  const handleMove = (id: string, direction: -1 | 1) => {
    LayoutAnimation.configureNext(LayoutAnimation.Presets.easeInEaseOut);
    onChange(moveStructureItem(setupModel, id, direction));
  };

  const handleContinue = () => {
    const issue = validateStructureStep(structure);
    if (issue) {
      setError(issue);
      return;
    }
    try {
      const next = completeStructureStep(setupModel);
      void hapticSuccess();
      onComplete(next);
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Struktur konnte nicht gespeichert werden.');
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
        templateName={templateName}
        onTabChange={switchTab}
        onBack={onBack}
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
            onMove={handleMove}
          />
        </View>
      </View>

      {activeTab === 'structure' && !readOnly ? (
        <View style={[styles.footer, { paddingBottom: systemBottomInset(insets) + spacing.sm }]}>
          {error ? (
            <View style={styles.errorBanner}>
              <MaterialCommunityIcons name="alert-circle-outline" size={18} color={colors.danger} />
              <Text style={styles.error}>{error}</Text>
            </View>
          ) : null}
          <View style={styles.footerSummary}>
            <Text style={styles.footerSummaryLabel}>
              {structure.length === 0
                ? 'Noch keine Bereiche definiert'
                : `${structure.length} ${structure.length === 1 ? 'Bereich' : 'Bereiche'} bereit`}
            </Text>
          </View>
          <PrimaryButton
            label="Weiter zur Feldzuordnung"
            onPress={handleContinue}
            disabled={structure.length === 0}
          />
          <PrimaryButton label="Später fortsetzen" variant="ghost" onPress={onBack} />
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
  footer: {
    gap: spacing.sm,
    paddingHorizontal: spacing.pageX,
    paddingTop: spacing.md,
    borderTopWidth: 1,
    borderTopColor: colors.border,
    backgroundColor: colors.panel,
    ...shadows.card
  },
  footerSummary: {
    alignItems: 'center'
  },
  footerSummaryLabel: {
    ...typography.caption,
    color: colors.muted
  },
  errorBanner: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: spacing.xs,
    padding: spacing.sm,
    borderRadius: 12,
    backgroundColor: 'rgba(161, 44, 36, 0.08)',
    borderWidth: 1,
    borderColor: 'rgba(161, 44, 36, 0.2)'
  },
  error: {
    ...typography.caption,
    color: colors.danger,
    flex: 1
  }
});

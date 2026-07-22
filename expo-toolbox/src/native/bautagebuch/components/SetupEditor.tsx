// @ts-nocheck
import { useEffect, useMemo, useState } from 'react';
import { Alert, Pressable, ScrollView, StyleSheet, Text, View } from 'react-native';
import { useSafeAreaInsets } from 'react-native-safe-area-context';

import { ListItem, PrimaryButton, TextField } from '../../../components/mobile';
import { colors, spacing, typography } from '../../../constants/theme';
import { validateSetupModel } from '../lib/setup-model.js';
import { mutateSetupModel } from '../hooks/useSetupAutosave';
import type { DetectedField } from '../types';
import { SetupPdfFieldPreview } from './SetupPdfFieldPreview';

type SetupField = {
  fieldId: string;
  fieldName?: string;
  label?: string;
  required?: boolean;
  skipped?: boolean;
  multiline?: boolean;
  page?: number;
};

type SetupSingleSection = {
  sectionId: string;
  label?: string;
  fields?: SetupField[];
};

type SetupTableColumn = {
  columnId: string;
  label?: string;
  required?: boolean;
  skipped?: boolean;
  multiline?: boolean;
};

type SetupTableCell = {
  columnId: string;
  fieldId?: string;
  fieldName?: string;
  page?: number;
};

type SetupTableSection = {
  tableId: string;
  label?: string;
  columns?: SetupTableColumn[];
  rows?: Array<{ rowId: string; index?: number; cells?: SetupTableCell[] }>;
};

type Props = {
  templateName: string;
  templatePdfPath?: string | null;
  detectedFields?: DetectedField[];
  setupModel: Record<string, unknown>;
  onChange: (next: Record<string, unknown>) => void;
  info?: string | null;
  error?: string | null;
};

function fieldBadges(field: SetupField) {
  const badges = [];
  if (field.required) badges.push('Pflicht');
  if (field.skipped) badges.push('Ausgeblendet');
  if (field.multiline) badges.push('Mehrzeilig');
  return badges.join(' · ') || 'Standard';
}

function columnBadges(column: SetupTableColumn) {
  const badges = [];
  if (column.required) badges.push('Pflicht');
  if (column.skipped) badges.push('Ausgeblendet');
  if (column.multiline) badges.push('Mehrzeilig');
  return badges.join(' · ') || 'Standard';
}

export function SetupEditor({
  templateName,
  templatePdfPath,
  detectedFields = [],
  setupModel,
  onChange,
  info,
  error
}: Props) {
  const insets = useSafeAreaInsets();
  const singleSections = (setupModel.single_sections || []) as SetupSingleSection[];
  const tableSections = (setupModel.table_sections || []) as SetupTableSection[];
  const validationErrors = useMemo(() => validateSetupModel(setupModel), [setupModel]);

  const [mode, setMode] = useState<'single' | 'table'>(() =>
    singleSections.length > 0 ? 'single' : 'table'
  );
  const [activeSingleId, setActiveSingleId] = useState(singleSections[0]?.sectionId || '');
  const [activeTableId, setActiveTableId] = useState(tableSections[0]?.tableId || '');
  const [activeFieldId, setActiveFieldId] = useState<string | null>(null);
  const [activeColumnId, setActiveColumnId] = useState<string | null>(null);

  const activeSingle = singleSections.find((section) => section.sectionId === activeSingleId) || singleSections[0];
  const activeTable = tableSections.find((table) => table.tableId === activeTableId) || tableSections[0];

  const activeField = useMemo(() => {
    if (mode === 'single' && activeFieldId) {
      for (const section of singleSections) {
        const field = (section.fields || []).find((entry) => String(entry.fieldId) === String(activeFieldId));
        if (field) return field;
      }
    }
    if (mode === 'table' && activeTable && activeColumnId) {
      const cell = activeTable.rows?.[0]?.cells?.find((entry) => entry.columnId === activeColumnId);
      if (cell?.fieldId) {
        return {
          fieldId: cell.fieldId,
          fieldName: cell.fieldName,
          label:
            activeTable.columns?.find((column) => column.columnId === activeColumnId)?.label ||
            cell.fieldName ||
            activeColumnId
        };
      }
    }
    return null;
  }, [mode, activeFieldId, activeColumnId, activeTable, singleSections]);

  const activeFieldPage = useMemo(() => {
    if (!activeField) return 1;
    const detected = detectedFields.find((entry) => String(entry.fieldId) === String(activeField.fieldId));
    return Number(activeField.page || detected?.page || 1);
  }, [activeField, detectedFields]);

  const activeFieldLabel =
    activeField?.label || activeField?.fieldName || activeField?.fieldId || 'Feld in der Liste auswählen';

  useEffect(() => {
    if (mode !== 'single' || !activeSingle) return;
    const fields = activeSingle.fields || [];
    if (fields.length === 0) {
      setActiveFieldId(null);
      return;
    }
    if (!fields.some((field) => field.fieldId === activeFieldId)) {
      setActiveFieldId(fields[0].fieldId);
    }
  }, [mode, activeSingle, activeFieldId]);

  useEffect(() => {
    if (mode !== 'table' || !activeTable) return;
    const columns = activeTable.columns || [];
    if (columns.length === 0) {
      setActiveColumnId(null);
      return;
    }
    if (!columns.some((column) => column.columnId === activeColumnId)) {
      setActiveColumnId(columns[0].columnId);
    }
  }, [mode, activeTable, activeColumnId]);

  const updateModel = (mutator: (model: Record<string, unknown>) => void) => {
    onChange(mutateSetupModel(setupModel, mutator));
  };

  const updateSingleField = (sectionId: string, fieldId: string, patch: Partial<SetupField>) => {
    updateModel((model) => {
      for (const section of model.single_sections || []) {
        if (String(section.sectionId) !== String(sectionId)) continue;
        for (const field of section.fields || []) {
          if (String(field.fieldId) === String(fieldId)) {
            Object.assign(field, patch);
          }
        }
      }
    });
  };

  const updateSingleSectionLabel = (sectionId: string, label: string) => {
    updateModel((model) => {
      const section = (model.single_sections || []).find((entry) => String(entry.sectionId) === String(sectionId));
      if (section) section.label = label;
    });
  };

  const moveFieldWithinSection = (sectionId: string, fieldId: string, direction: -1 | 1) => {
    updateModel((model) => {
      const section = (model.single_sections || []).find((entry) => String(entry.sectionId) === String(sectionId));
      if (!section?.fields) return;
      const index = section.fields.findIndex((field) => String(field.fieldId) === String(fieldId));
      const nextIndex = index + direction;
      if (index < 0 || nextIndex < 0 || nextIndex >= section.fields.length) return;
      const fields = [...section.fields];
      const [moved] = fields.splice(index, 1);
      fields.splice(nextIndex, 0, moved);
      section.fields = fields;
    });
  };

  const moveFieldToSection = (fromSectionId: string, fieldId: string) => {
    const targets = singleSections.filter((section) => section.sectionId !== fromSectionId);
    if (targets.length === 0) return;
    Alert.alert(
      'Gruppe wählen',
      'Feld in andere Gruppe verschieben',
      targets
        .map((section) => ({
          text: section.label || section.sectionId,
          onPress: () => {
            updateModel((model) => {
              const sections = model.single_sections || [];
              const from = sections.find((entry) => String(entry.sectionId) === String(fromSectionId));
              const to = sections.find((entry) => String(entry.sectionId) === String(section.sectionId));
              if (!from?.fields || !to) return;
              const index = from.fields.findIndex((field) => String(field.fieldId) === String(fieldId));
              if (index < 0) return;
              const [field] = from.fields.splice(index, 1);
              to.fields = [...(to.fields || []), field];
            });
          }
        }))
        .concat([{ text: 'Abbrechen', style: 'cancel' }])
    );
  };

  const updateTableLabel = (tableId: string, label: string) => {
    updateModel((model) => {
      const table = (model.table_sections || []).find((entry) => String(entry.tableId) === String(tableId));
      if (table) table.label = label;
    });
  };

  const updateTableColumn = (tableId: string, columnId: string, patch: Partial<SetupTableColumn>) => {
    updateModel((model) => {
      const table = (model.table_sections || []).find((entry) => String(entry.tableId) === String(tableId));
      for (const column of table?.columns || []) {
        if (String(column.columnId) === String(columnId)) {
          Object.assign(column, patch);
        }
      }
      if ('multiline' in patch) {
        for (const row of table?.rows || []) {
          for (const cell of row.cells || []) {
            if (String(cell.columnId) === String(columnId)) {
              cell.multiline = patch.multiline === true;
            }
          }
        }
      }
    });
  };

  const selectedSingleField =
    mode === 'single'
      ? (activeSingle?.fields || []).find((field) => field.fieldId === activeFieldId) || null
      : null;
  const selectedTableColumn =
    mode === 'table'
      ? (activeTable?.columns || []).find((column) => column.columnId === activeColumnId) || null
      : null;

  return (
    <View style={styles.root}>
      {templatePdfPath ? (
        <View style={styles.previewPane}>
          <SetupPdfFieldPreview
            variant="pinned"
            pdfPath={templatePdfPath}
            detectedFields={detectedFields}
            activeFieldId={activeField?.fieldId || null}
            activeFieldLabel={activeFieldLabel}
            activeFieldPage={activeFieldPage}
          />
        </View>
      ) : null}

      <ScrollView
        style={styles.editorScroll}
        contentContainerStyle={[styles.editorContent, { paddingBottom: Math.max(insets.bottom, spacing.lg) }]}
        keyboardShouldPersistTaps="handled"
        showsVerticalScrollIndicator={false}
      >
        <View style={styles.intro}>
          <Text style={styles.templateName}>{templateName}</Text>
          <Text style={styles.muted}>
            Feld antippen — die Markierung in der PDF-Vorschau zeigt die Position im Formular.
          </Text>
        </View>

        {info ? <Text style={styles.info}>{info}</Text> : null}
        {error ? <Text style={styles.error}>{error}</Text> : null}
        {validationErrors.length > 0 ? (
          <Text style={styles.warn}>
            {validationErrors.length} Validierungsproblem(e): {validationErrors[0]}
          </Text>
        ) : null}

        <View style={styles.modeRow}>
          <PrimaryButton
            label="Gruppen"
            variant={mode === 'single' ? 'primary' : 'secondary'}
            onPress={() => setMode('single')}
            disabled={singleSections.length === 0}
          />
          <PrimaryButton
            label="Tabellen"
            variant={mode === 'table' ? 'primary' : 'secondary'}
            onPress={() => setMode('table')}
            disabled={tableSections.length === 0}
          />
        </View>

        {mode === 'single' ? (
          <>
            <Text style={styles.sectionLabel}>Gruppe wählen</Text>
            <ScrollView horizontal showsHorizontalScrollIndicator={false} contentContainerStyle={styles.chipRow}>
              {singleSections.map((section) => {
                const active = activeSingle?.sectionId === section.sectionId;
                return (
                  <Pressable
                    key={section.sectionId}
                    style={[styles.chip, active ? styles.chipActive : null]}
                    onPress={() => {
                      setMode('single');
                      setActiveSingleId(section.sectionId);
                    }}
                  >
                    <Text style={[styles.chipText, active ? styles.chipTextActive : null]} numberOfLines={1}>
                      {section.label || section.sectionId}
                    </Text>
                    <Text style={[styles.chipMeta, active ? styles.chipTextActive : null]}>
                      {(section.fields || []).length} Felder
                    </Text>
                  </Pressable>
                );
              })}
            </ScrollView>

            {activeSingle ? (
              <>
                <View style={styles.editorCard}>
                  <TextField
                    label="Gruppenname"
                    value={activeSingle.label || ''}
                    onChangeText={(value) => updateSingleSectionLabel(activeSingle.sectionId, value)}
                  />
                </View>

                <Text style={styles.sectionLabel}>Felder in „{activeSingle.label || activeSingle.sectionId}“</Text>
                {(activeSingle.fields || []).map((field) => {
                  const active = activeFieldId === field.fieldId;
                  return (
                    <Pressable
                      key={field.fieldId}
                      style={[styles.selectorCard, active ? styles.selectorCardActive : null]}
                      onPress={() => setActiveFieldId(field.fieldId)}
                    >
                      <Text style={styles.selectorTitle} numberOfLines={2}>
                        {field.label || field.fieldName || field.fieldId}
                      </Text>
                      <Text style={styles.selectorMeta}>
                        {field.fieldName || field.fieldId} · {fieldBadges(field)}
                      </Text>
                    </Pressable>
                  );
                })}

                {selectedSingleField ? (
                  <View style={styles.detailCard}>
                    <Text style={styles.detailTitle}>Feld bearbeiten</Text>
                    <Text style={styles.fieldMeta}>
                      PDF-Feld: {selectedSingleField.fieldName || selectedSingleField.fieldId}
                      {selectedSingleField.page ? ` · Seite ${selectedSingleField.page}` : ''}
                    </Text>
                    <TextField
                      label="Beschriftung"
                      value={selectedSingleField.label || ''}
                      onChangeText={(value) =>
                        updateSingleField(activeSingle.sectionId, selectedSingleField.fieldId, { label: value })
                      }
                      onFocus={() => setActiveFieldId(selectedSingleField.fieldId)}
                    />
                    <ListItem
                      title="Pflichtfeld"
                      subtitle={selectedSingleField.required ? 'Ja' : 'Nein'}
                      onPress={() =>
                        updateSingleField(activeSingle.sectionId, selectedSingleField.fieldId, {
                          required: !selectedSingleField.required
                        })
                      }
                    />
                    <ListItem
                      title="Überspringen"
                      subtitle={selectedSingleField.skipped ? 'Ausgeblendet' : 'Sichtbar'}
                      onPress={() =>
                        updateSingleField(activeSingle.sectionId, selectedSingleField.fieldId, {
                          skipped: !selectedSingleField.skipped
                        })
                      }
                    />
                    <ListItem
                      title="Mehrzeilig"
                      subtitle={selectedSingleField.multiline ? 'Ja · großes Eingabefeld' : 'Nein · einzeilig'}
                      onPress={() =>
                        updateSingleField(activeSingle.sectionId, selectedSingleField.fieldId, {
                          multiline: !selectedSingleField.multiline
                        })
                      }
                    />
                    <View style={styles.row}>
                      <PrimaryButton
                        label="↑"
                        variant="ghost"
                        disabled={
                          (activeSingle.fields || []).findIndex((field) => field.fieldId === selectedSingleField.fieldId) === 0
                        }
                        onPress={() =>
                          moveFieldWithinSection(activeSingle.sectionId, selectedSingleField.fieldId, -1)
                        }
                      />
                      <PrimaryButton
                        label="↓"
                        variant="ghost"
                        disabled={
                          (activeSingle.fields || []).findIndex((field) => field.fieldId === selectedSingleField.fieldId) ===
                          (activeSingle.fields?.length || 0) - 1
                        }
                        onPress={() =>
                          moveFieldWithinSection(activeSingle.sectionId, selectedSingleField.fieldId, 1)
                        }
                      />
                      <PrimaryButton
                        label="Verschieben"
                        variant="secondary"
                        onPress={() => moveFieldToSection(activeSingle.sectionId, selectedSingleField.fieldId)}
                      />
                    </View>
                  </View>
                ) : null}
              </>
            ) : null}
          </>
        ) : (
          <>
            <Text style={styles.sectionLabel}>Tabelle wählen</Text>
            <ScrollView horizontal showsHorizontalScrollIndicator={false} contentContainerStyle={styles.chipRow}>
              {tableSections.map((table) => {
                const active = activeTable?.tableId === table.tableId;
                return (
                  <Pressable
                    key={table.tableId}
                    style={[styles.chip, active ? styles.chipActive : null]}
                    onPress={() => {
                      setMode('table');
                      setActiveTableId(table.tableId);
                    }}
                  >
                    <Text style={[styles.chipText, active ? styles.chipTextActive : null]} numberOfLines={1}>
                      {table.label || table.tableId}
                    </Text>
                    <Text style={[styles.chipMeta, active ? styles.chipTextActive : null]}>
                      {(table.columns || []).length} Spalten
                    </Text>
                  </Pressable>
                );
              })}
            </ScrollView>

            {activeTable ? (
              <>
                <View style={styles.editorCard}>
                  <TextField
                    label="Tabellenname"
                    value={activeTable.label || ''}
                    onChangeText={(value) => updateTableLabel(activeTable.tableId, value)}
                  />
                </View>

                <Text style={styles.sectionLabel}>Spalten in „{activeTable.label || activeTable.tableId}“</Text>
                {(activeTable.columns || []).map((column) => {
                  const active = activeColumnId === column.columnId;
                  const previewCell = activeTable.rows?.[0]?.cells?.find((cell) => cell.columnId === column.columnId);
                  return (
                    <Pressable
                      key={column.columnId}
                      style={[styles.selectorCard, active ? styles.selectorCardActive : null]}
                      onPress={() => setActiveColumnId(column.columnId)}
                    >
                      <Text style={styles.selectorTitle} numberOfLines={2}>
                        {column.label || column.columnId}
                      </Text>
                      <Text style={styles.selectorMeta}>
                        {column.columnId}
                        {previewCell?.fieldName ? ` · ${previewCell.fieldName}` : ''} · {columnBadges(column)}
                      </Text>
                    </Pressable>
                  );
                })}

                {selectedTableColumn ? (
                  <View style={styles.detailCard}>
                    <Text style={styles.detailTitle}>Spalte bearbeiten</Text>
                    <Text style={styles.fieldMeta}>Spalte: {selectedTableColumn.columnId}</Text>
                    <TextField
                      label="Beschriftung"
                      value={selectedTableColumn.label || ''}
                      onChangeText={(value) =>
                        updateTableColumn(activeTable.tableId, selectedTableColumn.columnId, { label: value })
                      }
                      onFocus={() => setActiveColumnId(selectedTableColumn.columnId)}
                    />
                    <ListItem
                      title="Pflichtfeld"
                      subtitle={selectedTableColumn.required ? 'Ja' : 'Nein'}
                      onPress={() =>
                        updateTableColumn(activeTable.tableId, selectedTableColumn.columnId, {
                          required: !selectedTableColumn.required
                        })
                      }
                    />
                    <ListItem
                      title="Überspringen"
                      subtitle={selectedTableColumn.skipped ? 'Ausgeblendet' : 'Sichtbar'}
                      onPress={() =>
                        updateTableColumn(activeTable.tableId, selectedTableColumn.columnId, {
                          skipped: !selectedTableColumn.skipped
                        })
                      }
                    />
                    <ListItem
                      title="Mehrzeilig"
                      subtitle={selectedTableColumn.multiline ? 'Ja · großes Eingabefeld' : 'Nein · einzeilig'}
                      onPress={() =>
                        updateTableColumn(activeTable.tableId, selectedTableColumn.columnId, {
                          multiline: !selectedTableColumn.multiline
                        })
                      }
                    />
                  </View>
                ) : null}
              </>
            ) : null}
          </>
        )}

      </ScrollView>
    </View>
  );
}

const styles = StyleSheet.create({
  root: {
    flex: 1
  },
  previewPane: {
    borderBottomWidth: 1,
    borderBottomColor: colors.border,
    backgroundColor: colors.panel
  },
  editorScroll: {
    flex: 1
  },
  editorContent: {
    gap: spacing.sm,
    paddingTop: spacing.sm,
    paddingHorizontal: spacing.pageX
  },
  intro: {
    gap: 4
  },
  templateName: {
    ...typography.bodyStrong,
    color: colors.ink
  },
  muted: {
    ...typography.caption,
    color: colors.muted
  },
  info: {
    ...typography.caption,
    color: colors.accent
  },
  error: {
    ...typography.body,
    color: colors.danger
  },
  warn: {
    ...typography.caption,
    color: colors.accent2
  },
  modeRow: {
    flexDirection: 'row',
    gap: spacing.xs
  },
  sectionLabel: {
    ...typography.label,
    color: colors.muted,
    marginTop: spacing.xxs
  },
  chipRow: {
    gap: spacing.xs,
    paddingRight: spacing.sm
  },
  chip: {
    minWidth: 132,
    maxWidth: 220,
    borderWidth: 1,
    borderColor: colors.border,
    borderRadius: 12,
    paddingHorizontal: spacing.sm,
    paddingVertical: spacing.sm,
    backgroundColor: colors.panel,
    gap: 2
  },
  chipActive: {
    borderColor: colors.accent,
    backgroundColor: colors.badgeBg
  },
  chipText: {
    ...typography.bodyStrong,
    color: colors.ink
  },
  chipTextActive: {
    color: colors.accent
  },
  chipMeta: {
    ...typography.caption,
    color: colors.muted
  },
  editorCard: {
    gap: spacing.sm,
    padding: spacing.sm,
    borderRadius: 12,
    backgroundColor: colors.panel,
    borderWidth: 1,
    borderColor: colors.border
  },
  selectorCard: {
    gap: 4,
    padding: spacing.sm,
    borderRadius: 12,
    borderWidth: 1,
    borderColor: colors.border,
    backgroundColor: colors.panel
  },
  selectorCardActive: {
    borderColor: colors.accent,
    backgroundColor: 'rgba(47, 111, 237, 0.08)'
  },
  selectorTitle: {
    ...typography.bodyStrong,
    color: colors.ink
  },
  selectorMeta: {
    ...typography.caption,
    color: colors.muted
  },
  detailCard: {
    gap: spacing.xs,
    padding: spacing.sm,
    borderRadius: 12,
    borderWidth: 1,
    borderColor: colors.accent,
    backgroundColor: colors.panel
  },
  detailTitle: {
    ...typography.label,
    color: colors.accent2
  },
  fieldMeta: {
    ...typography.caption,
    color: colors.muted
  },
  row: {
    flexDirection: 'row',
    flexWrap: 'wrap',
    gap: spacing.xs,
    alignItems: 'center'
  }
});

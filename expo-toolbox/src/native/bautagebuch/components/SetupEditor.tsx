// @ts-nocheck
import { useEffect, useMemo, useState } from 'react';
import { Alert, Pressable, ScrollView, StyleSheet, Switch, Text, View } from 'react-native';
import { useSafeAreaInsets } from 'react-native-safe-area-context';

import { PrimaryButton, TextField } from '../../../components/mobile';
import { colors, spacing, typography } from '../../../constants/theme';
import { validateSetupModel, syncSectionOrder } from '../lib/setup-model.js';
import { mutateSetupModel } from '../hooks/useSetupAutosave';
import type { BautagebuchTemplate, DetectedField } from '../types';
import { SetupPdfFieldPreview } from './SetupPdfFieldPreview';
import { SetupTemplateManager } from './SetupTemplateManager';

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
  templates: BautagebuchTemplate[];
  activeTemplateId: string;
  editingTemplateId: string;
  importing?: boolean;
  onSelectEdit: (templateId: string) => void;
  onSetActive: (templateId: string) => void;
  onImport: () => void;
  templateName: string;
  templatePdfPath?: string | null;
  detectedFields?: DetectedField[];
  setupModel: Record<string, unknown>;
  onChange: (next: Record<string, unknown>) => void;
  info?: string | null;
  error?: string | null;
  embedded?: boolean;
  showPreview?: boolean;
};


function SettingRow({
  title,
  hint,
  value,
  onValueChange,
  disabled = false
}: {
  title: string;
  hint: string;
  value: boolean;
  onValueChange: (next: boolean) => void;
  disabled?: boolean;
}) {
  return (
    <View style={styles.settingRow}>
      <View style={styles.settingCopy}>
        <Text style={styles.settingTitle}>{title}</Text>
        <Text style={styles.settingHint}>{hint}</Text>
      </View>
      <Switch
        value={value}
        onValueChange={onValueChange}
        disabled={disabled}
        trackColor={{ false: colors.border, true: colors.accent }}
        thumbColor={colors.panel}
      />
    </View>
  );
}

function StatusPills({ items }: { items: string[] }) {
  if (items.length === 0) {
    return <Text style={styles.pillNeutral}>Standard</Text>;
  }
  return (
    <View style={styles.pillRow}>
      {items.map((item) => (
        <View key={item} style={styles.pill}>
          <Text style={styles.pillText}>{item}</Text>
        </View>
      ))}
    </View>
  );
}

export function SetupEditor({
  templates,
  activeTemplateId,
  editingTemplateId,
  importing = false,
  onSelectEdit,
  onSetActive,
  onImport,
  templateName,
  templatePdfPath,
  detectedFields = [],
  setupModel,
  onChange,
  info,
  error,
  embedded = false,
  showPreview = false
}: Props) {
  const insets = useSafeAreaInsets();
  const singleSections = (setupModel.single_sections || []) as SetupSingleSection[];
  const tableSections = (setupModel.table_sections || []) as SetupTableSection[];
  const sectionOrder = useMemo(() => syncSectionOrder(setupModel), [setupModel]);
  const orderedSections = useMemo(
    () =>
      sectionOrder.map((entry) => {
        if (entry.kind === 'single') {
          const section = singleSections.find((item) => item.sectionId === entry.id);
          return {
            ...entry,
            label: section?.label || entry.id,
            typeLabel: 'Gruppe'
          };
        }
        const table = tableSections.find((item) => item.tableId === entry.id);
        return {
          ...entry,
          label: table?.label || entry.id,
          typeLabel: 'Tabelle'
        };
      }),
    [sectionOrder, singleSections, tableSections]
  );
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

  const moveSectionInOrder = (index: number, direction: -1 | 1) => {
    updateModel((model) => {
      const order = syncSectionOrder(model);
      const nextIndex = index + direction;
      if (nextIndex < 0 || nextIndex >= order.length) return;
      const next = [...order];
      [next[index], next[nextIndex]] = [next[nextIndex], next[index]];
      model.section_order = next;
    });
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
      <ScrollView
        style={styles.editorScroll}
        contentContainerStyle={[styles.editorContent, { paddingBottom: Math.max(insets.bottom, spacing.lg) }]}
        keyboardShouldPersistTaps="handled"
        showsVerticalScrollIndicator={false}
      >
        <View style={styles.intro}>
          <Text style={styles.templateName}>{templateName}</Text>
          <Text style={styles.muted}>
            Abschnitt wählen, Feld antippen und Einstellungen direkt darunter anpassen.
          </Text>
        </View>

        {showPreview && templatePdfPath ? (
          <View style={styles.inlinePreview}>
            <SetupPdfFieldPreview
              variant="default"
              pdfPath={templatePdfPath}
              detectedFields={detectedFields}
              activeFieldId={activeField?.fieldId || null}
              activeFieldLabel={activeFieldLabel}
              activeFieldPage={activeFieldPage}
            />
          </View>
        ) : null}

        {!embedded ? (
          <SetupTemplateManager
            templates={templates}
            activeTemplateId={activeTemplateId}
            editingTemplateId={editingTemplateId}
            importing={importing}
            onSelectEdit={onSelectEdit}
            onSetActive={onSetActive}
            onImport={onImport}
          />
        ) : null}

        {info ? <Text style={styles.info}>{info}</Text> : null}
        {error ? <Text style={styles.error}>{error}</Text> : null}
        {validationErrors.length > 0 ? (
          <Text style={styles.warn}>
            {validationErrors.length} Validierungsproblem(e): {validationErrors[0]}
          </Text>
        ) : null}

        {orderedSections.length > 1 ? (
          <>
            <Text style={styles.sectionLabel}>Reihenfolge im Bautagebuch</Text>
            <Text style={styles.muted}>
              Gruppen und Tabellen frei anordnen — die Reihenfolge gilt im Assistenten und in der PDF-Vorschau.
            </Text>
            {orderedSections.map((entry, index) => {
              const selected =
                entry.kind === 'single'
                  ? mode === 'single' && activeSingle?.sectionId === entry.id
                  : mode === 'table' && activeTable?.tableId === entry.id;
              return (
              <Pressable
                key={`${entry.kind}:${entry.id}`}
                style={[styles.orderRow, selected ? styles.orderRowActive : null]}
                onPress={() => {
                  if (entry.kind === 'single') {
                    setMode('single');
                    setActiveSingleId(entry.id);
                  } else {
                    setMode('table');
                    setActiveTableId(entry.id);
                  }
                }}
              >
                <View style={styles.orderMeta}>
                  <Text style={styles.orderTitle} numberOfLines={1}>
                    {entry.label}
                  </Text>
                  <Text style={styles.orderType}>{entry.typeLabel}</Text>
                </View>
                <View style={styles.row}>
                  <PrimaryButton
                    label="↑"
                    variant="ghost"
                    disabled={index === 0}
                    onPress={() => moveSectionInOrder(index, -1)}
                  />
                  <PrimaryButton
                    label="↓"
                    variant="ghost"
                    disabled={index === orderedSections.length - 1}
                    onPress={() => moveSectionInOrder(index, 1)}
                  />
                </View>
              </Pressable>
            );
            })}
          </>
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
                      <Text style={styles.selectorMeta}>{field.fieldName || field.fieldId}</Text>
                      <StatusPills
                        items={[
                          field.required ? 'Pflicht' : null,
                          field.skipped ? 'Ausgeblendet' : null,
                          field.multiline ? 'Mehrzeilig' : null
                        ].filter(Boolean)}
                      />
                    </Pressable>
                  );
                })}

                {selectedSingleField ? (
                  <View style={styles.detailCard}>
                    <Text style={styles.detailTitle}>{selectedSingleField.label || selectedSingleField.fieldName}</Text>
                    <Text style={styles.fieldMeta}>
                      PDF-Feld: {selectedSingleField.fieldName || selectedSingleField.fieldId}
                      {selectedSingleField.page ? ` · Seite ${selectedSingleField.page}` : ''}
                    </Text>
                    <TextField
                      label="Anzeigename im Assistenten"
                      value={selectedSingleField.label || ''}
                      onChangeText={(value) =>
                        updateSingleField(activeSingle.sectionId, selectedSingleField.fieldId, { label: value })
                      }
                      onFocus={() => setActiveFieldId(selectedSingleField.fieldId)}
                    />
                    <SettingRow
                      title="Pflichtfeld"
                      hint="Muss vor dem Export ausgefüllt sein"
                      value={selectedSingleField.required === true}
                      onValueChange={(value) =>
                        updateSingleField(activeSingle.sectionId, selectedSingleField.fieldId, { required: value })
                      }
                    />
                    <SettingRow
                      title="Im Assistenten ausblenden"
                      hint="Feld wird nicht angezeigt, bleibt aber im PDF"
                      value={selectedSingleField.skipped === true}
                      onValueChange={(value) =>
                        updateSingleField(activeSingle.sectionId, selectedSingleField.fieldId, { skipped: value })
                      }
                    />
                    <SettingRow
                      title="Mehrzeiliges Eingabefeld"
                      hint="Größeres Textfeld für längere Einträge"
                      value={selectedSingleField.multiline === true}
                      onValueChange={(value) =>
                        updateSingleField(activeSingle.sectionId, selectedSingleField.fieldId, { multiline: value })
                      }
                    />
                    <Text style={styles.sectionLabel}>Reihenfolge in der Gruppe</Text>
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
                        {previewCell?.fieldName ? ` · ${previewCell.fieldName}` : ''}
                      </Text>
                      <StatusPills
                        items={[
                          column.required ? 'Pflicht' : null,
                          column.skipped ? 'Ausgeblendet' : null,
                          column.multiline ? 'Mehrzeilig' : null
                        ].filter(Boolean)}
                      />
                    </Pressable>
                  );
                })}

                {selectedTableColumn ? (
                  <View style={styles.detailCard}>
                    <Text style={styles.detailTitle}>{selectedTableColumn.label || selectedTableColumn.columnId}</Text>
                    <Text style={styles.fieldMeta}>Spalten-ID: {selectedTableColumn.columnId}</Text>
                    <TextField
                      label="Anzeigename im Assistenten"
                      value={selectedTableColumn.label || ''}
                      onChangeText={(value) =>
                        updateTableColumn(activeTable.tableId, selectedTableColumn.columnId, { label: value })
                      }
                      onFocus={() => setActiveColumnId(selectedTableColumn.columnId)}
                    />
                    <SettingRow
                      title="Pflichtfeld"
                      hint="Spalte muss in jeder sichtbaren Zeile ausgefüllt sein"
                      value={selectedTableColumn.required === true}
                      onValueChange={(value) =>
                        updateTableColumn(activeTable.tableId, selectedTableColumn.columnId, { required: value })
                      }
                    />
                    <SettingRow
                      title="Im Assistenten ausblenden"
                      hint="Spalte wird nicht angezeigt, bleibt aber im PDF"
                      value={selectedTableColumn.skipped === true}
                      onValueChange={(value) =>
                        updateTableColumn(activeTable.tableId, selectedTableColumn.columnId, { skipped: value })
                      }
                    />
                    <SettingRow
                      title="Mehrzeiliges Eingabefeld"
                      hint="Größeres Textfeld für längere Einträge"
                      value={selectedTableColumn.multiline === true}
                      onValueChange={(value) =>
                        updateTableColumn(activeTable.tableId, selectedTableColumn.columnId, { multiline: value })
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
  },
  orderRow: {
    flexDirection: 'row',
    alignItems: 'center',
    justifyContent: 'space-between',
    gap: spacing.sm,
    padding: spacing.sm,
    borderRadius: 12,
    borderWidth: 1,
    borderColor: colors.border,
    backgroundColor: colors.panel
  },
  orderRowActive: {
    borderColor: colors.accent,
    backgroundColor: 'rgba(47, 111, 237, 0.08)'
  },
  orderMeta: {
    flex: 1,
    gap: 2
  },
  orderTitle: {
    ...typography.bodyStrong,
    color: colors.ink
  },
  orderType: {
    ...typography.caption,
    color: colors.muted
  },
  inlinePreview: {
    borderRadius: 12,
    overflow: 'hidden'
  },
  settingRow: {
    flexDirection: 'row',
    alignItems: 'center',
    justifyContent: 'space-between',
    gap: spacing.sm,
    paddingVertical: spacing.xs
  },
  settingCopy: {
    flex: 1,
    gap: 2
  },
  settingTitle: {
    ...typography.bodyStrong,
    color: colors.ink
  },
  settingHint: {
    ...typography.caption,
    color: colors.muted
  },
  pillRow: {
    flexDirection: 'row',
    flexWrap: 'wrap',
    gap: spacing.xxs
  },
  pill: {
    borderRadius: 999,
    paddingHorizontal: spacing.sm,
    paddingVertical: 4,
    backgroundColor: colors.badgeBg
  },
  pillText: {
    ...typography.caption,
    color: colors.accent
  },
  pillNeutral: {
    ...typography.caption,
    color: colors.muted
  }
});

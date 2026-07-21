// @ts-nocheck
import { useMemo, useState } from 'react';
import { Alert, Pressable, StyleSheet, Text, View } from 'react-native';

import { ListItem, PrimaryButton, TextField } from '../../../components/mobile';
import { colors, typography } from '../../../constants/theme';
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
};

type SetupTableSection = {
  tableId: string;
  label?: string;
  columns?: SetupTableColumn[];
  rows?: Array<{ rowId: string; index?: number }>;
};

type Props = {
  templateName: string;
  templatePdfPath?: string | null;
  detectedFields?: DetectedField[];
  setupModel: Record<string, unknown>;
  onChange: (next: Record<string, unknown>) => void;
  onFinish: () => void;
  onPreview?: () => void;
  saving?: boolean;
  previewBusy?: boolean;
  info?: string | null;
  error?: string | null;
};

export function SetupEditor({
  templateName,
  templatePdfPath,
  detectedFields = [],
  setupModel,
  onChange,
  onFinish,
  onPreview,
  saving,
  previewBusy,
  info,
  error
}: Props) {
  const singleSections = (setupModel.single_sections || []) as SetupSingleSection[];
  const tableSections = (setupModel.table_sections || []) as SetupTableSection[];
  const validationErrors = useMemo(() => validateSetupModel(setupModel), [setupModel]);

  const [mode, setMode] = useState<'single' | 'table'>('single');
  const [showPdfPanel, setShowPdfPanel] = useState(Boolean(templatePdfPath));
  const [activeSingleId, setActiveSingleId] = useState(singleSections[0]?.sectionId || '');
  const [activeTableId, setActiveTableId] = useState(tableSections[0]?.tableId || '');
  const [activeFieldId, setActiveFieldId] = useState<string | null>(null);

  const activeSingle = singleSections.find((section) => section.sectionId === activeSingleId) || singleSections[0];
  const activeTable = tableSections.find((table) => table.tableId === activeTableId) || tableSections[0];

  const activeField = useMemo(() => {
    if (!activeFieldId) return null;
    for (const section of singleSections) {
      const field = (section.fields || []).find((entry) => String(entry.fieldId) === String(activeFieldId));
      if (field) return field;
    }
    return null;
  }, [activeFieldId, singleSections]);

  const activeFieldPage = useMemo(() => {
    if (!activeField) return 1;
    const detected = detectedFields.find((entry) => String(entry.fieldId) === String(activeField.fieldId));
    return Number(activeField.page || detected?.page || 1);
  }, [activeField, detectedFields]);

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
      targets.map((section) => ({
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
      })).concat([{ text: 'Abbrechen', style: 'cancel' }])
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
    });
  };

  return (
    <View style={styles.root}>
      <Text style={styles.title}>Setup: {templateName}</Text>
      <Text style={styles.muted}>
        Beschriftungen, Pflichtfelder und Sichtbarkeit anpassen. Änderungen werden automatisch gespeichert.
      </Text>

      {info ? <Text style={styles.info}>{info}</Text> : null}
      {error ? <Text style={styles.error}>{error}</Text> : null}
      {validationErrors.length > 0 ? (
        <Text style={styles.warn}>
          {validationErrors.length} Validierungsproblem(e): {validationErrors[0]}
        </Text>
      ) : null}

      {templatePdfPath ? (
        <View style={styles.previewBlock}>
          <View style={styles.row}>
            <Text style={styles.section}>Vorlagen-PDF</Text>
            <PrimaryButton
              label={showPdfPanel ? 'Ausblenden' : 'Anzeigen'}
              variant="ghost"
              onPress={() => setShowPdfPanel((value) => !value)}
            />
          </View>
          {showPdfPanel ? (
            <SetupPdfFieldPreview
              pdfPath={templatePdfPath}
              detectedFields={detectedFields}
              activeFieldId={activeFieldId}
              activeFieldLabel={activeField?.label || activeField?.fieldName || activeField?.fieldId || null}
              activeFieldPage={activeFieldPage}
            />
          ) : null}
        </View>
      ) : null}

      <View style={styles.row}>
        <PrimaryButton
          label="Gruppen"
          variant={mode === 'single' ? 'primary' : 'secondary'}
          onPress={() => setMode('single')}
        />
        <PrimaryButton
          label="Tabellen"
          variant={mode === 'table' ? 'primary' : 'secondary'}
          onPress={() => setMode('table')}
        />
        {onPreview ? (
          <PrimaryButton
            label={previewBusy ? 'Vorschau…' : 'PDF-Vorschau'}
            variant="ghost"
            disabled={previewBusy || saving}
            onPress={onPreview}
          />
        ) : null}
      </View>

      {mode === 'single' ? (
        <>
          <Text style={styles.section}>Gruppen</Text>
          {singleSections.map((section) => (
            <ListItem
              key={section.sectionId}
              title={section.label || section.sectionId}
              subtitle={`${section.fields?.length || 0} Felder`}
              meta={activeSingle?.sectionId === section.sectionId ? 'Aktiv' : undefined}
              onPress={() => setActiveSingleId(section.sectionId)}
            />
          ))}

          {activeSingle ? (
            <View style={styles.editorCard}>
              <TextField
                label="Gruppenname"
                value={activeSingle.label || ''}
                onChangeText={(value) => updateSingleSectionLabel(activeSingle.sectionId, value)}
              />
              {(activeSingle.fields || []).map((field, index) => (
                <PressableFieldCard
                  key={field.fieldId}
                  active={activeFieldId === field.fieldId}
                  onPress={() => setActiveFieldId(field.fieldId)}
                >
                  <Text style={styles.fieldMeta}>
                    PDF-Feld: {field.fieldName || field.fieldId}
                    {field.page ? ` · Seite ${field.page}` : ''}
                  </Text>
                  <TextField
                    label="Beschriftung"
                    value={field.label || ''}
                    onChangeText={(value) => updateSingleField(activeSingle.sectionId, field.fieldId, { label: value })}
                  />
                  <ListItem
                    title="Pflichtfeld"
                    subtitle={field.required ? 'Ja' : 'Nein'}
                    onPress={() =>
                      updateSingleField(activeSingle.sectionId, field.fieldId, { required: !field.required })
                    }
                  />
                  <ListItem
                    title="Überspringen"
                    subtitle={field.skipped ? 'Ausgeblendet' : 'Sichtbar'}
                    onPress={() =>
                      updateSingleField(activeSingle.sectionId, field.fieldId, { skipped: !field.skipped })
                    }
                  />
                  <View style={styles.row}>
                    <PrimaryButton
                      label="↑"
                      variant="ghost"
                      disabled={index === 0}
                      onPress={() => moveFieldWithinSection(activeSingle.sectionId, field.fieldId, -1)}
                    />
                    <PrimaryButton
                      label="↓"
                      variant="ghost"
                      disabled={index === (activeSingle.fields?.length || 0) - 1}
                      onPress={() => moveFieldWithinSection(activeSingle.sectionId, field.fieldId, 1)}
                    />
                    <PrimaryButton
                      label="Verschieben"
                      variant="secondary"
                      onPress={() => moveFieldToSection(activeSingle.sectionId, field.fieldId)}
                    />
                  </View>
                </PressableFieldCard>
              ))}
            </View>
          ) : null}
        </>
      ) : (
        <>
          <Text style={styles.section}>Tabellen</Text>
          {tableSections.map((table) => (
            <ListItem
              key={table.tableId}
              title={table.label || table.tableId}
              subtitle={`${table.rows?.length || 0} Zeilen · ${table.columns?.length || 0} Spalten`}
              meta={activeTable?.tableId === table.tableId ? 'Aktiv' : undefined}
              onPress={() => setActiveTableId(table.tableId)}
            />
          ))}

          {activeTable ? (
            <View style={styles.editorCard}>
              <TextField
                label="Tabellenname"
                value={activeTable.label || ''}
                onChangeText={(value) => updateTableLabel(activeTable.tableId, value)}
              />
              {(activeTable.columns || []).map((column) => (
                <View key={column.columnId} style={styles.fieldCard}>
                  <Text style={styles.fieldMeta}>Spalte: {column.columnId}</Text>
                  <TextField
                    label="Beschriftung"
                    value={column.label || ''}
                    onChangeText={(value) => updateTableColumn(activeTable.tableId, column.columnId, { label: value })}
                  />
                  <ListItem
                    title="Pflichtfeld"
                    subtitle={column.required ? 'Ja' : 'Nein'}
                    onPress={() =>
                      updateTableColumn(activeTable.tableId, column.columnId, { required: !column.required })
                    }
                  />
                  <ListItem
                    title="Überspringen"
                    subtitle={column.skipped ? 'Ausgeblendet' : 'Sichtbar'}
                    onPress={() =>
                      updateTableColumn(activeTable.tableId, column.columnId, { skipped: !column.skipped })
                    }
                  />
                </View>
              ))}
            </View>
          ) : null}
        </>
      )}

      <PrimaryButton
        label={saving ? 'Wird abgeschlossen…' : 'Setup abschließen'}
        disabled={saving || validationErrors.length > 0}
        onPress={onFinish}
      />
    </View>
  );
}

function PressableFieldCard({
  active,
  onPress,
  children
}: {
  active: boolean;
  onPress: () => void;
  children: React.ReactNode;
}) {
  return (
    <Pressable
      onPress={onPress}
      style={[styles.fieldCard, active ? styles.fieldCardActive : null]}
    >
      {children}
    </Pressable>
  );
}

const styles = StyleSheet.create({
  root: { gap: 12 },
  title: { ...typography.bodyStrong, color: colors.ink },
  muted: { ...typography.body, color: colors.muted },
  info: { ...typography.caption, color: colors.accent },
  error: { ...typography.body, color: colors.danger },
  warn: { ...typography.caption, color: colors.accent2 },
  section: { ...typography.label, color: colors.muted, marginTop: 4 },
  editorCard: { gap: 12, padding: 12, borderRadius: 12, backgroundColor: colors.panel, borderWidth: 1, borderColor: colors.border },
  fieldCard: {
    gap: 8,
    paddingTop: 8,
    paddingHorizontal: 4,
    borderTopWidth: 1,
    borderTopColor: colors.border
  },
  fieldCardActive: {
    borderRadius: 8,
    backgroundColor: 'rgba(47, 111, 237, 0.08)',
    borderTopColor: colors.accent
  },
  fieldMeta: { ...typography.caption, color: colors.muted },
  previewBlock: { gap: 8 },
  row: { flexDirection: 'row', flexWrap: 'wrap', gap: 8, alignItems: 'center' }
});

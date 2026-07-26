// @ts-nocheck
import { useEffect, useMemo, useRef } from 'react';
import { Image, Pressable, StyleSheet, Text, TextInput, View } from 'react-native';

import { Card, PrimaryButton, TextField } from '../../../components/mobile';
import { colors, spacing, typography } from '../../../constants/theme';
import { isClockTimeLabel, normalizeClockTime } from '../lib/time-format.js';
import { buildRunSections, inputKeyForField, requiredMissingCount, sectionProgressState } from '../lib/setup-model.js';
import {
  buildRunSectionsWithPhotoDoc,
  runSectionMissingCount,
  isPhotoDocRequiredMissing,
  sectionRunOptions,
  visibleRowCountForSection
} from '../lib/run-validation';
import { RunSectionNav } from './RunSectionNav';

type RunSection = ReturnType<typeof buildRunSections>[number];

type Props = {
  setupModel: Record<string, unknown>;
  values: Record<string, unknown>;
  sectionIndex: number;
  onChange: (nextValues: Record<string, unknown>) => void;
  onSectionChange: (index: number) => void;
  onWeatherSync?: () => Promise<void>;
  weatherBusy?: boolean;
  photoDoc?: { enabled: boolean | null; entries: Array<{ id: string; localPath?: string }> };
  onPhotoDocChange?: (enabled: boolean) => void;
  onAddPhoto?: () => void;
  onPickPhotos?: () => void;
  onRemovePhoto?: (entryId: string) => void;
  photoBusy?: boolean;
  totalMissingRequired?: number;
  onRequestExport?: () => void;
  showPreview?: boolean;
  previewPanel?: React.ReactNode;
};

const GEWERK_FIELDS = ['Text3', 'Text5', 'Text6', 'Text7', 'Text8'];
const SHIFT_FIELDS = ['Check Box1', 'Check Box2', 'Check Box3'];
const MAIN_PERSONAL_TABLE_ID = 'table_main_personal';

function fieldKey(fieldId: string) {
  return `field:${fieldId}`;
}

function cellKey(cellId: string) {
  return `cell:${cellId}`;
}

function findFieldIdByName(section: RunSection, fieldName: string): string {
  if (section.kind !== 'single') return '';
  const field = section.fields.find((entry) => entry.fieldName === fieldName);
  return String(field?.fieldId || '');
}

function isFieldMissing(section: RunSection, field: { fieldId: string; type?: string; required?: boolean }, values: Record<string, unknown>) {
  if (!field.required) return false;
  const key = fieldKey(field.fieldId);
  const value = values[key];
  if (field.type === 'checkbox') return value !== true;
  return String(value ?? '').trim().length === 0;
}

function isMainPersonalTable(section: RunSection) {
  if (section.kind !== 'table') return false;
  const tableId = String(section.tableId || '').trim();
  if (tableId === MAIN_PERSONAL_TABLE_ID) return true;
  const label = String(section.label || '').toLowerCase();
  return label.includes('firmen') || label.includes('personal') || label.includes('besetzung');
}

function computeFocusKeys(section: RunSection, values: Record<string, unknown>): string[] {
  if (!section) return [];

  if (section.kind === 'single') {
    const gewerkFieldIds = GEWERK_FIELDS.map((name) => findFieldIdByName(section, name)).filter(Boolean);
    const shiftFieldIds = SHIFT_FIELDS.map((name) => findFieldIdByName(section, name)).filter(Boolean);
    return section.fields
      .filter(
        (field) =>
          !gewerkFieldIds.includes(field.fieldId) &&
          !shiftFieldIds.includes(field.fieldId) &&
          field.type !== 'checkbox' &&
          !(field.type === 'dropdown' && field.options.length > 0)
      )
      .map((field) => fieldKey(field.fieldId));
  }

  if (section.kind === 'table') {
    const visibleCount = visibleRowCountForSection(section, values);
    const rows = section.rows.slice(0, Math.max(1, visibleCount));
    return rows.flatMap((row) =>
      row.cells
        .filter((cell) => {
          const column = section.columns.find((entry) => entry.columnId === cell.columnId);
          return column?.type !== 'checkbox' && !(column?.type === 'dropdown' && (column.options || []).length > 0);
        })
        .map((cell) => cellKey(cell.cellId))
    );
  }

  return [];
}

export function RunWizard({
  setupModel,
  values,
  sectionIndex,
  onChange,
  onSectionChange,
  onWeatherSync,
  weatherBusy,
  photoDoc,
  onPhotoDocChange,
  onAddPhoto,
  onPickPhotos,
  onRemovePhoto,
  photoBusy = false,
  totalMissingRequired = 0,
  onRequestExport,
  showPreview = false,
  previewPanel = null
}: Props) {
  const sections = useMemo(() => buildRunSectionsWithPhotoDoc(setupModel), [setupModel]);
  const section = sections[sectionIndex] || sections[0];
  const safeSectionIndex = sections.length > 0 ? Math.min(sectionIndex, sections.length - 1) : 0;
  const inputRefs = useRef(new Map<string, TextInput | null>());
  const focusKeys = useMemo(() => computeFocusKeys(section, values), [section, values]);

  const getFocusProps = (focusKey: string, multiline = false) => {
    const index = focusKeys.indexOf(focusKey);
    if (index < 0 || multiline) {
      return {};
    }
    const isLast = index === focusKeys.length - 1;
    return {
      returnKeyType: isLast ? 'done' : 'next',
      blurOnSubmit: isLast,
      onSubmitEditing: () => {
        if (!isLast) {
          inputRefs.current.get(focusKeys[index + 1])?.focus();
        }
      }
    };
  };

  useEffect(() => {
    if (sections.length === 0) return;
    if (sectionIndex > sections.length - 1) {
      onSectionChange(sections.length - 1);
    }
  }, [sectionIndex, sections.length, onSectionChange]);

  const setFieldValue = (fieldId: string, value: unknown) => {
    onChange({ ...values, [fieldKey(fieldId)]: value });
  };

  const setCellValue = (cellId: string, value: unknown) => {
    onChange({ ...values, [cellKey(cellId)]: value });
  };

  const renderSingleField = (field: RunSection['fields'][number]) => {
    const current = values[fieldKey(field.fieldId)];
    const label = field.label || field.fieldName;
    const missing = isFieldMissing(section, field, values);

    if (field.type === 'checkbox') {
      const checked = current === true;
      return (
        <Pressable
          key={field.fieldId}
          style={[styles.checkboxRow, missing ? styles.missingField : null]}
          onPress={() => setFieldValue(field.fieldId, !checked)}
        >
          <View style={[styles.checkbox, checked ? styles.checkboxOn : null]} />
          <Text style={styles.checkboxLabel}>{label}</Text>
        </Pressable>
      );
    }

    if (field.type === 'dropdown' && field.options.length > 0) {
      return (
        <View key={field.fieldId} style={[styles.choiceBlock, missing ? styles.missingField : null]}>
          <Text style={styles.label}>
            {label}
            {field.required ? ' *' : ''}
          </Text>
          <View style={styles.chipRow}>
            {field.options.map((option) => (
              <Pressable
                key={option}
                style={[styles.chip, current === option ? styles.chipActive : null]}
                onPress={() => setFieldValue(field.fieldId, option)}
              >
                <Text style={[styles.chipText, current === option ? styles.chipTextActive : null]}>
                  {option}
                </Text>
              </Pressable>
            ))}
          </View>
        </View>
      );
    }

    return (
      <View key={field.fieldId} style={missing ? styles.missingField : null}>
        <TextField
          ref={(node) => {
            if (node) inputRefs.current.set(fieldKey(field.fieldId), node);
            else inputRefs.current.delete(fieldKey(field.fieldId));
          }}
          label={`${label}${field.required ? ' *' : ''}`}
          value={String(current ?? '')}
          onChangeText={(text) => setFieldValue(field.fieldId, text)}
          onEndEditing={(event) => {
            if (isClockTimeLabel(label)) {
              setFieldValue(field.fieldId, normalizeClockTime(event.nativeEvent.text));
            }
          }}
          keyboardType={isClockTimeLabel(label) ? 'numbers-and-punctuation' : 'default'}
          multiline={field.multiline === true}
          autoGrow={field.multiline === true}
          {...getFocusProps(fieldKey(field.fieldId), field.multiline === true)}
        />
      </View>
    );
  };

  const renderHeaderSection = () => {
    if (section.kind !== 'single') return null;
    const gewerkFieldIds = GEWERK_FIELDS.map((name) => findFieldIdByName(section, name)).filter(Boolean);
    const shiftFieldIds = SHIFT_FIELDS.map((name) => findFieldIdByName(section, name)).filter(Boolean);
    const otherFields = section.fields.filter(
      (field) => !gewerkFieldIds.includes(field.fieldId) && !shiftFieldIds.includes(field.fieldId)
    );
    const gewerkMissing = requiredMissingCount(section, values, sectionRunOptions(section, values)) > 0;

    return (
      <View style={styles.sectionBody}>
        {otherFields.map(renderSingleField)}
        {gewerkFieldIds.length > 0 ? (
          <View style={[styles.groupPanel, gewerkMissing ? styles.missingField : null]}>
            <Text style={styles.groupTitle}>Gewerk</Text>
            <Text style={styles.groupHint}>Bitte genau ein Gewerk auswählen.</Text>
            <View style={styles.chipRow}>
              {section.fields
                .filter((field) => gewerkFieldIds.includes(field.fieldId))
                .map((field) => {
                  const active = values[fieldKey(field.fieldId)] === 'X';
                  return (
                    <Pressable
                      key={field.fieldId}
                      style={[styles.chip, active ? styles.chipActive : null]}
                      onPress={() => {
                        const next = { ...values };
                        for (const id of gewerkFieldIds) {
                          next[fieldKey(id)] = id === field.fieldId ? 'X' : '';
                        }
                        onChange(next);
                      }}
                    >
                      <Text style={[styles.chipText, active ? styles.chipTextActive : null]}>
                        {field.label}
                      </Text>
                    </Pressable>
                  );
                })}
            </View>
          </View>
        ) : null}
        {shiftFieldIds.length > 0 ? (
          <View style={styles.groupPanel}>
            <Text style={styles.groupTitle}>Schicht</Text>
            <Text style={styles.groupHint}>Bitte genau eine Schicht auswählen.</Text>
            <View style={styles.groupChoices}>
              {section.fields
                .filter((field) => shiftFieldIds.includes(field.fieldId))
                .map((field) => {
                  const checked = values[fieldKey(field.fieldId)] === true;
                  return (
                    <Pressable
                      key={field.fieldId}
                      style={[styles.shiftChip, checked ? styles.shiftChipActive : null]}
                      onPress={() => {
                        const next = { ...values };
                        for (const id of shiftFieldIds) {
                          next[fieldKey(id)] = id === field.fieldId ? !checked : false;
                        }
                        onChange(next);
                      }}
                    >
                      <View style={[styles.checkbox, checked ? styles.checkboxOn : null]} />
                      <Text style={[styles.shiftChipLabel, checked ? styles.shiftChipLabelActive : null]}>
                        {field.label}
                      </Text>
                    </Pressable>
                  );
                })}
            </View>
          </View>
        ) : null}
      </View>
    );
  };

  const renderWeatherSection = () => {
    if (section.kind !== 'single') return null;
    return (
      <View style={styles.sectionBody}>
        {section.fields.map(renderSingleField)}
        {onWeatherSync ? (
          <PrimaryButton
            label={weatherBusy ? 'Wetter wird geladen…' : 'Wetter aktualisieren'}
            variant="secondary"
            disabled={weatherBusy}
            onPress={() => void onWeatherSync()}
          />
        ) : null}
      </View>
    );
  };

  const renderTableSection = () => {
    if (section.kind !== 'table') return null;
    const tableId = String(section.tableId || '');
    const visibleCount = visibleRowCountForSection(section, values);
    const rows = section.rows.slice(0, Math.max(1, visibleCount));
    const personal = isMainPersonalTable(section);

    return (
      <View style={styles.sectionBody}>
        {rows.map((row) => (
          <Card key={row.rowId} style={styles.tableCard}>
            <Text style={styles.tableRowTitle}>Zeile {row.rowId.replace('r', '')}</Text>
            {row.cells.map((cell) => {
              const label = cell.label;
              const column = section.columns.find((entry) => entry.columnId === cell.columnId);
              const multiline = column?.multiline === true || cell.multiline === true;
              const isTime = personal && isClockTimeLabel(label);
              const missing =
                cell.required &&
                String(values[cellKey(cell.cellId)] ?? '').trim().length === 0 &&
                row.rowId === rows[0]?.rowId;

              return (
                <View key={cell.cellId} style={missing ? styles.missingField : null}>
                  <TextField
                    ref={(node) => {
                      if (node) inputRefs.current.set(cellKey(cell.cellId), node);
                      else inputRefs.current.delete(cellKey(cell.cellId));
                    }}
                    label={`${label}${cell.required ? ' *' : ''}`}
                    value={String(values[cellKey(cell.cellId)] ?? '')}
                    onChangeText={(text) => setCellValue(cell.cellId, text)}
                    onEndEditing={(event) => {
                      if (isTime) {
                        setCellValue(cell.cellId, normalizeClockTime(event.nativeEvent.text));
                      }
                    }}
                    keyboardType={isTime ? 'numbers-and-punctuation' : 'default'}
                    multiline={multiline}
                    autoGrow={multiline}
                    {...getFocusProps(cellKey(cell.cellId), multiline)}
                  />
                </View>
              );
            })}
          </Card>
        ))}
        {visibleCount < section.rows.length ? (
          <PrimaryButton
            label="Weitere Zeile hinzufügen"
            variant="secondary"
            onPress={() =>
              onChange({ ...values, [`__tableRows:${tableId}`]: visibleCount + 1 })
            }
          />
        ) : null}
      </View>
    );
  };

  const renderPhotoDocSection = () => {
    const choiceMissing = isPhotoDocRequiredMissing(photoDoc?.enabled ?? null);
    const photoCount = photoDoc?.entries?.length || 0;
    const photosMissing = photoDoc?.enabled === true && photoCount === 0;
    return (
      <View style={styles.sectionBody}>
        <View style={choiceMissing ? styles.missingField : null}>
          <Text style={styles.label}>Fotodokumentation anhängen? *</Text>
          <Text style={styles.hint}>
            Wähle Ja, wenn Baustellenfotos als separates PDF oder zusammen mit dem BTB exportiert werden sollen.
          </Text>
          <View style={styles.chipRow}>
            <Pressable
              style={[styles.chip, photoDoc?.enabled === true ? styles.chipActive : null]}
              onPress={() => onPhotoDocChange?.(true)}
            >
              <Text style={[styles.chipText, photoDoc?.enabled === true ? styles.chipTextActive : null]}>Ja</Text>
            </Pressable>
            <Pressable
              style={[styles.chip, photoDoc?.enabled === false ? styles.chipActive : null]}
              onPress={() => onPhotoDocChange?.(false)}
            >
              <Text style={[styles.chipText, photoDoc?.enabled === false ? styles.chipTextActive : null]}>Nein</Text>
            </Pressable>
          </View>
        </View>
        {photoDoc?.enabled ? (
          <>
            <View style={[styles.photoSummary, photosMissing ? styles.missingField : null]}>
              <Text style={styles.label}>
                Fotos ({photoCount})
                {photosMissing ? ' · mindestens 1 Foto empfohlen' : ''}
              </Text>
              <Text style={styles.hint}>
                Nimm ein Foto auf oder wähle mehrere Bilder aus der Galerie. Du kannst beliebig viele Fotos hinzufügen.
              </Text>
            </View>
            <View style={styles.photoActions}>
              <PrimaryButton
                label={photoBusy ? 'Wird gespeichert…' : 'Foto aufnehmen'}
                variant="secondary"
                disabled={photoBusy}
                onPress={() => onAddPhoto?.()}
              />
              <PrimaryButton
                label={photoBusy ? 'Wird geladen…' : 'Mehrere aus Galerie'}
                variant="secondary"
                disabled={photoBusy}
                onPress={() => onPickPhotos?.()}
              />
            </View>
            {(photoDoc.entries || []).map((entry, index) => (
              <View key={entry.id} style={styles.photoRow}>
                <Text style={styles.photoMeta}>Foto {photoCount - index}</Text>
                {entry.localPath ? <Image source={{ uri: entry.localPath }} style={styles.photoPreview} /> : null}
                <PrimaryButton
                  label="Entfernen"
                  variant="ghost"
                  disabled={photoBusy}
                  onPress={() => onRemovePhoto?.(entry.id)}
                />
              </View>
            ))}
            {photoCount === 0 ? (
              <Text style={styles.hint}>Noch keine Fotos vorhanden. Tippe auf „Foto aufnehmen“ oder „Mehrere aus Galerie“.</Text>
            ) : null}
          </>
        ) : null}
      </View>
    );
  };

  const renderSectionContent = () => {
    if (!section) return null;
    if (section.kind === 'photo-doc' || section.sectionId === 'photo-doc') return renderPhotoDocSection();
    if (section.sectionId === 'single:header') return renderHeaderSection();
    if (section.sectionId === 'single:weather') return renderWeatherSection();
    if (section.kind === 'table') return renderTableSection();
    if (section.kind === 'single') {
      return <View style={styles.sectionBody}>{section.fields.map(renderSingleField)}</View>;
    }
    return null;
  };

  const isLastSection = safeSectionIndex >= sections.length - 1;
  const exportBlocked = totalMissingRequired > 0;

  const sectionNavItems = sections.map((entry) => {
    const options = entry.kind === 'photo-doc' ? {} : sectionRunOptions(entry, values);
    const progress =
      entry.kind === 'photo-doc'
        ? isPhotoDocRequiredMissing(photoDoc?.enabled ?? null)
          ? 'todo'
          : photoDoc?.enabled === true && (photoDoc.entries || []).length === 0
            ? 'progress'
            : 'done'
        : sectionProgressState(entry, values, options);
    const missingCount =
      entry.kind === 'photo-doc'
        ? isPhotoDocRequiredMissing(photoDoc?.enabled ?? null)
          ? 1
          : 0
        : runSectionMissingCount(entry, values);

    return {
      sectionId: entry.sectionId,
      label: entry.label,
      progress,
      missingCount
    };
  });

  return (
    <View style={styles.root}>
      <RunSectionNav
        sections={sectionNavItems}
        activeIndex={safeSectionIndex}
        totalMissingRequired={totalMissingRequired}
        onSelect={onSectionChange}
      />

      <Card>
        <View style={styles.sectionHeader}>
          <View style={styles.sectionStepBadge}>
            <Text style={styles.sectionStepText}>{safeSectionIndex + 1}</Text>
          </View>
          <View style={styles.sectionHeaderText}>
            <Text style={styles.sectionEyebrow}>
              Abschnitt {safeSectionIndex + 1} von {sections.length}
            </Text>
            <Text style={styles.sectionTitle}>{section?.label}</Text>
          </View>
        </View>
        {renderSectionContent()}
      </Card>

      {showPreview && previewPanel ? <View style={styles.previewBlock}>{previewPanel}</View> : null}

      <View style={styles.footerRow}>
        <PrimaryButton
          label="Zurück"
          variant="secondary"
          disabled={safeSectionIndex <= 0}
          onPress={() => onSectionChange(Math.max(0, safeSectionIndex - 1))}
        />
        <PrimaryButton
          label={isLastSection ? 'Abschließen' : 'Weiter'}
          disabled={isLastSection && exportBlocked}
          onPress={() => {
            if (isLastSection) {
              onRequestExport?.();
              return;
            }
            onSectionChange(Math.min(sections.length - 1, safeSectionIndex + 1));
          }}
        />
      </View>
    </View>
  );
}

const styles = StyleSheet.create({
  root: { gap: spacing.md },
  sectionHeader: {
    flexDirection: 'row',
    alignItems: 'flex-start',
    gap: spacing.sm,
    marginBottom: spacing.md,
    paddingBottom: spacing.sm,
    borderBottomWidth: 1,
    borderBottomColor: colors.border
  },
  sectionStepBadge: {
    width: 36,
    height: 36,
    borderRadius: 18,
    backgroundColor: colors.badgeBg,
    borderWidth: 1,
    borderColor: colors.accent,
    alignItems: 'center',
    justifyContent: 'center'
  },
  sectionStepText: {
    ...typography.bodyStrong,
    color: colors.accent
  },
  sectionHeaderText: {
    flex: 1,
    gap: 2
  },
  sectionEyebrow: {
    ...typography.caption,
    color: colors.muted
  },
  sectionTitle: { ...typography.subtitle, color: colors.ink },
  sectionBody: { gap: spacing.sm },
  groupPanel: {
    gap: spacing.xs,
    borderWidth: 1,
    borderColor: colors.border,
    borderRadius: spacing.inputRadius,
    backgroundColor: colors.bg,
    padding: spacing.sm
  },
  groupTitle: { ...typography.bodyStrong, color: colors.ink },
  groupHint: { ...typography.caption, color: colors.muted },
  groupChoices: { gap: spacing.xs },
  label: { ...typography.label, color: colors.muted },
  hint: { ...typography.caption, color: colors.muted },
  choiceBlock: { gap: spacing.xs },
  chipRow: { flexDirection: 'row', flexWrap: 'wrap', gap: spacing.xs },
  chip: {
    borderWidth: 1,
    borderColor: colors.border,
    borderRadius: 999,
    paddingHorizontal: spacing.sm,
    paddingVertical: 10,
    minHeight: 44,
    justifyContent: 'center',
    backgroundColor: colors.panel
  },
  chipActive: { borderColor: colors.accent, backgroundColor: colors.badgeBg },
  chipText: { ...typography.caption, color: colors.ink },
  chipTextActive: { color: colors.accent, fontFamily: 'SpaceGrotesk_600SemiBold' },
  shiftChip: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: spacing.sm,
    minHeight: 48,
    borderWidth: 1,
    borderColor: colors.border,
    borderRadius: spacing.inputRadius,
    paddingHorizontal: spacing.sm,
    paddingVertical: spacing.xs,
    backgroundColor: colors.panelElevated
  },
  shiftChipActive: {
    borderColor: colors.accent,
    backgroundColor: colors.badgeBg
  },
  shiftChipLabel: { ...typography.body, color: colors.ink, flex: 1 },
  shiftChipLabelActive: { color: colors.accent, fontFamily: 'SpaceGrotesk_600SemiBold' },
  checkboxRow: { flexDirection: 'row', alignItems: 'center', gap: spacing.sm, minHeight: 48 },
  checkbox: {
    width: 24,
    height: 24,
    borderRadius: 6,
    borderWidth: 1,
    borderColor: colors.border,
    backgroundColor: colors.white
  },
  checkboxOn: { backgroundColor: colors.accent, borderColor: colors.accent },
  checkboxLabel: { ...typography.body, color: colors.ink },
  tableCard: { gap: spacing.sm },
  tableRowTitle: { ...typography.bodyStrong, color: colors.ink },
  photoActions: { flexDirection: 'row', gap: spacing.sm, flexWrap: 'wrap' },
  photoSummary: { gap: spacing.xxs },
  photoRow: { gap: spacing.xs },
  photoMeta: { ...typography.caption, color: colors.muted },
  photoPreview: { width: '100%', height: 180, borderRadius: 12, backgroundColor: colors.border },
  missingField: {
    borderWidth: 1,
    borderColor: colors.danger,
    borderRadius: 12,
    padding: 6,
    backgroundColor: 'rgba(161, 44, 36, 0.05)'
  },
  footerRow: { flexDirection: 'row', gap: spacing.sm },
  previewBlock: { gap: spacing.sm }
});

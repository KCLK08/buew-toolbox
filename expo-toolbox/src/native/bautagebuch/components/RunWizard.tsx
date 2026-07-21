// @ts-nocheck
import { useMemo } from 'react';
import { Image, Pressable, StyleSheet, Text, View } from 'react-native';

import { Card, PrimaryButton, TextField } from '../../../components/mobile';
import { colors, spacing, typography } from '../../../constants/theme';
import { normalizeClockTime } from '../lib/time-format.js';
import { buildRunSections, sectionProgressState } from '../lib/setup-model.js';
import { RunValuesPreview } from './RunValuesPreview';

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
  onRemovePhoto?: (entryId: string) => void;
};

const GEWERK_FIELDS = ['Text3', 'Text5', 'Text6', 'Text7', 'Text8'];
const SHIFT_FIELDS = ['Check Box1', 'Check Box2', 'Check Box3'];

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
  onRemovePhoto
}: Props) {
  const sections = useMemo(() => {
    const base = buildRunSections(setupModel);
    return [...base, { sectionId: 'photo-doc', kind: 'photo-doc', label: 'Fotodokumentation' }];
  }, [setupModel]);
  const section = sections[sectionIndex] || sections[0];

  const setFieldValue = (fieldId: string, value: unknown) => {
    onChange({ ...values, [fieldKey(fieldId)]: value });
  };

  const setCellValue = (cellId: string, value: unknown) => {
    onChange({ ...values, [cellKey(cellId)]: value });
  };

  const renderSingleField = (field: RunSection['fields'][number]) => {
    const current = values[fieldKey(field.fieldId)];
    const label = field.label || field.fieldName;

    if (field.type === 'checkbox') {
      const checked = current === true;
      return (
        <Pressable
          key={field.fieldId}
          style={styles.checkboxRow}
          onPress={() => setFieldValue(field.fieldId, !checked)}
        >
          <View style={[styles.checkbox, checked ? styles.checkboxOn : null]} />
          <Text style={styles.checkboxLabel}>{label}</Text>
        </Pressable>
      );
    }

    if (field.type === 'dropdown' && field.options.length > 0) {
      return (
        <View key={field.fieldId} style={styles.choiceBlock}>
          <Text style={styles.label}>{label}</Text>
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
      <TextField
        key={field.fieldId}
        label={label}
        value={String(current ?? '')}
        onChangeText={(text) => {
          const normalized =
            label.toLowerCase().includes('beginn') || label.toLowerCase().includes('ende')
              ? normalizeClockTime(text)
              : text;
          setFieldValue(field.fieldId, normalized);
        }}
        multiline={label.length > 30}
      />
    );
  };

  const renderHeaderSection = () => {
    if (section.kind !== 'single') return null;
    const gewerkFieldIds = GEWERK_FIELDS.map((name) => findFieldIdByName(section, name)).filter(Boolean);
    const shiftFieldIds = SHIFT_FIELDS.map((name) => findFieldIdByName(section, name)).filter(Boolean);
    const otherFields = section.fields.filter(
      (field) => !gewerkFieldIds.includes(field.fieldId) && !shiftFieldIds.includes(field.fieldId)
    );

    return (
      <View style={styles.sectionBody}>
        {otherFields.map(renderSingleField)}
        {gewerkFieldIds.length > 0 ? (
          <View style={styles.choiceBlock}>
            <Text style={styles.label}>Gewerk</Text>
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
          <View style={styles.choiceBlock}>
            <Text style={styles.label}>Schicht</Text>
            {section.fields
              .filter((field) => shiftFieldIds.includes(field.fieldId))
              .map((field) => renderSingleField(field))}
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
    const rowCountKey = `__tableRows:${tableId}`;
    const visibleCount = Number(values[rowCountKey] ?? 1);
    const rows = section.rows.slice(0, Math.max(1, visibleCount));

    return (
      <View style={styles.sectionBody}>
        {rows.map((row) => (
          <Card key={row.rowId} style={styles.tableCard}>
            <Text style={styles.tableRowTitle}>Zeile {row.rowId.replace('r', '')}</Text>
            {row.cells.map((cell) => (
              <TextField
                key={cell.cellId}
                label={cell.label}
                value={String(values[cellKey(cell.cellId)] ?? '')}
                onChangeText={(text) => {
                  const normalized =
                    cell.label.toLowerCase().includes('beginn') ||
                    cell.label.toLowerCase().includes('ende')
                      ? normalizeClockTime(text)
                      : text;
                  setCellValue(cell.cellId, normalized);
                }}
                multiline={cell.label.length > 24}
              />
            ))}
          </Card>
        ))}
        {visibleCount < section.rows.length ? (
          <PrimaryButton
            label="Weitere Zeile hinzufügen"
            variant="secondary"
            onPress={() => onChange({ ...values, [rowCountKey]: visibleCount + 1 })}
          />
        ) : null}
      </View>
    );
  };

  const renderPhotoDocSection = () => (
    <View style={styles.sectionBody}>
      <Text style={styles.label}>Fotodokumentation anhängen?</Text>
      <View style={styles.chipRow}>
        <Pressable
          style={[styles.chip, photoDoc?.enabled === true ? styles.chipActive : null]}
          onPress={() => onPhotoDocChange?.(true)}
        >
          <Text style={styles.chipText}>Ja</Text>
        </Pressable>
        <Pressable
          style={[styles.chip, photoDoc?.enabled === false ? styles.chipActive : null]}
          onPress={() => onPhotoDocChange?.(false)}
        >
          <Text style={styles.chipText}>Nein</Text>
        </Pressable>
      </View>
      {photoDoc?.enabled ? (
        <>
          <PrimaryButton label="+ Foto aufnehmen" variant="secondary" onPress={() => onAddPhoto?.()} />
          {(photoDoc.entries || []).map((entry) => (
            <View key={entry.id} style={styles.photoRow}>
              {entry.localPath ? <Image source={{ uri: entry.localPath }} style={styles.photoPreview} /> : null}
              <PrimaryButton
                label="Entfernen"
                variant="ghost"
                onPress={() => onRemovePhoto?.(entry.id)}
              />
            </View>
          ))}
        </>
      ) : null}
    </View>
  );

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

  return (
    <View style={styles.root}>
      <View style={styles.navRow}>
        {sections.map((entry, index) => {
          const progress = sectionProgressState(entry, values, {
            visibleRowCount: Number(values[`__tableRows:${entry.tableId || ''}`] ?? 1)
          });
          return (
            <Pressable key={entry.sectionId} style={styles.navDotWrap} onPress={() => onSectionChange(index)}>
              <View
                style={[
                  styles.navDot,
                  index === sectionIndex ? styles.navDotActive : null,
                  progress === 'done' ? styles.navDotDone : null
                ]}
              />
              <Text style={styles.navLabel} numberOfLines={1}>
                {entry.label}
              </Text>
            </Pressable>
          );
        })}
      </View>

      <Card>
        <Text style={styles.sectionTitle}>{section?.label}</Text>
        {renderSectionContent()}
      </Card>

      <RunValuesPreview setupModel={setupModel} values={values} sectionIndex={sectionIndex} />

      <View style={styles.footerRow}>
        <PrimaryButton
          label="Zurück"
          variant="secondary"
          disabled={sectionIndex <= 0}
          onPress={() => onSectionChange(Math.max(0, sectionIndex - 1))}
        />
        <PrimaryButton
          label={sectionIndex >= sections.length - 1 ? 'Fertig' : 'Weiter'}
          onPress={() => onSectionChange(Math.min(sections.length - 1, sectionIndex + 1))}
        />
      </View>
    </View>
  );
}

const styles = StyleSheet.create({
  root: { gap: spacing.md },
  navRow: { flexDirection: 'row', flexWrap: 'wrap', gap: spacing.sm },
  navDotWrap: { alignItems: 'center', width: 72 },
  navDot: {
    width: 12,
    height: 12,
    borderRadius: 6,
    backgroundColor: colors.border,
    marginBottom: 4
  },
  navDotActive: { backgroundColor: colors.accent },
  navDotDone: { backgroundColor: colors.accent2 },
  navLabel: { ...typography.caption, color: colors.muted, textAlign: 'center' },
  sectionTitle: { ...typography.subtitle, color: colors.ink, marginBottom: spacing.sm },
  sectionBody: { gap: spacing.sm },
  label: { ...typography.label, color: colors.muted },
  choiceBlock: { gap: spacing.xs },
  chipRow: { flexDirection: 'row', flexWrap: 'wrap', gap: spacing.xs },
  chip: {
    borderWidth: 1,
    borderColor: colors.border,
    borderRadius: 999,
    paddingHorizontal: spacing.sm,
    paddingVertical: 8,
    backgroundColor: colors.panel
  },
  chipActive: { borderColor: colors.accent, backgroundColor: colors.badgeBg },
  chipText: { ...typography.caption, color: colors.ink },
  chipTextActive: { color: colors.accent, fontFamily: 'SpaceGrotesk_600SemiBold' },
  checkboxRow: { flexDirection: 'row', alignItems: 'center', gap: spacing.sm, minHeight: 44 },
  checkbox: {
    width: 22,
    height: 22,
    borderRadius: 6,
    borderWidth: 1,
    borderColor: colors.border,
    backgroundColor: colors.white
  },
  checkboxOn: { backgroundColor: colors.accent, borderColor: colors.accent },
  checkboxLabel: { ...typography.body, color: colors.ink },
  tableCard: { gap: spacing.sm },
  tableRowTitle: { ...typography.bodyStrong, color: colors.ink },
  photoRow: { gap: spacing.xs },
  photoPreview: { width: '100%', height: 160, borderRadius: 12, backgroundColor: colors.border },
  footerRow: { flexDirection: 'row', gap: spacing.sm }
});

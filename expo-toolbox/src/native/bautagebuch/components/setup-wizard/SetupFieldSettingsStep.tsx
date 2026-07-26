import { useEffect, useMemo, useState } from 'react';
import { ScrollView, StyleSheet, Text, View, useWindowDimensions } from 'react-native';
import { useSafeAreaInsets } from 'react-native-safe-area-context';

import { colors, spacing, typography } from '../../../../constants/theme';
import { listSetupSections, updateSetupField } from '../../lib/setup-mapping';
import type { DetectedField, SetupFieldConfig } from '../../types';
import { PreviewOverlayPanel } from '../PreviewOverlayPanel';
import { SetupPdfFieldPreview } from '../SetupPdfFieldPreview';
import { SetupFieldCard } from './SetupFieldCard';
import { SetupGroupNav } from './SetupGroupNav';
import { SetupValidationList } from './SetupValidationList';

type Props = {
  pdfPath: string | null;
  detectedFields: DetectedField[];
  setupModel: Record<string, unknown>;
  validationIssues: string[];
  readOnly?: boolean;
  showPreview?: boolean;
  onClosePreview?: () => void;
  onChange: (next: Record<string, unknown>) => void;
};

export function SetupFieldSettingsStep({
  pdfPath,
  detectedFields,
  setupModel,
  validationIssues,
  readOnly = false,
  showPreview = false,
  onClosePreview,
  onChange
}: Props) {
  const insets = useSafeAreaInsets();
  const { width } = useWindowDimensions();
  const isTablet = width >= 900;
  const sections = useMemo(() => listSetupSections(setupModel), [setupModel]);
  const [activeSectionId, setActiveSectionId] = useState(sections[0]?.sectionId || '');
  const [activeFieldId, setActiveFieldId] = useState<string | null>(
    sections[0]?.fields[0]?.fieldId || null
  );

  const activeSection = sections.find((section) => section.sectionId === activeSectionId) || sections[0];
  const activeField =
    activeSection?.fields.find((field) => field.fieldId === activeFieldId) ||
    activeSection?.fields[0] ||
    null;

  const groupItems = sections.map((section) => ({
    sectionId: section.sectionId,
    label: section.label,
    count: section.fields.length
  }));

  useEffect(() => {
    if (sections.length === 0) return;
    const sectionExists = sections.some((section) => section.sectionId === activeSectionId);
    if (!sectionExists) {
      setActiveSectionId(sections[0].sectionId);
      setActiveFieldId(sections[0].fields[0]?.fieldId || null);
    }
  }, [sections, activeSectionId]);

  const handleFieldChange = (field: SetupFieldConfig, patch: Partial<SetupFieldConfig>) => {
    setActiveFieldId(field.fieldId);
    onChange(updateSetupField(setupModel, activeSection?.sectionId || '', field.fieldId, patch));
  };

  return (
    <View style={styles.root}>
      <View style={[styles.body, isTablet ? styles.bodyTablet : null]}>
        <View style={isTablet ? styles.sideNav : null}>
          <SetupGroupNav
            groups={groupItems}
            activeSectionId={activeSection?.sectionId || ''}
            horizontal={!isTablet}
            onSelect={(sectionId) => {
              setActiveSectionId(sectionId);
              const section = sections.find((entry) => entry.sectionId === sectionId);
              setActiveFieldId(section?.fields[0]?.fieldId || null);
            }}
          />
        </View>

        <ScrollView
          style={styles.fieldScroll}
          contentContainerStyle={[
            styles.fieldScrollContent,
            { paddingBottom: Math.max(insets.bottom + spacing.xl, spacing.xxl) }
          ]}
          keyboardShouldPersistTaps="handled"
        >
          {activeSection ? (
            <Text style={styles.sectionTitle}>{activeSection.label}</Text>
          ) : null}

          {(activeSection?.fields || []).map((field) => {
            const expanded = activeFieldId === field.fieldId;
            return (
              <SetupFieldCard
                key={field.fieldId}
                field={field}
                detectedFields={detectedFields}
                expanded={expanded}
                readOnly={readOnly}
                onPress={() => setActiveFieldId(expanded ? null : field.fieldId)}
                onChange={(patch) => handleFieldChange(field, patch)}
              />
            );
          })}

          <SetupValidationList
            issues={validationIssues}
            onSelectIssue={() => {
              // Validation messages are textual; user can scroll and fix manually.
            }}
          />
        </ScrollView>
      </View>

      {showPreview && pdfPath ? (
        <PreviewOverlayPanel title="PDF-Vorschau" onClose={onClosePreview}>
          <SetupPdfFieldPreview
            variant="overlay"
            emphasizeActiveHighlight
            pdfPath={pdfPath}
            detectedFields={detectedFields}
            activeFieldId={activeField?.fieldId || null}
            activeFieldLabel={activeField?.label || activeField?.fieldName || null}
            activeFieldPage={activeField?.page || 1}
          />
        </PreviewOverlayPanel>
      ) : null}
    </View>
  );
}

const styles = StyleSheet.create({
  root: {
    flex: 1,
    position: 'relative'
  },
  body: {
    flex: 1
  },
  bodyTablet: {
    flexDirection: 'row'
  },
  sideNav: {
    width: 220,
    borderRightWidth: 1,
    borderRightColor: colors.border
  },
  fieldScroll: {
    flex: 1
  },
  fieldScrollContent: {
    gap: spacing.sm,
    padding: spacing.pageX
  },
  sectionTitle: {
    ...typography.subtitle,
    color: colors.ink
  }
});

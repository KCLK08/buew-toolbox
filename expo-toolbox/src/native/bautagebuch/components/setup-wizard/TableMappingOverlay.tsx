import { useMemo, useState } from 'react';
import { Pressable, ScrollView, StyleSheet, Text, TextInput, View } from 'react-native';

import { PrimaryButton, SingleLineText } from '../../../../components/mobile';
import { colors, shadows, spacing, typography } from '../../../../constants/theme';
import {
  addWizardTableColumn,
  addWizardTableRow,
  addWizardTableSection,
  type MappingField
} from '../../lib/setup-mapping';
import { findFieldAssignedToTableCell, tableAssignmentKey } from '../../lib/setup-wizard-tables';
import type { OverlayPlacement } from '../../lib/setup-mapping';
import type { SetupWizardState, SetupWizardTableAssignment } from '../../types';

type Props = {
  wizard: SetupWizardState;
  currentField: MappingField | null;
  placement: OverlayPlacement;
  selectedAssignment: SetupWizardTableAssignment | null;
  activeTableId: string | null;
  setupModel: Record<string, unknown>;
  onChange: (next: Record<string, unknown>) => void;
  onAssignCell: (assignment: SetupWizardTableAssignment) => void;
  onSelectTable: (tableId: string) => void;
};

export function TableMappingOverlay({
  wizard,
  currentField,
  placement,
  selectedAssignment,
  activeTableId,
  setupModel,
  onChange,
  onAssignCell,
  onSelectTable
}: Props) {
  const [creatingTable, setCreatingTable] = useState(false);
  const [tableLabel, setTableLabel] = useState('');
  const [columnLabel, setColumnLabel] = useState('');
  const [columnType, setColumnType] = useState<'text' | 'checkbox'>('text');

  const activeTable = wizard.tables.find((table) => table.tableId === activeTableId) || wizard.tables[0] || null;

  const containerStyle = useMemo(() => {
    if (placement === 'top') return styles.overlayTop;
    if (placement === 'left') return styles.overlayLeft;
    if (placement === 'right') return styles.overlayRight;
    return styles.overlayBottom;
  }, [placement]);

  const createTable = () => {
    const label = tableLabel.trim();
    if (!label) return;
    const result = addWizardTableSection(setupModel, label);
    onSelectTable(result.table.tableId);
    onChange(result.setupModel);
    setTableLabel('');
    setCreatingTable(false);
  };

  const addColumn = () => {
    if (!activeTable) return;
    onChange(addWizardTableColumn(setupModel, activeTable.tableId, { label: columnLabel, type: columnType }));
    setColumnLabel('');
  };

  const addRow = () => {
    if (!activeTable) return;
    onChange(addWizardTableRow(setupModel, activeTable.tableId));
  };

  if (wizard.tables.length === 0) {
    return (
      <View style={[styles.host, styles.overlayBottom]} pointerEvents="box-none">
        <View style={[styles.panel, styles.panelEmpty, shadows.card]} pointerEvents="auto">
          <Text style={styles.emptyTitle}>Tabellen-Zuordnung</Text>
          <Text style={styles.emptyCopy}>
            Erstelle eine Tabelle mit Text- oder Checkbox-Spalten und ordne PDF-Felder Tabellenzellen zu.
          </Text>
          {creatingTable ? (
            <View style={styles.formBlock}>
              <Text style={styles.formLabel}>Tabellenname</Text>
              <TextInput
                value={tableLabel}
                onChangeText={setTableLabel}
                placeholder="z. B. Personal"
                placeholderTextColor={colors.muted}
                style={styles.input}
                autoFocus
              />
              <PrimaryButton label="Tabelle anlegen" onPress={createTable} />
            </View>
          ) : (
            <PrimaryButton label="+ Tabelle erstellen" onPress={() => setCreatingTable(true)} />
          )}
        </View>
      </View>
    );
  }

  return (
    <View style={[styles.host, containerStyle]} pointerEvents="box-none">
      <View style={[styles.panel, shadows.card]} pointerEvents="auto">
        <Text style={styles.heading}>Tabellenzelle wählen</Text>

        <ScrollView horizontal showsHorizontalScrollIndicator={false} contentContainerStyle={styles.tableTabs}>
          {wizard.tables.map((table) => {
            const active = activeTable?.tableId === table.tableId;
            return (
              <Pressable
                key={table.tableId}
                style={[styles.tab, active ? styles.tabActive : null]}
                onPress={() => onSelectTable(table.tableId)}
              >
                <SingleLineText style={active ? styles.tabLabelActive : styles.tabLabel}>{table.label}</SingleLineText>
              </Pressable>
            );
          })}
          <Pressable style={[styles.tab, styles.tabCreate]} onPress={() => setCreatingTable(true)}>
            <Text style={styles.tabCreateLabel}>+ Tabelle</Text>
          </Pressable>
        </ScrollView>

        {creatingTable ? (
          <View style={styles.formBlock}>
            <TextInput
              value={tableLabel}
              onChangeText={setTableLabel}
              placeholder="Neuer Tabellenname"
              placeholderTextColor={colors.muted}
              style={styles.input}
              autoFocus
            />
            <PrimaryButton label="Anlegen" onPress={createTable} />
          </View>
        ) : null}

        {activeTable ? (
          <>
            <View style={styles.toolbar}>
              <View style={styles.columnTypeRow}>
                <Pressable
                  style={[styles.typeChip, columnType === 'text' ? styles.typeChipActive : null]}
                  onPress={() => setColumnType('text')}
                >
                  <Text style={columnType === 'text' ? styles.typeChipLabelActive : styles.typeChipLabel}>Text</Text>
                </Pressable>
                <Pressable
                  style={[styles.typeChip, columnType === 'checkbox' ? styles.typeChipActive : null]}
                  onPress={() => setColumnType('checkbox')}
                >
                  <Text style={columnType === 'checkbox' ? styles.typeChipLabelActive : styles.typeChipLabel}>
                    Checkbox
                  </Text>
                </Pressable>
              </View>
              <TextInput
                value={columnLabel}
                onChangeText={setColumnLabel}
                placeholder={columnType === 'checkbox' ? 'Checkbox-Spalte' : 'Spaltenname'}
                placeholderTextColor={colors.muted}
                style={[styles.input, styles.inputCompact]}
              />
              <PrimaryButton label="+ Spalte" variant="secondary" onPress={addColumn} />
              <PrimaryButton label="+ Zeile" variant="ghost" onPress={addRow} />
            </View>

            {activeTable.columns.length === 0 ? (
              <Text style={styles.hint}>Füge mindestens eine Text- oder Checkbox-Spalte hinzu.</Text>
            ) : (
              <ScrollView horizontal showsHorizontalScrollIndicator={false} contentContainerStyle={styles.gridWrap}>
                <View style={styles.grid}>
                  <View style={styles.gridHeaderRow}>
                    <View style={styles.cornerCell} />
                    {activeTable.columns.map((column) => (
                      <View key={column.columnId} style={styles.headerCell}>
                        <SingleLineText style={styles.headerLabel}>{column.label}</SingleLineText>
                        <Text style={styles.headerMeta}>{column.type === 'checkbox' ? '☑' : 'Aa'}</Text>
                      </View>
                    ))}
                  </View>
                  {Array.from({ length: activeTable.rowCount }, (_, rowIndex) => (
                    <View key={`row-${rowIndex}`} style={styles.gridRow}>
                      <View style={styles.rowLabelCell}>
                        <Text style={styles.rowLabel}>Z{rowIndex + 1}</Text>
                      </View>
                      {activeTable.columns.map((column) => {
                        const assignment = {
                          tableId: activeTable.tableId,
                          rowIndex,
                          columnId: column.columnId
                        };
                        const selected =
                          selectedAssignment &&
                          tableAssignmentKey(selectedAssignment) === tableAssignmentKey(assignment);
                        const occupiedFieldId = findFieldAssignedToTableCell(wizard, assignment);
                        const occupiedByCurrent = occupiedFieldId === currentField?.fieldId;
                        return (
                          <Pressable
                            key={`${rowIndex}-${column.columnId}`}
                            style={[
                              styles.cell,
                              selected ? styles.cellSelected : null,
                              occupiedFieldId ? styles.cellOccupied : null
                            ]}
                            onPress={() => onAssignCell(assignment)}
                            disabled={!currentField}
                          >
                            <Text style={styles.cellLabel}>
                              {occupiedByCurrent ? 'Aktuell' : occupiedFieldId ? 'Belegt' : 'Zuordnen'}
                            </Text>
                          </Pressable>
                        );
                      })}
                    </View>
                  ))}
                </View>
              </ScrollView>
            )}
          </>
        ) : null}
      </View>
    </View>
  );
}

const styles = StyleSheet.create({
  host: {
    position: 'absolute',
    zIndex: 20,
    maxWidth: '100%'
  },
  overlayTop: {
    top: spacing.sm,
    left: spacing.pageX,
    right: spacing.pageX
  },
  overlayBottom: {
    bottom: spacing.sm,
    left: spacing.pageX,
    right: spacing.pageX
  },
  overlayLeft: {
    left: spacing.pageX,
    top: '18%',
    maxWidth: 280
  },
  overlayRight: {
    right: spacing.pageX,
    top: '18%',
    maxWidth: 280
  },
  panel: {
    backgroundColor: colors.panelElevated,
    borderRadius: spacing.cardRadius,
    borderWidth: 1,
    borderColor: colors.border,
    padding: spacing.md,
    gap: spacing.sm,
    maxHeight: 320
  },
  panelEmpty: {
    gap: spacing.md
  },
  emptyTitle: {
    ...typography.bodyStrong,
    color: colors.ink
  },
  emptyCopy: {
    ...typography.body,
    color: colors.muted
  },
  heading: {
    ...typography.label,
    color: colors.muted
  },
  tableTabs: {
    gap: spacing.xs
  },
  tab: {
    minHeight: 36,
    paddingHorizontal: spacing.md,
    borderRadius: 999,
    borderWidth: 1,
    borderColor: colors.border,
    backgroundColor: colors.panel,
    justifyContent: 'center'
  },
  tabActive: {
    borderColor: colors.accent,
    backgroundColor: colors.badgeBg
  },
  tabCreate: {
    borderStyle: 'dashed',
    borderColor: colors.accent
  },
  tabLabel: {
    ...typography.caption,
    color: colors.ink
  },
  tabLabelActive: {
    ...typography.caption,
    color: colors.accent2,
    fontFamily: 'SpaceGrotesk_600SemiBold'
  },
  tabCreateLabel: {
    ...typography.caption,
    color: colors.accent,
    fontFamily: 'SpaceGrotesk_600SemiBold'
  },
  formBlock: {
    gap: spacing.sm
  },
  formLabel: {
    ...typography.caption,
    color: colors.muted
  },
  toolbar: {
    gap: spacing.xs
  },
  columnTypeRow: {
    flexDirection: 'row',
    gap: spacing.xs
  },
  typeChip: {
    minHeight: 32,
    paddingHorizontal: spacing.md,
    borderRadius: 999,
    borderWidth: 1,
    borderColor: colors.border,
    justifyContent: 'center'
  },
  typeChipActive: {
    borderColor: colors.accent,
    backgroundColor: colors.badgeBg
  },
  typeChipLabel: {
    ...typography.caption,
    color: colors.ink
  },
  typeChipLabelActive: {
    ...typography.caption,
    color: colors.accent2,
    fontFamily: 'SpaceGrotesk_600SemiBold'
  },
  input: {
    ...typography.body,
    color: colors.ink,
    borderWidth: 1,
    borderColor: colors.border,
    borderRadius: spacing.inputRadius,
    paddingHorizontal: spacing.md,
    minHeight: spacing.touchMin,
    paddingVertical: spacing.sm,
    backgroundColor: colors.panel
  },
  inputCompact: {
    minHeight: 40
  },
  hint: {
    ...typography.caption,
    color: colors.muted
  },
  gridWrap: {
    paddingBottom: spacing.xxs
  },
  grid: {
    gap: spacing.xxs
  },
  gridHeaderRow: {
    flexDirection: 'row',
    gap: spacing.xxs
  },
  gridRow: {
    flexDirection: 'row',
    gap: spacing.xxs
  },
  cornerCell: {
    width: 34
  },
  headerCell: {
    width: 92,
    minHeight: 42,
    borderRadius: 10,
    backgroundColor: colors.panel,
    borderWidth: 1,
    borderColor: colors.border,
    padding: spacing.xxs,
    justifyContent: 'center',
    alignItems: 'center',
    gap: 2
  },
  headerLabel: {
    ...typography.caption,
    color: colors.ink,
    textAlign: 'center'
  },
  headerMeta: {
    ...typography.caption,
    color: colors.muted
  },
  rowLabelCell: {
    width: 34,
    justifyContent: 'center',
    alignItems: 'center'
  },
  rowLabel: {
    ...typography.caption,
    color: colors.muted
  },
  cell: {
    width: 92,
    minHeight: 44,
    borderRadius: 10,
    backgroundColor: colors.panel,
    borderWidth: 1,
    borderColor: colors.border,
    justifyContent: 'center',
    alignItems: 'center',
    paddingHorizontal: spacing.xxs
  },
  cellSelected: {
    borderColor: colors.accent,
    backgroundColor: colors.badgeBg,
    borderWidth: 2
  },
  cellOccupied: {
    borderColor: colors.accent2
  },
  cellLabel: {
    ...typography.caption,
    color: colors.ink,
    textAlign: 'center'
  }
});

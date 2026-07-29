# Technische Prüfung: Schritt 2 nach hybrider Feldarchitektur

Stand: Juli 2026  
Branch-Kontext: `cursor/setup-step2-field-management-fe37` (baut auf PR #109, #110, #111 auf)  
Art: Code-Review / Fehleranalyse — keine neuen Features

---

## Prüfkatalog

1. Persistenz manueller Feldgeometrien  
2. Wiederherstellung nach App-Neustart  
3. Gleichbehandlung AcroForm / manual  
4. Feldtypänderung ohne Änderung der `source`  
5. Löschbereinigung aller Zuordnungen  
6. Synchronisation PDF-Overlay ↔ Feldliste  

---

## 1. Persistenz manueller Feldgeometrien

### Ergebnis: **Funktioniert (DB-Pfad korrekt)**

Manuelle Felder werden über `createManualFieldInput()` mit `source: 'manual'` und vollständiger `geometry` angelegt und via `addTemplateField()` → `serializeFieldForDb()` persistiert.

Relevante Pfade:

- `src/native/bautagebuch/lib/template-field.ts` — `createManualFieldInput()`, `serializeFieldForDb()`, `normalizeDetectedField()`
- `src/native/bautagebuch/db/database.ts` — `addTemplateField()`

Gespeichert werden:

- `geometryJson` (Seite + Rechteck)
- `rectJson` (Legacy-Kompatibilität)
- `source = 'manual'`

Roundtrip-Test (Serialize → `normalizeDetectedField`): Geometrie und `source: manual` kommen unverändert zurück.

**Architektur-Hinweis:** Geometrie liegt nur in `detected_fields`, nicht im Wizard/`setupModel`. Das ist konsistent mit dem hybriden Modell.

**Randfall:** Nach dem Erstellen kann `reloadFields()` asynchron laufen — die PDF aktualisiert sich ggf. einen Moment später (kein Datenverlust).

---

## 2. Wiederherstellung nach App-Neustart

### Ergebnis: **Funktioniert für Felddaten, mit Label-Dualität**

Beim Laden von Schritt 2 (`app/bautagebuch/setup/[templateId]/assign.tsx`):

1. `getDetectedFields()` → `normalizeDetectedField()` stellt `geometry`, `source`, `rect` wieder her  
2. Migration bei DB-Boot (`migrateDetectedFieldsSchema`) füllt `geometryJson` aus Legacy-`rectJson` nach  
3. `getSetupModel()` lädt Wizard-Zustand (`assignments`, `fieldLabels`, …) separat  

**Label-Dualität (mittleres Risiko):**

| Speicherort | Inhalt |
|-------------|--------|
| `detected_fields.labelCandidate` | DB-Feldname |
| `wizard.fieldLabels` | Anzeigename im Setup-JSON |

`resolveFieldDisplayLabel()` bevorzugt `fieldLabels`. Wenn Autosave des Setup-Modells (`useSetupAutosave`, ~420 ms) nach einem Namenswechsel nicht mehr flusht, kann nach Neustart der Wizard-Name von `labelCandidate` abweichen.

**Kein separater Geometrie-Verlust** nach Neustart, solange SQLite-Daten intakt sind.

---

## 3. Gleichbehandlung AcroForm / manual

### Ergebnis: **Im Kern gleich, Darstellung bewusst unterschiedlich**

| Bereich | AcroForm | Manuell | Bewertung |
|---------|----------|---------|-----------|
| Zuordnung (`assignFieldToGroup`, …) | gleicher Code | gleicher Code | OK |
| `isMappingComplete` | gleich | gleich | OK |
| `MappingField` / `sortMappingFields` | `source` + `geometry` | `source` + `geometry` | OK |
| PDF-Overlay | nur bei `fieldHasGeometry()` | nur bei `fieldHasGeometry()` | OK |
| Overlay-Farbe | grün (`source-acroform`) | amber (`source-manual`) | beabsichtigt |
| Felder ohne Geometrie | in Liste, kein Overlay | in Liste, kein Overlay | OK |

**Keine AcroForm-only-Pfade** mehr in Schritt 2 — Highlights laufen über `fieldToPreviewLegacyRect()` / `geometry`.

**Asymmetrie beim Import:** AcroForm über Scan-Pipeline, manual über `addTemplateField()` — konvergieren auf dasselbe DB-Schema.

---

## 4. Feldtypänderung ohne Änderung der `source`

### Ergebnis: **DB-Logik korrekt, UI-Race möglich**

`updateTemplateField()` schließt `source` vom Patch aus und überschreibt die DB-Spalte `source` nicht:

```typescript
// database.ts — merged.source = existing.source
// UPDATE setzt nur fieldName, labelCandidate, type, … — nicht source
```

**Bug (niedrig): Race bei Typänderung**

In `SetupAssignStep.tsx`:

```typescript
const handleFieldTypeChange = (fieldId, type) => {
  void onUpdateField?.(fieldId, { type });
  onFieldsChanged?.();
};
```

`onFieldsChanged()` (Reload) läuft parallel zu `await updateTemplateField()` — kurz kann der alte Typ in der Liste stehen. `source` bleibt davon unberührt.

**Namensänderung:** Jeder Tastendruck schreibt sofort in DB (`labelCandidate`) und Wizard (`fieldLabels`) — funktional ok, aber ohne Debounce und mit Dualitäts-Risiko (s. Abschnitt 2).

---

## 5. Löschbereinigung aller Zuordnungen

### Ergebnis: **Wizard bereinigt, Setup-Sections nicht**

**Was beim Löschen passiert:**

| Aktion | Status |
|--------|--------|
| `deleteTemplateField()` — DB-Zeile | OK |
| `removeFieldFromWizard()` — `assignments`, `tableAssignments`, `fieldLabels`, `deferredFieldIds` | OK |
| `currentFieldIndex` anpassen | OK |
| Autosave Setup-Modell via `onChange` | OK |
| `single_sections` / `table_sections` bereinigen | **Fehlt** |

`removeFieldFromWizard()` in `setup-mapping.ts` touchiert `single_sections` / `table_sections` nicht.

**Auswirkung:**

- **Erst-Setup (draft):** meist unkritisch — `single_sections` existieren erst nach „Weiter zu Schritt 3“ (`rebuildSectionsFromWizard`).
- **Re-Edit fertiger Vorlagen:** Gelöschte Felder können in `single_sections`/`table_sections` als Geister-Einträge bleiben, bis Schritt 2 erneut abgeschlossen wird.
- **Re-Edit + Sprung zu Schritt 3** (Step-Nav ohne `rebuildSectionsFromWizard`): Schritt 3 kann über `listFieldSettingsTargets()` noch gelöschte `fieldId`s aus `single_sections` anzeigen.

**Schweregrad:** Mittel — betrifft vor allem Re-Edit, nicht den Erst-Durchlauf bis Schritt 3.

---

## 6. Synchronisation PDF-Overlay ↔ Feldliste

### Ergebnis: **Liste → PDF größtenteils OK, mit Lücken**

**Funktioniert:**

- Auswahl in der Feldliste setzt `wizard.currentFieldIndex` → `activeFieldId` / `activeFieldPage` an `SetupPdfFieldPreview`
- WebView erhält `setActive` → `refreshAllOverlays()` + `scrollToActiveField()` (scrollbare Assign-Ansicht)
- Overlays basieren auf `field.geometry`; Labels aus `fieldLabels` + `labelCandidate`
- Quellen-Farben (grün/amber) werden gesetzt

**Probleme:**

| Problem | Schweregrad | Details |
|---------|-------------|---------|
| Feldnummern divergieren | Mittel | Liste: Index über alle `mappingFields`. PDF: Index nur über Felder mit Geometrie. Bei Feldern ohne Rect unterschiedliche Nummern. |
| Kein PDF → Liste Tap | Niedrig | Overlays haben `pointer-events: none` — keine Rückwärts-Sync per Tap auf Markierung. |
| Voller WebView-Reload bei Label-Änderung | Niedrig (Performance) | `highlights` in `useEffect`-Deps → HTML-Neuladen pro Tastendruck. Sync ok, teuer. |
| Felder ohne Geometrie | Erwartbar | Kein Overlay, kein Scroll-Ziel — Liste zeigt Feld, PDF nicht. |
| Tab „Felder“ aktiv | OK | PDF-Tab hidden, aber gemountet — `postPreviewCommand` läuft weiter. |

**Scroll-Sync:** `setActive` im Scroll-Modus scrollt per `.highlight.active` zur richtigen Seite — für Felder mit Geometrie ausreichend.

---

## Gesamtbewertung

| Prüfpunkt | Status | Kurzfassung |
|-----------|--------|-------------|
| Persistenz manueller Geometrien | OK | DB-Roundtrip sauber |
| Wiederherstellung nach Neustart | OK mit Vorbehalt | Geometrie ja; Label-Dualität möglich |
| Gleichbehandlung AcroForm/manual | OK | Logisch gleich, visuell differenziert |
| Typänderung ohne source-Wechsel | OK | DB schützt `source`; UI-Race bei Reload |
| Löschbereinigung | Teilweise | Wizard ja, `single_sections` bei Re-Edit nein |
| PDF ↔ Feldliste Sync | Teilweise | Liste→PDF ok; Nummern, kein Rückweg-Tap |

---

## Priorisierte Fehlerliste

1. **Mittel:** `removeFieldFromWizard()` bereinigt keine `single_sections`/`table_sections` — Geisterfelder bei Re-Edit / direktem Sprung zu Schritt 3.  
2. **Mittel:** Feldnummern PDF vs. Liste bei gemischten Feldern mit/ohne Geometrie inkonsistent.  
3. **Niedrig:** Race bei Typänderung (`onFieldsChanged` vor Ende von `updateTemplateField`).  
4. **Niedrig:** Label in DB vs. Wizard kann bei abgebrochenem Autosave divergieren.  
5. **Niedrig:** Keine PDF→Liste-Interaktion (nur eine Richtung).

---

## Fazit

Für den **Erst-Durchlauf** (Import → Schritt 2 → Schritt 3) ist die Architektur **weitgehend stabil**. Die größten Risiken liegen im **Re-Edit-Pfad** (Löschen ohne Rebuild) und bei der **Overlay-Nummerierung** bei unvollständigen Geometrien.

---

## Relevante Dateien

| Bereich | Pfad |
|---------|------|
| Schritt 2 UI | `src/native/bautagebuch/components/setup-wizard/SetupAssignStep.tsx` |
| Feldliste | `SetupAssignFieldListPanel.tsx` |
| PDF-Preview | `src/native/bautagebuch/components/SetupPdfFieldPreview.tsx` |
| Overlay-HTML | `src/native/bautagebuch/lib/pdf-preview-html.ts` |
| Feldmodell | `src/native/bautagebuch/lib/template-field.ts` |
| Wizard/Mapping | `src/native/bautagebuch/lib/setup-mapping.ts` |
| DB | `src/native/bautagebuch/db/database.ts` |
| Screen | `app/bautagebuch/setup/[templateId]/assign.tsx` |
| Autosave | `src/native/bautagebuch/hooks/useSetupAutosave.ts` |
| Schritt 3 Targets | `src/native/bautagebuch/lib/setup-field-settings.ts` |

# STEP 0 – Analyse des PDF-Vorlagen-Setups (BÜW Toolbox)

**Stand:** Analyse ohne Codeänderungen  
**Scope:** `expo-toolbox/` (Expo/React-Native-App, Bautagebuch-Modul)  
**Ziel der späteren Erweiterung:** 2-stufiger Workflow — (1) geführter Feld-Zuordnungsmodus mit PDF-Overlay, (2) bestehender Gruppen-/Feld-Editor — ohne den bestehenden Editor zu beschädigen oder zu ersetzen.

---

## 1. Aktueller Workflow

### Überblick

Es existieren **zwei parallele Setup-Pfade**, die über Routing und Template-Typ unterschieden werden:

| Pfad | Auslöser | Schritt 1 (Mapping) | Schritt 2 (Editor) |
|------|----------|---------------------|---------------------|
| **Neuer Wizard** | User-importierte PDF (`templateKind === ''`) | `SetupMappingStep` | `SetupFieldSettingsStep` |
| **Legacy-Editor** | `builtin-etb` oder `table_sections.length > 0` | entfällt | `SetupEditor` |

### End-to-End-Flow (User-importierte PDF)

```
/bautagebuch/config (Setup-Tab)
  │
  ├─► PDF importieren (TemplateOverviewList → templateService.importTemplateFromDocument)
  │     ├─ Document Picker (expo-document-picker)
  │     ├─ Scan (pdf-scan → Fallback pdf-scan-lite)
  │     ├─ PDF speichern: {documentDirectory}bautagebuch/templates/{templateId}_{fileName}
  │     ├─ SQLite: templates (status draft), detected_fields, setup_models (status in_progress)
  │     └─ setupModel mit wizard.step = 'mapping'
  │
  └─► resolveSetupEntryPath() → /bautagebuch/setup/[templateId]/mapping  (Schritt 1)
        │
        ├─ SetupMappingStep: Feld-für-Feld Gruppenzuordnung auf PDF-Overlay
        │     ├─ Zuweisung → wizard.assignments[fieldId] = sectionId
        │     ├─ Überspringen → wizard.deferredFieldIds
        │     └─ Abschluss → rebuildSectionsFromWizard() → wizard.step = 'fields'
        │
        └─► /bautagebuch/setup/[templateId]/fields  (Schritt 2)
              ├─ SetupFieldSettingsStep: Feldeigenschaften pro Gruppe
              ├─ Autosave (420 ms Debounce)
              └─ „Setup abschließen“ → status 'ready' → zurück zu /bautagebuch/config
```

### Built-in ETB-Vorlage

- Wird beim App-Start via `ensureBuiltinTemplate()` bereitgestellt.
- PDF-Download von `{toolboxWebBaseUrl}/bautagebuch/templates/Vorlage-eBTB.pdf`.
- Fertiges Setup via `buildEtbSetupModel()` — **Status sofort `ready`**, kein Wizard.
- Enthält `table_sections` (Personal- und Leistungstabellen) → Legacy-Editor.

### Template-Status-Lebenszyklus

| Status | Bedeutung |
|--------|-----------|
| `draft` | Frisch importiert |
| `in_progress` | Setup begonnen, noch nicht abgeschlossen |
| `ready` | Setup abgeschlossen, kann aktiviert werden |
| `archived` | Archiviert, nicht mehr editierbar (Wizard: readOnly) |

---

## 2. Relevante Dateien

### Routen & Screens

| Datei | Rolle |
|-------|-------|
| `expo-toolbox/app/bautagebuch/(tabs)/config.tsx` | Setup-Hub: Import, Liste, Aktivieren, Archivieren, Navigation |
| `expo-toolbox/app/bautagebuch/setup/index.tsx` | Redirect → `/bautagebuch/config` |
| `expo-toolbox/app/bautagebuch/setup/_layout.tsx` | Stack: `mapping`, `fields` |
| `expo-toolbox/app/bautagebuch/setup/[templateId]/mapping.tsx` | Screen Schritt 1 |
| `expo-toolbox/app/bautagebuch/setup/[templateId]/fields.tsx` | Screen Schritt 2 / Legacy-Router |

### Services & Persistenz

| Datei | Rolle |
|-------|-------|
| `expo-toolbox/src/native/bautagebuch/services/templateService.ts` | Import, ETB-Bootstrap, Bundle-Laden, Aktivierung |
| `expo-toolbox/src/native/bautagebuch/db/database.ts` | SQLite CRUD: templates, detected_fields, setup_models |
| `expo-toolbox/src/native/bautagebuch/hooks/useSetupAutosave.ts` | Debounced Autosave (420 ms) |

### Scan & Modelle

| Datei | Rolle |
|-------|-------|
| `expo-toolbox/src/native/bautagebuch/lib/pdf-scan.web.ts` | Vollscan (Web): pdf-lib + pdfjs-dist |
| `expo-toolbox/src/native/bautagebuch/lib/pdf-scan-lite.ts` | Lite-Scan: pdf-lib only, rect=null |
| `expo-toolbox/src/native/bautagebuch/lib/pdf-scan.ts` | Native: re-exportiert Lite |
| `expo-toolbox/src/native/bautagebuch/lib/scan-meta.ts` | Rescan-Version & -Trigger |
| `expo-toolbox/src/native/bautagebuch/lib/setup-mapping.ts` | Wizard-Logik, Gruppen, Section-Rebuild |
| `expo-toolbox/src/native/bautagebuch/lib/generic-setup-model.js` | Initiales generisches Modell (Seiten-Gruppen) |
| `expo-toolbox/src/native/bautagebuch/lib/etb-template.js` | ETB-Festlayout, Upgrade, Tabellen |
| `expo-toolbox/src/native/bautagebuch/lib/setup-model.js` | Validierung, Run-Sections, PDF-Export, Feldtypen |
| `expo-toolbox/src/native/bautagebuch/lib/setup-field-hints.js` | Checkbox-Default-Helfer |
| `expo-toolbox/src/native/bautagebuch/types.ts` | TypeScript-Typen |

### UI-Komponenten

| Datei | Rolle |
|-------|-------|
| `components/setup-wizard/TemplateOverviewList.tsx` | Vorlagenliste + Import-Button |
| `components/setup-wizard/SetupMappingStep.tsx` | Schritt 1: Feld-Zuordnung |
| `components/setup-wizard/GroupOverlayCards.tsx` | Gruppenauswahl als Overlay auf PDF |
| `components/setup-wizard/SetupProgressHeader.tsx` | Fortschrittsanzeige Mapping |
| `components/setup-wizard/SetupFieldSettingsStep.tsx` | Schritt 2: Feldeinstellungen (Wizard) |
| `components/setup-wizard/SetupFieldCard.tsx` | Einzelnes Feld (Wizard) |
| `components/setup-wizard/SetupGroupNav.tsx` | Gruppen-Navigation |
| `components/setup-wizard/SetupValidationList.tsx` | Validierungsfehler |
| `components/SetupEditor.tsx` | Legacy-Editor (Gruppen + Tabellen) |
| `components/SetupTemplateManager.tsx` | Legacy Template-Picker (nur non-embedded) |
| `components/SetupPdfFieldPreview.tsx` | PDF-Vorschau mit Feld-Overlays |
| `components/PdfPreviewPanel.tsx` | Fallback-Vorschau ohne Overlays |
| `components/PreviewOverlayPanel.tsx` | Vollbild-Overlay-Wrapper |

---

## 3. Komponentenübersicht

```
config.tsx
└── TemplateOverviewList
      └── templateService.importTemplateFromDocument()
            └── resolveSetupEntryPath() → mapping | fields

mapping.tsx (Schritt 1)
└── SetupMappingStep
      ├── SetupProgressHeader
      ├── SetupPdfFieldPreview (variant="mapping")
      └── GroupOverlayCards

fields.tsx (Schritt 2)
├── [Wizard] SetupFieldSettingsStep
│     ├── SetupGroupNav
│     ├── SetupFieldCard (pro Feld)
│     ├── SetupValidationList
│     ├── SetupPdfFieldPreview (overlay)
│     └── PreviewOverlayPanel
└── [Legacy] SetupEditor
      ├── SetupTemplateManager (nur wenn !embedded)
      ├── SetupPdfFieldPreview (overlay)
      └── PreviewOverlayPanel
```

### Routing-Entscheidung (`resolveSetupEntryPath`)

```typescript
if (templateKind === 'builtin-etb' || hasTableSections(setupModel))
  → /setup/[id]/fields        // Legacy, kein Mapping
if (wizard.step === 'fields')
  → /setup/[id]/fields
else
  → /setup/[id]/mapping
```

### Editor-Auswahl in `fields.tsx`

```typescript
useLegacyEditor = templateKind === 'builtin-etb' || hasTableSections(setupModel)
```

Guard in `fields.tsx`: Wizard-Pfad redirectet zu Mapping, wenn `wizard.step !== 'fields'`.

---

## 4. Datenmodelle

### Template (`BautagebuchTemplate`)

```typescript
{
  templateId: string;        // z.B. "tplv2_1720000000000"
  templateName: string;
  fileName: string;
  templateKind: string;      // "" | "builtin-etb"
  mimeType: string;          // "application/pdf"
  sizeBytes: number;
  pageCount: number;
  pdfPath: string;           // absoluter Pfad im Document Directory
  status: 'draft' | 'in_progress' | 'ready' | 'archived';
  createdAt: string;
  updatedAt: string;
  deleted_at: string | null;
}
```

### Erkanntes Feld (`DetectedField`) — Rohdaten aus PDF-Scan

```typescript
{
  id: string;                // "{templateId}::{fieldId}"
  templateId: string;
  fieldId: string;           // z.B. "text1-p1-o3"
  fieldName: string;         // AcroForm-Name aus PDF
  labelCandidate: string;    // inferiert oder humanisiert
  type: string;              // text | checkbox | radio | dropdown | unsupported
  options: string[];         // für dropdown/radio
  page: number;
  orderIndex: number;
  rect: number[] | null;     // [x1, y1, x2, y2] PDF-Koordinaten
  createdAt: string;
  updatedAt: string;
}
```

**Speicherung:** SQLite-Tabelle `detected_fields`, Spalte `rectJson` als JSON-Array.

### Setup-Modell (Top-Level)

```typescript
{
  modelId: string;
  version: number;             // Generic: 1, ETB: 8
  status: string;
  templateId: string;
  templateName: string;
  pageCount: number;
  single_sections: SingleSection[];
  table_sections: TableSection[];   // nur ETB/Legacy
  section_order: { kind: 'single'|'table', id: string }[];
  wizard?: SetupWizardState;        // nur Generic-Wizard
  createdAt: string;
  updatedAt: string;
}
```

### Single Section (Gruppe)

```typescript
{
  sectionId: string;         // z.B. "kopfdaten" oder "page_1"
  label: string;
  page?: number;
  fields: SetupFieldConfig[];
}
```

### Feld-Konfiguration (`SetupFieldConfig`)

```typescript
{
  fieldId: string;
  fieldName?: string;
  label?: string;
  type?: string;
  options?: string[];
  required?: boolean;
  skipped?: boolean;
  multiline?: boolean;
  defaultValue?: string;
  hint?: string;             // nur Wizard-Editor
  page?: number;
  rect?: number[] | null;
}
```

### Wizard-State (`SetupWizardState`) — transient bis Rebuild

```typescript
{
  step: 'mapping' | 'fields';
  currentFieldIndex: number;
  groups: { sectionId: string; label: string }[];
  assignments: Record<fieldId, sectionId>;
  deferredFieldIds: string[];
}
```

**Persistiert in:** `setupModelJson.wizard` (SQLite `setup_models`).

### Tabellen-Sektion (nur Legacy/ETB)

```typescript
{
  tableId: string;
  label: string;
  page?: number;
  source?: 'etb-fixed' | 'etb-code-lock';
  columns: {
    columnId: string;
    label: string;
    required?: boolean;
    skipped?: boolean;
    multiline?: boolean;
  }[];
  rows: {
    rowId: string;
    index?: number;
    skipped?: boolean;
    cells: {
      cellId?: string;
      tableId?: string;
      rowId?: string;
      columnId: string;
      fieldId?: string;
      fieldName?: string;
      page?: number;
      rect?: number[];
      required?: boolean;
      skipped?: boolean;
      multiline?: boolean;
    }[];
  }[];
}
```

---

## 5. Bestehender Editor

### Wizard-Editor (`SetupFieldSettingsStep` + `SetupFieldCard`)

**Navigation:**
- Horizontale Gruppen-Chips (Mobil) oder Sidebar (Tablet ≥ 900 px).
- Felder als expandierbare Karten pro Gruppe.
- Optional: PDF-Vorschau-Overlay (Toggle im Screen-Footer).

**Bearbeitbare Feldeigenschaften:**

| Parameter | UI-Label | Typ | Beschreibung |
|-----------|----------|-----|--------------|
| `label` | Anzeigename im Assistenten | string | Anzeigename im BTB-Run-Wizard |
| `defaultValue` | Standardtext / Checkbox-Default | string | Vorausfüllung; Checkbox: `'true'`/`'false'` |
| `required` | Pflichtfeld | boolean | Export-Validierung |
| `skipped` | Im Assistenten ausblenden | boolean | Versteckt im Run, bleibt im PDF |
| `multiline` | Mehrzeiliges Eingabefeld | boolean | Nur Nicht-Checkbox |
| `hint` | Hilfetext | string | **Nur Wizard** — noch nicht im RunWizard konsumiert |

**Read-only Meta:** `fieldName`, `fieldId`, `page`, `type`.

**Nicht verfügbar im Wizard:** Feld-Reihenfolge ändern, Felder zwischen Gruppen verschieben, Gruppen umbenennen, Section-Reihenfolge.

### Legacy-Editor (`SetupEditor`)

**Navigation:**
- Modus-Umschaltung: **Gruppen** | **Tabellen**.
- Gruppen-/Tabellen-Auswahl in Sidebar.
- Collapsible **„Reihenfolge im Bautagebuch"** (`section_order`).
- Feld-Auswahl per Karte; Vorschau-Highlight synchronisiert.

**Single-Field Bearbeitung:**

| Parameter | UI | Legacy |
|-----------|-----|--------|
| `label` | ✓ | ✓ |
| `defaultValue` | ✓ | ✓ (Checkbox via setup-field-hints) |
| `required` | ✓ | ✓ |
| `skipped` | ✓ | ✓ |
| `multiline` | ✓ | ✓ |
| `hint` | ✓ | ✗ |

**Zusätzliche Legacy-Funktionen (nicht im Wizard):**
- Gruppenname bearbeiten.
- Felder innerhalb Gruppe ↑/↓ verschieben.
- Feld in andere Gruppe verschieben (Alert-Auswahl).
- Tabellen: Spaltenlabel, required/skipped/multiline pro Spalte.
- Tabellen: Zell-Feld-Zuordnung (fest bei ETB).

**Embedded-Modus** (Setup-Routen): `SetupTemplateManager` ausgeblendet, kein Import.

> **Wichtig für spätere Erweiterung:** Der Legacy-Editor (`SetupEditor`) und der Wizard-Editor (`SetupFieldSettingsStep`) sind getrennte Komponenten. Änderungen am Mapping-Wizard dürfen `SetupEditor` nicht verändern. ETB-Pfad muss unverändert bleiben.

---

## 6. PDF Rendering

### Vorschau-Komponente: `SetupPdfFieldPreview`

| Aspekt | Detail |
|--------|--------|
| **Host** | `react-native-webview` |
| **Renderer** | pdf.js **3.11.174** via CDN (`cdnjs.cloudflare.com`) |
| **PDF-Daten** | Base64 eingebettet in HTML (kein postMessage für PDF-Bytes) |
| **Overlays** | Absolute `<div>`-Elemente über Canvas, Koordinaten aus `rect` |
| **Y-Achse** | PDF-Koordinaten → Canvas: Y gespiegelt |
| **Varianten** | `default`, `pinned`, `mapping`, `overlay` |
| **Mapping-Modus** | Höhere Canvas-Skalierung, aktives Feld pulsiert, andere abgedunkelt |
| **Seitenwechsel** | WebView `postMessage`: `{ type: 'setPage', page }`, `{ type: 'setActive', fieldId, page }` |
| **Fallback** | Bei Fehler → `PdfPreviewPanel` (ohne Overlays) |

### Seiten laden

1. PDF von `pdfPath` als Base64 gelesen (`expo-file-system/legacy`).
2. pdf.js `getDocument({ data: bytes })` im WebView.
3. `renderPage(n)` mit viewport-basierter Skalierung (max CSS-Breite 860 px, DPR bis 3).
4. Overlays pro Seite gefiltert nach `field.page`.

### Bibliotheken (Scan vs. Preview)

| Kontext | Bibliothek | Version |
|---------|------------|---------|
| Scan (Web) | pdf-lib + pdfjs-dist | pdf-lib ^1.17.1, pdfjs-dist ^4.8.69 |
| Scan (Native) | pdf-lib (Lite) | — |
| Preview (WebView) | pdf.js CDN | 3.11.174 |

---

## 7. Felderkennung

### Ablauf

1. **pdf-lib** `PDFDocument.load()` → `getForm().getFields()`.
2. **Typ-Erkennung** via `detectPdfFieldType()` in `setup-model.js`:
   - `PDFTextField` → `text`
   - `PDFCheckBox` → `checkbox`
   - `PDFRadioGroup` → `radio`
   - `PDFDropdown` / `PDFOptionList` → `dropdown`
   - Sonst → `unsupported` (wird aus Mapping gefiltert)
3. **Web-Vollscan** (`pdf-scan.web.ts`):
   - pdfjs `getAnnotations({ intent: 'display' })` für Widget-Rects und Seitennummer.
   - `getTextContent()` für Label-Inferenz (`inferLabelCandidate()` — links/oben vom Feld).
4. **Lite-Scan** (`pdf-scan-lite.ts`, Native-Fallback):
   - Nur AcroForm-Metadaten, `page: 1`, `rect: null`.
5. **Sortierung:** page → rect-Position → orderIndex → fieldName.
6. **ID-Vergabe:** `{slugify(fieldName)}-p{page}-o{orderIndex}` (dedupliziert).

### Rescan-Trigger (`detectedFieldsNeedRescan`)

- Keine Felder vorhanden.
- Beliebiges Feld ohne `rect`.
- Dropdown/Radio ohne `options`.
- Version `ETB_SCAN_VERSION = 3`.

### Mapping PDF ↔ Felder

| Ebene | Verknüpfung |
|-------|-------------|
| Scan → DB | `detected_fields` pro `templateId` |
| Scan → Setup | `fieldId` als stabiler Schlüssel |
| Setup-Feld | Kopiert `fieldId`, `fieldName`, `page`, `rect`, `type` aus Scan |
| Preview | `detectedFields[].rect` + `fieldId` für Highlight |
| PDF-Export | `fieldName` → AcroForm-Feld in pdf-lib |

---

## 8. Gruppenlogik

### Default-Gruppen (Wizard)

| sectionId | label |
|-----------|-------|
| `kopfdaten` | Kopfdaten |
| `witterung` | Witterung |
| `baustellenbesetzung` | Baustellenbesetzung |
| `leistungsblock` | Leistungsblock |
| `abschluss` | Abschluss |
| `sonstiges` | Sonstiges |

### Gruppen-Lebenszyklus

1. **Import:** `buildGenericSetupModel()` erzeugt vorläufige Gruppen **pro PDF-Seite** (`page_1`, `page_2`, …).
2. **Wizard Init:** `ensureWizardInitialized()` setzt Default-Gruppen, `step: 'mapping'`.
3. **Mapping:** User ordnet Felder zu via `assignFieldToGroup()` oder überspringt via `deferField()`.
4. **Neue Gruppe:** `addWizardGroup()` → generiert `sectionId` via `createId('group')`.
5. **Rebuild:** `rebuildSectionsFromWizard()`:
   - Erzeugt `single_sections` aus `assignments`.
   - Deferred Felder → Fallback-Gruppe `sonstiges`.
   - Setzt `wizard.step = 'fields'`.
   - Leere Gruppen werden herausgefiltert.

### Overlay-Platzierung (`resolveOverlayPlacement`)

Basierend auf Feld-Rect-Mittelpunkt:
- `centerY > 700` → Overlay oben
- `centerY < 220` → Overlay unten
- `centerX < 220` → Overlay rechts
- `centerX > 520` → Overlay links
- Sonst → unten

### Gruppen im Legacy-Editor

- Gruppen kommen aus `single_sections` (ETB: fest codiert in `etb-template.js`).
- Umbenennen, Feld-Reihenfolge, Verschieben zwischen Gruppen direkt im Editor.
- `section_order` steuert BTB-Run-Reihenfolge (Single + Table gemischt).

---

## 9. Empfohlene Erweiterungspunkte

Für den geplanten **2-stufigen Workflow** (Schritt 1: geführtes Mapping mit Overlay, Schritt 2: bestehender Editor):

### Bereits vorhanden — wiederverwenden

| Komponente / Modul | Grund |
|--------------------|-------|
| `SetupMappingStep` | Schritt 1 existiert bereits als Wizard |
| `SetupPdfFieldPreview` (variant `mapping`) | PDF-Overlay mit Feld-Highlight |
| `GroupOverlayCards` | Gruppenauswahl auf PDF |
| `setup-mapping.ts` | Wizard-State, Assignment, Rebuild |
| `SetupFieldSettingsStep` | Schritt 2 für Generic-PDFs |
| `SetupEditor` | Schritt 2 für ETB/Tabellen — **unverändert lassen** |
| `useSetupAutosave` | Persistenz während Bearbeitung |
| `resolveSetupEntryPath` | Einstiegspunkt für Routing-Erweiterung |

### Mögliche Erweiterungen (neu oder Anpassung)

| Bereich | Vorschlag |
|---------|-----------|
| **Schritt 1 UX** | Geführter Modus kann `SetupMappingStep` erweitern oder ersetzen — State-Modell (`SetupWizardState`) bleibt kompatibel |
| **Native Scan** | Vollscan auf iOS/Android aktivieren (aktuell Lite-only) für zuverlässige Overlays |
| **Routing** | Optionaler dritter Zwischenschritt über `wizard.step` erweitern (z.B. `'intro'`) |
| **Gruppen-Editor Schritt 2** | Wizard-Editor um Gruppen-Umbenennung / Feld-Verschieben ergänzen — separat vom Legacy-Editor |
| **`hint`-Feld** | In `RunWizard` anzeigen, wenn Schritt-2-Erweiterung gewünscht |

### Integrationspunkt für neuen Assistenten

```
mapping.tsx
  └── [NEUER Assistent ODER erweiterter SetupMappingStep]
        ├── SetupPdfFieldPreview (wiederverwendet)
        ├── GroupOverlayCards (wiederverwendet)
        └── setup-mapping.ts API (assignFieldToGroup, deferField, rebuildSectionsFromWizard)
              │
              ▼
fields.tsx → SetupFieldSettingsStep (unverändert) | SetupEditor (unverändert)
```

**Empfehlung:** Neuen Assistenten als **Alternative/Erweiterung von Schritt 1** einhängen, nicht Schritt 2 ersetzen. `rebuildSectionsFromWizard()` als stabile Schnittstelle zwischen Schritt 1 und 2 beibehalten.

---

## 10. Risiken

| Risiko | Schwere | Detail |
|--------|---------|--------|
| **Native Scan ohne Rects** | Hoch | `pdf-scan.ts` exportiert nur Lite → Mapping-Overlays auf iOS/Android ohne Positionsdaten |
| **CDN-Abhängigkeit Preview** | Mittel | pdf.js von cdnjs — offline/Netzwerkfehler → Fallback ohne Overlays |
| **pdf.js Versions-Mismatch** | Mittel | Preview 3.11.174 vs. Scan 4.8.69 — unterschiedliche Codepfade |
| **Zwei Editor-Implementierungen** | Mittel | Wizard vs. Legacy — Feature-Parität schwer zu halten |
| **`hint` ungenutzt** | Niedrig | Im Wizard editierbar, `RunWizard` zeigt es nicht |
| **Deferred → sonstiges** | Niedrig | Übersprungene Felder landen still in `sonstiges` |
| **Zero-Field PDFs** | Niedrig | Mapping wird übersprungen, leere Sections |
| **40 MB Import-Limit** | Niedrig | Harte Grenze in `templateService` |
| **AcroForm-only** | Mittel | PDFs ohne interaktive Felder werden abgelehnt |
| **ETB hard-coded** | Hoch (ETB) | Feste Feldnamen in `etb-template.js` — PDF-Strukturänderungen brechen Mapping |
| **Autosave-Race** | Niedrig | 420 ms Debounce; `flush()` bei Navigation kritisch |
| **Legacy readOnly-Lücke** | Niedrig | Archivierte Templates: Wizard readOnly, Legacy-Switches nicht überall disabled |
| **Editor-Beschädigung** | Hoch | Jede Änderung an `SetupEditor` oder ETB-Pfad kann Produktions-Setup brechen |

### Kritische Schutzmaßnahmen für die Erweiterung

1. **`SetupEditor.tsx` nicht anfassen** — ETB und Tabellen-Setup müssen unverändert funktionieren.
2. **Routing-Guards beibehalten** — `useLegacyEditor`-Logik und `resolveSetupEntryPath` nicht vereinfachen.
3. **Wizard-State backward-compatible** — `SetupWizardState` erweitern statt ersetzen.
4. **`rebuildSectionsFromWizard` als Contract** — Schritt 1 muss valide `single_sections` erzeugen, die Schritt 2 konsumiert.
5. **Native Rect-Scan priorisieren** — ohne Rects ist PDF-Overlay-Assistent auf Mobilgeräten eingeschränkt.

---

## Anhang: Persistierte Daten während Bearbeitung

| Daten | Wo | Wann geschrieben |
|-------|-----|------------------|
| PDF-Datei | Filesystem `bautagebuch/templates/` | Import |
| `detected_fields` | SQLite | Import, Rescan |
| `setup_models.setupModelJson` | SQLite | Autosave, Mapping-Complete, Finish |
| `wizard.*` | Innerhalb setupModelJson | Mapping-Schritte |
| `single_sections` | Innerhalb setupModelJson | Nach `rebuildSectionsFromWizard` |
| `templates.status` | SQLite | Parallel zu setup_models.status |

**Datenüberleben bei Navigation:** Alles in SQLite + Filesystem. Screen-State (aktive Gruppe, expandiertes Feld, Preview-Toggle) ist **nur im React-State** und geht bei Screen-Wechsel verloren — Setup-Daten bleiben erhalten.

---

## Anhang: Verwendete NPM-Pakete

| Paket | Verwendung |
|-------|------------|
| `pdf-lib` | AcroForm-Scan, PDF-Export/Befüllung |
| `pdfjs-dist` | Vollscan (Web), Worker in `assets/pdf.worker.min.mjs` |
| `react-native-webview` | PDF-Vorschau |
| `expo-document-picker` | PDF-Import |
| `expo-file-system` | PDF-Speicherung, Base64-Lesen |
| `expo-sqlite` | Template-Datenbank |

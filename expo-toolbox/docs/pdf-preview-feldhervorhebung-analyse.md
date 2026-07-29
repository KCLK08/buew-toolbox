# Technische Analyse: PDF-Feldhervorhebung

**Stand:** 29.07.2026  
**Kontext:** Erkannte PDF-Felder werden in der Setup-Vorschau (Schritt 2/3) nicht visuell hervorgehoben.

---

## Architektur-Überblick

```
PDF Import
  └─► scanTemplateBytes()
        └─► scanTemplatePdf (Native) = scanTemplatePdfLite
              └─► rect: null, page: 1
                    └─► SQLite detected_fields.rectJson = null
                          └─► SetupPdfFieldPreview: highlights = []
                                └─► WebView drawOverlay() zeichnet nichts

Setup Schritt 2/3
  └─► SetupAssignStep / SetupFieldsStep
        └─► SetupPdfFieldPreview (variant="assign")
              └─► buildFieldPreviewHtml (Einzelseite + #overlay)
                    └─► setActive via postMessage (funktioniert, aber ohne rect-Daten wirkungslos)

Setup Schritt 1
  └─► SetupStructureStep
        └─► PdfPreviewPanel (buildScrollablePdfPreviewHtml, ohne Overlays)
```

---

## 1. Wo werden erkannte PDF-Felder gespeichert?

| Speicherort | Inhalt | Zugriff |
|---|---|---|
| SQLite `detected_fields` | Primäre Quelle nach Import | `saveDetectedFields()` / `getDetectedFields()` in `database.ts` |
| Setup-Modell `single_sections[].fields[]` | Kopie nach Zuordnung (Schritt 2) | `fieldConfigFromDetected()` in `setup-mapping.ts` |
| Laufzeit `MappingField` | Abgeleitet aus `detected_fields` | `sortMappingFields()` |

**Schema** (`detected_fields`):

- `fieldId`, `fieldName`, `labelCandidate`, `type`, `optionsJson`
- `page`, `orderIndex`
- `rectJson` (JSON-Array mit 4 PDF-Koordinaten)

**Relevante Dateien:**

- `expo-toolbox/src/native/bautagebuch/db/database.ts`
- `expo-toolbox/src/native/bautagebuch/types.ts` (`DetectedField`)
- `expo-toolbox/src/native/bautagebuch/services/templateService.ts` (Import/Rescan)

---

## 2. Wo werden Positionsdaten gespeichert?

Es gibt **kein Feld `pdfPosition`** im Codebase.

Position = **`rect`** (PDF-Punkte, Ursprung unten links: `[x1, y1, x2, y2]`) + **`page`**.

| Pipeline | `rect` | `page` |
|---|---|---|
| `pdf-scan-full.ts` (pdf.js Widget-Annotations) | Aus `annotation.rect` | Korrekte Seitenzahl |
| `pdf-scan-lite.ts` (nur pdf-lib AcroForm) | Immer `null` | Immer `1` |

### Native-Import (Hauptproblem)

`pdf-scan.ts` (Native-Einstieg):

```typescript
/**
 * Native scan: pdf-lib AcroForm extraction only.
 * pdfjs-dist is excluded from the Hermes bundle (dynamic import() unsupported).
 */
export async function scanTemplatePdf(pdfBytes: Uint8Array) {
  return scanTemplatePdfLite(pdfBytes);
}
```

Damit landen in SQLite typischerweise **`rectJson = null`** und **`page = 1`** für alle Felder.

`detectedFieldsNeedRescan()` in `scan-meta.ts` erkennt fehlende Rechtecke und soll Rescan auslösen — aber `rescanExistingTemplate()` ruft wieder `scanTemplateBytes()` auf, also erneut den Lite-Scan. Der Full-Scan in `pdf-scan-full.ts` ist auf Native **nicht angebunden** (nur über `pdf-scan.web.ts` für Web).

**Relevante Dateien:**

- `expo-toolbox/src/native/bautagebuch/lib/pdf-scan-lite.ts`
- `expo-toolbox/src/native/bautagebuch/lib/pdf-scan-full.ts`
- `expo-toolbox/src/native/bautagebuch/lib/pdf-scan.ts`
- `expo-toolbox/src/native/bautagebuch/lib/scan-meta.ts`

---

## 3. Wird das aktive Feld aus Schritt 2/3 an den PDF Viewer übergeben?

**Ja, korrekt verdrahtet** — aber nur an `SetupPdfFieldPreview`, nicht an `PdfPreviewPanel`.

### Schritt 2 (`SetupAssignStep`)

```tsx
<SetupPdfFieldPreview
  pdfPath={pdfPath}
  detectedFields={detectedFields}
  activeFieldId={currentField?.fieldId || null}
  activeFieldLabel={currentLabel}
  activeFieldPage={currentField?.page || 1}
  assignedFieldIds={assignedFieldIds}
  variant="assign"
/>
```

### Schritt 3 (`SetupFieldsStep`)

Gleiches Muster mit `activeFieldId`, `activeFieldLabel`, `activeFieldPage`, `variant="assign"`.

### WebView-Befehl (`SetupPdfFieldPreview`)

```typescript
postPreviewCommand({
  type: 'setActive',
  fieldId: resolvedActiveFieldId,
  page: Number(activeFieldPage || previewState.page || 1),
  overlayPlacement,
  assignedFieldIds
});
```

### Schritt 1 (`SetupStructureStep`)

Nutzt nur `PdfPreviewPanel` ohne Felddaten — dort sind keine Hervorhebungen vorgesehen.

### Datenverlust vor dem Viewer

Highlights werden aus `detectedFields` gebaut, Felder ohne `rect` werden herausgefiltert:

```typescript
detectedFields
  .filter((field) => Array.isArray(field.rect) && field.rect.length >= 4)
  .map(/* ... */)
```

Bei Lite-Scan → **`highlights = []`** → WebView rendert PDF ohne Boxen, obwohl `setActive` korrekt ankommt.

**Relevante Dateien:**

- `expo-toolbox/src/native/bautagebuch/components/setup-wizard/SetupAssignStep.tsx`
- `expo-toolbox/src/native/bautagebuch/components/setup-wizard/SetupFieldsStep.tsx`
- `expo-toolbox/src/native/bautagebuch/components/SetupPdfFieldPreview.tsx`

---

## 4. Existiert eine Overlay-Komponente über dem PDF?

Es gibt **zwei verschiedene „Overlay“-Konzepte**:

| Komponente | Typ | Zweck |
|---|---|---|
| `#overlay` in WebView-HTML | Div über Canvas | Feld-Rechtecke (`.highlight`, `.highlight.active`, `.highlight.assigned`) |
| `GroupOverlayCards` / `TableMappingOverlay` | RN-UI über Preview | Gruppen-/Tabellen-Auswahl, **keine** Feldmarkierung |
| `PreviewOverlayPanel` | Modal-Container | BTB-Laufzeit-Vorschau, kein Feld-Overlay |
| `PdfPreviewPanel` | Einfache WebView | **Kein** Overlay |

Die eigentliche Feldhervorhebung ist **rein im WebView-HTML** (`pdf-preview-html.ts`), eingebunden über `SetupPdfFieldPreview`.

Bei WebView-Fehler: Fallback auf `PdfPreviewPanel` → komplett ohne Overlays.

**Relevante Dateien:**

- `expo-toolbox/src/native/bautagebuch/lib/pdf-preview-html.ts`
- `expo-toolbox/src/native/bautagebuch/components/SetupPdfFieldPreview.tsx`
- `expo-toolbox/src/native/bautagebuch/components/PdfPreviewPanel.tsx`
- `expo-toolbox/src/native/bautagebuch/components/setup-wizard/GroupOverlayCards.tsx`
- `expo-toolbox/src/native/bautagebuch/components/setup-wizard/TableMappingOverlay.tsx`

---

## 5. Werden x/y/width/height korrekt auf die gerenderte PDF-Größe skaliert?

**Theoretisch ja, praktisch zwei Probleme:**

### A) Keine Koordinaten → nichts zu skalieren

Hauptursache auf Native (siehe Abschnitt 2).

### B) Skalierungsfehler in der Einzelseiten-Ansicht (`buildFieldPreviewHtml`)

Schritt 2/3 nutzen `variant="assign"` → **`buildFieldPreviewHtml`** (Einzelseite), nicht die scrollbare Variante.

In `drawOverlay()`:

```javascript
overlay.style.width = viewportObj.width + 'px';   // Pixel-Dimensionen
overlay.style.height = viewportObj.height + 'px';
const left = Math.min(x1, x2) * viewportScale;   // CSS-Skalierung
const top = (viewportObj.height / viewportScale - Math.max(y1, y2)) * viewportScale;
```

- Canvas-CSS-Größe: `viewportObj.width / dpr`
- Overlay-Größe: `viewportObj.width` (Pixel, ≈ CSS × dpr)

→ Auf Retina-Displays (dpr > 1) sind **Overlay-Container und Highlight-Positionen nicht im gleichen Koordinatensystem**.

Die scrollbare Variante (`buildScrollableFieldPreviewHtml`) macht es korrekt:

```javascript
const cssWidth = viewportObj.width / dpr;
const cssScale = cssWidth / baseViewport.width;
```

**Fazit:** Selbst mit vorhandenen `rect`-Daten wären Markierungen in Schritt 2/3 auf Hi-DPI-Geräten wahrscheinlich **verschoben**, weil `variant="assign"` die fehlerhafte Einzelseiten-HTML-Variante wählt.

---

## 6. Wird die aktuelle PDF-Seite berücksichtigt?

| Modus | Seitenhandling |
|---|---|
| `buildFieldPreviewHtml` (assign/mapping) | Einzelseite; `setActive` / `setPage` → `renderPage(page)`; `pageHighlights(pageNumber)` filtert nach Seite |
| `buildScrollableFieldPreviewHtml` (overlay) | Alle Seiten gestapelt; Overlays pro `.page-sheet` |
| Lite-Scan | Alle Felder auf `page: 1` → falsche Seite selbst bei vorhandenen Rechtecken |

Die Seitenlogik im HTML ist vorhanden; die **Eingangsdaten** (`page` + `rect`) sind auf Native meist unbrauchbar.

---

## 7. Funktioniert die Hervorhebung bei Zoom und Scroll?

| Verhalten | Status |
|---|---|
| Scroll | `#viewport` scrollt `#wrap` (Canvas + Overlay gemeinsam) |
| Pinch-Zoom | CSS `transform: scale()` auf `#wrap` — Canvas und Overlay skalieren gemeinsam |
| Aktives Feld zentrieren | `scrollActiveFieldErgonomic()` nach `setActive` / Seitenwechsel |
| Seitenwechsel bei setActive | Löst `renderPage()` aus → Pinch-Zoom wird zurückgesetzt (`resetPinchZoom()`) |

Zoom/Scroll sind grundsätzlich mitgedacht. Einschränkungen:

1. Ohne Highlights keine sichtbare Wirkung.
2. DPR-Skalierungsfehler → Markierungen nicht pixelgenau am Feld.
3. `setActive` triggert Seiten-Neurendern → Zoom springt zurück.

Es gibt **kein separates RN-Overlay**, das unabhängig vom WebView scrollt/zoomt.

---

## Wo die Daten verloren gehen

```
PDF Import
  └─► scanTemplatePdf (Native) = Lite-Scan
        └─► rect: null, page: 1          ◄── Haupt-Bruch
              └─► detected_fields.rectJson = null
                    └─► SetupPdfFieldPreview filtert alle Felder raus
                          └─► highlights = []
                                └─► drawOverlay() zeichnet nichts

Zusätzlich (sekundär):
  └─► variant="assign" → buildFieldPreviewHtml (DPR-Overlay-Bug)
  └─► WebView-Fehler → Fallback PdfPreviewPanel (kein Overlay)
  └─► Schritt 1: PdfPreviewPanel bewusst ohne Feld-Overlays
```

---

## Beteiligte Komponenten

| Datei | Rolle |
|---|---|
| `pdf-scan-lite.ts` / `pdf-scan.ts` | Feld-Erkennung Native (ohne Positionen) |
| `pdf-scan-full.ts` | Feld-Erkennung mit `rect` (nicht auf Native aktiv) |
| `database.ts` | Persistenz `detected_fields.rectJson` |
| `SetupPdfFieldPreview.tsx` | Brücke RN → WebView (highlights, setActive) |
| `pdf-preview-html.ts` | Rendering + Overlay-Logik im WebView |
| `PdfPreviewPanel.tsx` | Einfache Vorschau ohne Overlays |
| `SetupAssignStep.tsx` / `SetupFieldsStep.tsx` | Übergabe aktives Feld |
| `SetupStructureStep.tsx` | Nur `PdfPreviewPanel` |
| `GroupOverlayCards.tsx` | UI-Overlay, keine PDF-Markierung |
| `pdf-preview-overlay.ts` | Overlay-Platzierung für Scroll-Bias (nicht Feld-Rechtecke) |

---

## Notwendige Änderungen (Empfehlung, ohne Implementierung)

Priorität nach Wirkung:

### 1. Positionsdaten beim Import (kritisch)

Native-Scan muss `rect` und korrekte `page` liefern. Optionen:

- Full-Scan (`pdf-scan-full.ts`) auf Native verfügbar machen (z. B. Scan in WebView/Worker statt Hermes-Bundle), oder
- Positionsdaten aus dem PDF-Preview-WebView (pdf.js) nach dem Laden extrahieren und in `detected_fields` zurückschreiben.

Solange `rectJson` null bleibt, kann keine Hervorhebung funktionieren.

### 2. Preview-Variante für Setup-Schritte

Schritt 2/3 sollten für Feld-Markierung eher `buildScrollableFieldPreviewHtml` mit `variant="overlay"` nutzen (korrekte Skalierung, alle Seiten) — oder den DPR-Fehler in `buildFieldPreviewHtml.drawOverlay()` beheben (Overlay-Größe und Positionen auf CSS-Pixel normalisieren).

### 3. Fallback-Verhalten absichern

Wenn `highlights.length === 0`, sollte sichtbar gemacht werden, dass Positionsdaten fehlen (der Hint existiert nur für `variant="default"`, nicht für `assign`).

### 4. Rescan-Pipeline reparieren

`detectedFieldsNeedRescan()` erkennt fehlende Rechtecke, aber `rescanExistingTemplate()` nutzt wieder Lite-Scan — Rescan bringt aktuell keine Verbesserung.

### 5. Terminologie

Es gibt kein `pdfPosition`-Feld. Für künftige APIs reicht das bestehende Modell: `rect: number[]` + `page: number`.

---

## Kurzantwort auf die 7 Prüfpunkte

| # | Frage | Ergebnis |
|---|---|---|
| 1 | Speicherort Felder | SQLite `detected_fields` + Setup-Modell nach Zuordnung |
| 2 | pdfPosition | **Existiert nicht**; Position = `rect` + `page` in `rectJson` |
| 3 | Aktives Feld übergeben | **Ja** in Schritt 2/3 via `SetupPdfFieldPreview` |
| 4 | Overlay-Komponente | **Ja**, WebView-intern (`#overlay`); RN-Overlays sind UI, keine Feldmarkierung |
| 5 | Skalierung | Logik vorhanden, aber **Native ohne rect** + **DPR-Bug in Einzelseiten-Modus** |
| 6 | Seite | HTML filtert nach Seite; Lite-Scan liefert `page: 1` für alle |
| 7 | Zoom/Scroll | Konzept vorhanden; wirkungslos ohne Highlights, DPR-Fehler bei vorhandenen Daten |

---

## Root Cause

Auf Native werden beim PDF-Import **keine Widget-Rechtecke** extrahiert → leeres `highlights`-Array → leeres Overlay im WebView. Die Viewer-Pipeline ist größtenteils gebaut, bekommt aber **keine Positionsdaten**.

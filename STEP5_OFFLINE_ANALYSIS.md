# STEP 5 – Offline PDF Preview Engine Analyse

**Stand:** Vor Migration auf lokale pdf.js-Assets  
**Scope:** Nur PDF-Vorschau-Pipeline — keine Änderungen an Mapping, ETB, Export, Datenmodell oder Wizard-State

---

## 1. Aktuelle CDN-Nutzung (Vor Migration)

### Betroffene Dateien

| Datei | CDN-Nutzung |
|-------|-------------|
| `expo-toolbox/src/native/bautagebuch/lib/pdf-preview-html.ts` | `PDFJS_CDN_BASE` → cdnjs.cloudflare.com |
| `SetupPdfFieldPreview.tsx` | Indirekt via `buildFieldPreviewHtml()` |
| `PdfPreviewPanel.tsx` | Indirekt via `buildSimplePdfPreviewHtml()` |

### CDN-URLs (pdf.js 3.11.174)

```
https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js
https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js
```

### Pipeline vor Migration

```
React Native
  → WebView (inline HTML)
    → <script src="CDN/pdf.min.js">     ← Netzwerk erforderlich
    → GlobalWorkerOptions.workerSrc = CDN/pdf.worker.min.js  ← Netzwerk erforderlich
      → Canvas-Rendering + Overlays
```

### Nicht betroffen (bleiben unverändert)

| Bereich | Runtime |
|---------|---------|
| PDF-Scan (STEP 4) | `pdfjs-dist` npm + `assets/pdf.worker.min.mjs` (lokal) |
| PDF-Export | pdf-lib / eigene Pipeline |
| ETB / Legacy-Editor | Keine WebView-Vorschau |

---

## 2. Benötigte Assets

| Asset | Pfad | Größe (ca.) | Zweck |
|-------|------|-------------|-------|
| pdf.js Core | `expo-toolbox/assets/pdfjs/pdf.min.js` | ~313 KB | WebView-Runtime (`window.pdfjsLib`) |
| pdf.js Worker | `expo-toolbox/assets/pdfjs/pdf.worker.min.js` | ~1,1 MB | Hintergrund-Parsing im WebView |

**Version:** pdf.js **3.11.174** (identisch zur bisherigen CDN-Version — keine Versionsänderung)

### Bereits vorhanden (Scan-Pipeline, separate Runtime)

- `expo-toolbox/assets/pdf.worker.min.mjs` — für `pdfjs-dist` v4 Scan, nicht für WebView-Vorschau

---

## 3. Abhängigkeiten & Risiken

### WebView-Constraints

| Thema | Risiko | Mitigation |
|-------|--------|------------|
| Worker-URL | WebView kann keine relativen Worker-Pfade auflösen | `Asset.fromModule()` → `file://` URI |
| Inline-Script-Größe | pdf.min.js ~313 KB im HTML | Einmalig pro Session cachen (`loadPdfPreviewAssets`) |
| `</script>` in JS-Quelltext | Bricht HTML-Parsing | `escapeInlineScript()` |
| Große PDFs (50+ Seiten) | Speicher / Render-Zeit | Bestehendes 12-Mio-Pixel-Cap bleibt |

### Metro-Bundling

`.js`-Dateien unter `assets/` werden nicht automatisch als Assets behandelt.  
→ Metro `resolveRequest` erweitern (analog zu `pdf.worker.min.mjs`).

---

## 4. Migrationsstrategie

### Phase A — Assets bereitstellen

1. `assets/pdfjs/pdf.min.js` und `pdf.worker.min.js` von cdnjs 3.11.174 kopieren
2. Metro-Konfiguration: beide Dateien als `assetFiles` registrieren

### Phase B — Asset-Loader

Neue Datei `pdf-preview-assets.ts`:

```typescript
Asset.fromModule(require('assets/pdfjs/pdf.min.js'))
  → FileSystem.readAsStringAsync(localUri)  // Core inline
Asset.fromModule(require('assets/pdfjs/pdf.worker.min.js'))
  → workerSrc = localUri                    // Worker URI
```

Ergebnis wird pro App-Session gecacht.

### Phase C — HTML-Engine umbauen

`pdf-preview-html.ts`:

| Vorher | Nachher |
|--------|---------|
| `<script src="CDN/...">` | `<script>/* inline pdf.min.js */</script>` |
| `workerSrc = CDN URL` | `workerSrc = Expo Asset URI` |
| Generische Fehlermeldung | `"PDF Vorschau konnte nicht geladen werden"` + `console.error` |

### Phase D — Komponenten anbinden

`SetupPdfFieldPreview` und `PdfPreviewPanel`:

```typescript
Promise.all([readPdfBase64(), loadPdfPreviewAssets()])
  → buildFieldPreviewHtml({ base64, pdfJsSource, workerSrc, ... })
```

### Phase E — Tests & Validierung

- Unit-Tests: keine CDN-URLs, Worker-URL vorhanden, Offline-Simulation
- Manuell: Flugmodus → PDF importieren → Mapping → Vorschau + Zoom + Highlights

---

## 5. Unveränderte Funktionen (Regression-Schutz)

| Feature | Mechanismus |
|---------|-------------|
| Vollbild (Overlay) | `PreviewOverlayPanel` + flex Layout |
| Pinch-Zoom | Touch-Handler in WebView-HTML |
| Feld-Highlight | `#overlay` + `.highlight.active/.dim` |
| Scroll zum Feld | `scrollActiveFieldErgonomic()` |
| Overlay-Position | `postMessage({ overlayPlacement })` |
| Multi-Page | `setPage` / `renderPage()` |

---

## 6. Abgrenzung

**In Scope (STEP 5):**

- `pdf-preview-html.ts`
- `pdf-preview-assets.ts` (neu)
- `SetupPdfFieldPreview.tsx`
- `PdfPreviewPanel.tsx`
- `assets/pdfjs/*`
- `metro.config.js` (Asset-Registrierung)
- Tests + Dokumentation

**Out of Scope:**

- `SetupEditor.tsx`, ETB, PDF-Export
- Mapping Workflow STEP 1–4
- Datenmodell, Wizard State
- Scan-Pipeline (`pdf-scan*.ts`)

---

## 7. Empfehlung für STEP 6

Nach erfolgreichem Offline-Preview:

1. **Asset-Größen optimieren** — optional pdf.js Legacy-Build prüfen
2. **Scan/Vorschau-Vereinheitlichung** — zwei pdf.js-Versionen (3.11 WebView, 4.x Scan) dokumentieren oder langfristig harmonisieren
3. **On-Device Performance-Profiling** — 50-Seiten-PDF mit `renderMs`-Metrik auswerten

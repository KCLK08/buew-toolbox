# STEP 5 – Offline PDF Preview Engine Ergebnis

**Stand:** Nach Migration auf lokale pdf.js 3.11.174 Assets  
**Scope:** Nur PDF-Vorschau-Pipeline

---

## Geänderte Dateien

| Datei | Änderung |
|-------|----------|
| `expo-toolbox/assets/pdfjs/pdf.min.js` | **Neu** — pdf.js 3.11.174 Core (~313 KB) |
| `expo-toolbox/assets/pdfjs/pdf.worker.min.js` | **Neu** — pdf.js 3.11.174 Worker (~1,1 MB) |
| `expo-toolbox/src/native/bautagebuch/lib/pdf-preview-assets.ts` | **Neu** — Expo Asset Loader + Session-Cache |
| `expo-toolbox/src/native/bautagebuch/lib/pdf-preview-html.ts` | CDN entfernt, Inline-Bundle, Fehlerbehandlung |
| `expo-toolbox/src/native/bautagebuch/lib/pdf-preview-html.test.ts` | **Neu** — 7 Offline-Unit-Tests |
| `expo-toolbox/src/native/bautagebuch/components/SetupPdfFieldPreview.tsx` | Lädt lokale Assets vor HTML-Build |
| `expo-toolbox/src/native/bautagebuch/components/PdfPreviewPanel.tsx` | Lädt lokale Assets vor HTML-Build |
| `expo-toolbox/metro.config.js` | Registriert `assets/pdfjs/*.js` als Metro-Assets |
| `STEP5_OFFLINE_ANALYSIS.md` | Analyse-Dokument (Vor Migration) |

**Unverändert:** `SetupEditor.tsx`, ETB, PDF-Export, Mapping STEP 1–4, Datenmodell, Wizard State, `PreviewOverlayPanel.tsx` (nur Container)

---

## Architektur vorher/nachher

### Vorher

```
SetupPdfFieldPreview / PdfPreviewPanel
  → build*PreviewHtml(base64)
    → WebView HTML
      → <script src="https://cdnjs.cloudflare.com/.../pdf.min.js">  ❌ Netzwerk
      → workerSrc = CDN URL                                           ❌ Netzwerk
```

### Nachher

```
SetupPdfFieldPreview / PdfPreviewPanel
  → loadPdfPreviewAssets()  [einmal pro Session, gecacht]
      → Asset.fromModule(pdf.min.js) → readAsStringAsync → pdfJsSource
      → Asset.fromModule(pdf.worker.min.js) → workerSrc (file://)
  → build*PreviewHtml({ base64, pdfJsSource, workerSrc })
    → WebView HTML
      → <script>/* inline pdf.min.js */</script>  ✅ offline
      → GlobalWorkerOptions.workerSrc = file://… ✅ offline
```

---

## Offline Status

| Prüfpunkt | Status |
|-----------|--------|
| Keine cdnjs.cloudflare.com URLs im HTML | ✅ |
| pdf.js Core aus App-Bundle | ✅ |
| Worker aus App-Assets (Expo Asset URI) | ✅ |
| Keine CDN-Fallbacks | ✅ |
| Fehlertext bei Render-Fehler | ✅ `"PDF Vorschau konnte nicht geladen werden"` |
| console.error bei Boot/Worker/Render-Fehler | ✅ |

### Offline-Testszenario (manuell auf Gerät)

1. App installieren und einmal öffnen (Assets werden gebündelt)
2. Flugmodus aktivieren
3. PDF importieren
4. Mapping starten

**Erwartung:** PDF wird geladen, Felder angezeigt, Highlight + Zoom funktionieren — ohne Netzwerk.

---

## Tests

### Automatisiert

```bash
npx tsx src/native/bautagebuch/lib/pdf-preview-html.test.ts
```

| Test | Ergebnis |
|------|----------|
| CDN URL darf nicht enthalten sein | ✅ |
| Lokale Assets werden inline geladen | ✅ |
| Worker URL in GlobalWorkerOptions | ✅ |
| Offline-Modus-Simulation (kein `<script src=`) | ✅ |
| Bundled Asset-Dateien vorhanden | ✅ |
| Legacy API ohne Assets wird abgelehnt | ✅ |
| Script-Tag-Escaping | ✅ |

**7/7 Tests bestanden**

### Typecheck

```bash
npm run typecheck  # ✅
```

### Performance (bestehende Mechanismen unverändert)

| Szenario | Schutz |
|----------|--------|
| 1 Seite | Standard-Render |
| 10 Seiten | Seitenwechsel via `setPage`, kein Multi-Render |
| 50 Seiten | 12-Mio-Pixel-Cap reduziert Scale automatisch |

Features erhalten: Vollbild, Pinch-Zoom, Feld-Highlight, Scroll zum Feld, Overlay-Position, Multi-Page.

---

## Bekannte Einschränkungen

1. **Zwei pdf.js-Runtimes** — WebView-Vorschau nutzt 3.11.174 (assets/pdfjs), Scan-Pipeline nutzt pdfjs-dist v4 + `.mjs` Worker. Bewusst getrennt gehalten (Out of Scope STEP 5).

2. **HTML-Größe** — pdf.min.js wird inline in WebView-HTML eingebettet (~313 KB pro HTML-Instanz). Session-Cache vermeidet wiederholtes FileSystem-Lesen, nicht die HTML-Größe selbst.

3. **Erststart** — Assets müssen einmal aus dem App-Bundle geladen werden (beim ersten Preview-Aufruf). Kein separates Netzwerk, aber kurzer Ladeindikator möglich.

4. **Web-Plattform** — Worker-URI kann auf Web `http://localhost` sein (Dev-Server). Native Offline-Szenario (Flugmodus) ist das primäre Ziel.

---

## Empfehlung STEP 6

1. **Preview-Asset-Warmup** — `loadPdfPreviewAssets()` beim Wizard-Start vorladen, um ersten Preview-Aufruf zu beschleunigen
2. **pdf.js-Versionen harmonisieren** — langfristig Scan (v4) und Preview (v3.11) auf eine Version evaluieren
3. **Geräte-Performance-Log** — `renderMs` aus WebView-State für 1/10/50-Seiten-PDFs auf Low-End-Geräten sammeln

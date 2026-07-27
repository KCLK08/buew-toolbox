# STEP 5 Ergebnis

**Stand:** Offline PDF Preview Engine — pdf.js 3.11.174 lokal gebündelt  
**Scope:** Nur PDF-Vorschau-Pipeline (keine Änderungen an Mapping, Scan, ETB, Export, Datenmodell)

---

## Geänderte Dateien

| Datei | Änderung |
|-------|----------|
| `expo-toolbox/assets/pdfjs/pdf.min.js` | **Neu** — pdf.js 3.11.174 Core (~313 KB, aus CDN-Version) |
| `expo-toolbox/assets/pdfjs/pdf.worker.min.js` | **Neu** — pdf.js 3.11.174 Worker (~1,1 MB, aus CDN-Version) |
| `expo-toolbox/src/native/bautagebuch/lib/pdf-preview-assets.ts` | **Neu** — `loadPdfPreviewAssets()` mit Expo Asset + Session-Cache |
| `expo-toolbox/src/native/bautagebuch/lib/pdf-preview-html.ts` | CDN entfernt, Inline-Script, `escapeInlineScript()`, Fehlerbehandlung |
| `expo-toolbox/src/native/bautagebuch/lib/pdf-preview-html.test.ts` | **Neu** — 7 Unit-Tests für Offline-Preview |
| `expo-toolbox/src/native/bautagebuch/components/SetupPdfFieldPreview.tsx` | `Promise.all` Base64 + Assets vor HTML-Build |
| `expo-toolbox/src/native/bautagebuch/components/PdfPreviewPanel.tsx` | `Promise.all` Base64 + Assets vor HTML-Build |
| `expo-toolbox/metro.config.js` | `assets/pdfjs/*.js` als Metro-Assets registriert |

**Unverändert:** `SetupEditor.tsx`, ETB, PDF-Export, Datenmodell, Wizard State, Mapping STEP 1–4, Scan-Pipeline STEP 4

---

## Architektur vorher/nachher

### Vorher

```
React Native
  → WebView (inline HTML)
    → <script src="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js">
    → GlobalWorkerOptions.workerSrc = CDN-Worker-URL
```

Netzwerk für pdf.js Core und Worker **erforderlich**.

### Nachher

```
React Native
  → loadPdfPreviewAssets()  [Session-Cache]
      → Asset.fromModule(pdf.min.js) → readAsStringAsync → pdfJsSource
      → Asset.fromModule(pdf.worker.min.js) → workerSrc (file://)
  → WebView (inline HTML)
    → <script>${pdfJsSource}</script>   ← lokal inline
    → GlobalWorkerOptions.workerSrc = workerSrc  ← lokale Asset-URI
```

Kein CDN, keine externen Requests für die Preview-Runtime.

---

## Offline Status

| Bereich | Status | Anmerkung |
|---------|--------|-----------|
| PDF Import | ✓ | Unverändert (lokale Datei, Scan-Pipeline STEP 4) — kein CDN |
| PDF Vorschau | ✓ | pdf.min.js inline aus App-Bundle |
| Overlay | ✓ | Feld-Highlights + Overlay-Position unverändert |
| Zoom | ✓ | Pinch-Zoom im WebView-HTML unverändert |

**Manueller Offline-Test (Gerät):**

1. App installieren und einmal öffnen
2. Flugmodus aktivieren
3. PDF importieren → Mapping starten
4. Erwartung: PDF sichtbar, Felder markiert, Zoom funktioniert

---

## Tests

```bash
cd expo-toolbox
npm run typecheck
npx tsx src/native/bautagebuch/lib/pdf-preview-html.test.ts
```

| # | Test | Ergebnis |
|---|------|----------|
| 1 | CDN URL — HTML enthält nicht `cdnjs.cloudflare.com` | ✅ |
| 2 | Lokale Assets — `pdfJsSource` inline im HTML | ✅ |
| 3 | Worker — `workerSrc` in `GlobalWorkerOptions.workerSrc` | ✅ |
| 4 | Script Escape — `</script>` wird zu `<\/script>` | ✅ |
| 5 | HTML Generator — gültiges Preview-HTML mit Fehler-UI | ✅ |
| 6 | Legacy API ohne Assets wird abgelehnt | ✅ |
| 7 | Bundled Asset-Dateien vorhanden (3.11.174) | ✅ |

**Ergebnis: 7/7 Tests bestanden · Typecheck ✅**

---

## Bekannte Einschränkungen

1. **Zwei pdf.js-Runtimes** — WebView-Vorschau: 3.11.174 (`assets/pdfjs/`); Scan-Pipeline: pdfjs-dist v4 — bewusst getrennt (Out of Scope STEP 5).

2. **HTML-Größe** — pdf.min.js (~313 KB) wird pro WebView-Instanz inline eingebettet. Session-Cache vermeidet wiederholtes FileSystem-Lesen, nicht die HTML-Größe.

3. **Erststart** — Beim ersten Preview-Aufruf kurzer Ladeindikator während Asset-Laden aus dem Bundle.

4. **Web Dev** — Worker-URI kann unter Expo Web `http://localhost` sein; primäres Ziel ist natives Offline (Flugmodus).

---

## Empfehlung STEP 6

1. **Asset-Warmup** — `loadPdfPreviewAssets()` beim Wizard-Start vorladen
2. **Setup-Tab Bugfixes** — Read-only für archivierte Legacy-Vorlagen, Mapping-Index-Korrektur
3. **Performance-Profiling** — `renderMs`-Metrik für 1/10/50-Seiten-PDFs auf Low-End-Geräten auswerten

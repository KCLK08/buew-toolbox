# STEP 4 – Scan Ergebnis

**Stand:** Native PDF Scan Pipeline verbessert  
**Scope:** Scan-Bibliotheken only — keine Änderungen an SetupEditor, ETB, PDF Export, Mapping UX (STEP 1–3)

---

## Geänderte Dateien

| Datei | Änderung |
|-------|----------|
| `pdf-scan-types.ts` | **Neu** — `PdfScanResult`, `ScanFieldResult`, `hasRects` |
| `pdf-scan-shared.ts` | **Neu** — Gemeinsame Scan-Hilfsfunktionen |
| `pdf-scan-full.ts` | **Neu** — Vollscan (pdf-lib + pdfjs-dist), Web + Native |
| `pdf-scan.ts` | Native: Full Scan → Lite Fallback |
| `pdf-scan.web.ts` | Re-Export aus `pdf-scan-full.ts` |
| `pdf-scan-lite.ts` | Refactor auf Shared + `PdfScanResult` |
| `scan-meta.ts` | Version 4, erweiterte Rescan-Prüfung |
| `pdf-scan.test.ts` | **Neu** — 10 Testfälle |
| `STEP4_SCAN_ANALYSIS.md` | Analyse-Dokumentation |

**Unverändert:** `SetupEditor.tsx`, ETB Template, PDF Export, Mapping UX, DB-Schema, `SetupPdfFieldPreview` (filtert weiterhin Felder ohne Rect)

---

## Neue Funktionen

| Funktion | Beschreibung |
|----------|--------------|
| `scanTemplatePdfFull()` | Vollscan mit Widget-Rects, Seiten, Label-Inferenz |
| `PdfScanResult` | Einheitliches `{ fields, pageCount, scanVersion, hasRects }` |
| `withDetectedFieldsAlias()` | Rückwärtskompatibilität `detectedFields` |
| `resultHasRects()` | Qualitätsflag für Pipeline |
| `checkMappingTransition` / Rescan | Erweitert um fieldId, page, unsupported |

---

## Scan Architektur vorher / nachher

### Vorher

```
Native: pdf-scan.ts → pdf-scan-lite (page=1, rect=null)
Web:    pdf-scan.web.ts → pdf-lib + pdfjs (vollständig)
```

### Nachher

```
Native: pdf-scan.ts
    ├─ Stufe 1: pdf-scan-full (pdf-lib + pdfjs-dist + Worker Asset)
    └─ Stufe 2: pdf-scan-lite (Fallback)

Web: pdf-scan.web.ts → pdf-scan-full (identische Logik)

Beide: PdfScanResult { fields, pageCount, scanVersion, hasRects }
         + detectedFields Alias für templateService
```

### Scan-Strategie

1. **Stufe 1:** pdfjs Widget-Annotations → `page`, `rect`, `options`
2. **Stufe 2:** Bei Fehler → Lite (Felder importieren, `rect: null`)
3. **Rescan:** `ETB_SCAN_VERSION = 4` triggert Re-Import bestehender Templates

---

## Test Ergebnisse

| Test | Ergebnis |
|------|----------|
| PDF mit Textfeldern (Lite) | ✓ fieldName, type, fieldId |
| PDF mit Textfeldern (Full) | ✓ page + rect + hasRects |
| PDF Checkbox | ✓ type checkbox |
| PDF ohne AcroForm | ✓ Fehler geworfen |
| Mehrseitiges PDF (Lite) | ✓ pageCount=2, page fallback 1 |
| Mehrseitiges PDF (Full) | ✓ page=2 |
| Rescan bei fehlendem rect | ✓ true |
| Rescan bei fehlendem fieldId | ✓ true |
| Rescan bei vollständigen Daten | ✓ false |
| `resultHasRects()` | ✓ |
| `npm run typecheck` | ✓ 0 Fehler |

Ausführung: `npx tsx src/native/bautagebuch/lib/pdf-scan.test.ts`

---

## Bekannte Einschränkungen

| Einschränkung | Detail |
|---------------|--------|
| pdfjs Worker | Erfordert Expo Asset oder node_modules Worker-Pfad |
| Font-Warnings in Node | pdfjs `standardFontDataUrl` — in App via Expo Asset OK |
| Lite Fallback | Weiterhin `page: 1`, `rect: null` wenn Full Scan fehlschlägt |
| Kein OCR | Nur AcroForm — nicht-interaktive PDFs weiterhin abgelehnt |
| Label-Inferenz | Abhängig von PDF-Textlayer — nicht immer verfügbar |

---

## Empfehlung STEP 5

1. **Offline pdf.js** in WebView (STEP 2 Dokumentation umsetzen)
2. **Geräte-Tests** mit echten Baustellen-PDFs auf Android/iOS
3. **Automatischer Rescan** beim Öffnen in_progress Templates ohne `hasRects`
4. **Monitoring:** `hasRects` in Setup-UI anzeigen wenn Fallback aktiv

---

STEP 4 abgeschlossen. Keine weiteren automatischen Verbesserungen.

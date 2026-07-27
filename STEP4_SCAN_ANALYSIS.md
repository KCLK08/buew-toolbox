# STEP 4 – Scan Pipeline Analyse

**Stand:** Vor Implementierung der Native Full-Scan-Pipeline  
**Scope:** `expo-toolbox/src/native/bautagebuch/lib/pdf-scan*`

---

## Aktuelle Architektur

```
templateService.scanTemplateBytes()
    │
    ├─► scanTemplatePdf()     ← Platform-Datei (.ts / .web.ts)
    │       │
    │       ├─ [Web] pdf-scan.web.ts → pdf-scan-full.ts (pdf-lib + pdfjs-dist)
    │       │
    │       └─ [Native] pdf-scan.ts → bisher nur pdf-scan-lite.ts Re-Export
    │
    └─► catch → scanTemplatePdfLite() (pdf-lib only)
            │
            ▼
        SQLite detected_fields + buildGenericSetupModel / ETB rescan
```

### Dateien

| Datei | Plattform | Funktion |
|-------|-----------|----------|
| `pdf-scan.ts` | iOS/Android | Re-exportierte **Lite** als `scanTemplatePdf` |
| `pdf-scan.web.ts` | Web | Vollscan mit pdfjs Widget-Annotations |
| `pdf-scan-lite.ts` | Fallback | AcroForm via pdf-lib, `page: 1`, `rect: null` |
| `scan-meta.ts` | Alle | `ETB_SCAN_VERSION`, `detectedFieldsNeedRescan()` |

### Genutzte Bibliotheken

| Bibliothek | Lite | Full (Web) |
|------------|------|------------|
| pdf-lib ^1.17.1 | ✓ AcroForm-Felder, Typ, Options | ✓ AcroForm-Basis |
| pdfjs-dist ^4.8.69 | ✗ | ✓ Widget-Rects, Seiten, Label-Inferenz |
| assets/pdf.worker.min.mjs | ✗ | ✓ Worker für pdfjs |

---

## Probleme

| Problem | Auswirkung |
|---------|------------|
| Native = Lite only | Keine `rect`-Daten auf iOS/Android |
| `page: 1` Fallback | Mehrseitige PDFs falsch zugeordnet |
| Overlay / Mapping | `SetupPdfFieldPreview` filtert Felder ohne `rect` |
| Ergonomisches Scrollen | `resolveOverlayPlacement` ohne Rect nutzlos |
| Rescan v3 | Erkennt fehlende Rects, aber Native liefert keine |

### Fehlende Informationen (Native Lite)

- `rect` (Feldposition)
- Korrekte `page` (Widget-Seite)
- `labelCandidate` (nur humanisierter fieldName, keine PDF-Text-Inferenz)
- `hasRects`-Flag für Pipeline-Qualität

---

## Lösungsvorschlag (STEP 4)

1. **`pdf-scan-full.ts`** — Gemeinsame Vollscan-Implementierung (pdf-lib + pdfjs-dist), bereits auf Web bewährt
2. **`pdf-scan.ts` (Native)** — Stufe 1: Full Scan, Stufe 2: Lite Fallback
3. **`pdf-scan-types.ts`** — Einheitliches `PdfScanResult` mit `fields`, `hasRects`, `scanVersion`
4. **`pdf-scan-shared.ts`** — Gemeinsame Hilfsfunktionen (IDs, Sortierung, Label-Inferenz)
5. **`scan-meta.ts`** — Version 4, erweiterte `detectedFieldsNeedRescan()` (fieldId, page, rect, unsupported)
6. **Keine neuen Dependencies** — pdfjs-dist und Worker-Asset bereits vorhanden
7. **SetupPdfFieldPreview** — unverändert; filtert weiterhin Felder ohne Rect für Overlays

### Erwartetes Ergebnis nach STEP 4

```typescript
{
  fields: [{ fieldId, fieldName, type, page, rect, labelCandidate, options, orderIndex }],
  pageCount: number,
  scanVersion: "4",
  hasRects: boolean
}
```

Native und Web liefern identisches Format; bei pdfjs-Fehler Fallback auf Lite mit `hasRects: false`.

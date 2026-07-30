# AcroForm-Feldtyperkennung (Setup Wizard)

## Ausgangslage

Beim PDF-Import werden AcroForm-Felder über `pdf-lib` gescannt (`setup-model.js`). Zuvor landeten nahezu alle erkannten Felder als `text`, weil die Typableitung hauptsächlich auf JavaScript-`instanceof`-Prüfungen basierte. In manchen Laufzeitumgebungen (Bundler, Hermes) stimmen Konstruktor-Namen bzw. Prototyp-Ketten nicht zuverlässig mit den pdf-lib-Klassen überein.

## Technische Einordnung

**Kein grundsätzliches AcroForm-Limit**, sondern ein Erkennungsproblem in unserer Scan-Pipeline:

| Quelle | Verfügbarkeit |
|--------|----------------|
| AcroForm-Dictionary `/FT` (Field Type) | Ja — direkt aus dem PDF |
| AcroForm-Flags `/Ff` (Field Flags) | Ja — für Btn/Radio/Checkbox |
| pdf-lib-Klassen (`PDFTextField`, `PDFCheckBox`, …) | Ja, wenn `instanceof` greift |
| Datum als eigener PDF-Typ | Nein — kein separater AcroForm-Typ; Datumsfelder sind `/FT /Tx` (Text) |
| OCR / visuelle Heuristik | Nicht Teil des AcroForm-Scans |

## Implementierte Erkennung

`detectAcroFormFieldType()` liest das AcroForm-Wörterbuch und mappt:

| `/FT` | `/Ff` / Zusatz | App-Typ |
|-------|----------------|---------|
| `Tx` | — | `text` |
| `Ch` | — | `dropdown` |
| `Sig` | — | `signature` |
| `Btn` | Bit 16 (`0x8000`) gesetzt | `unsupported` (Pushbutton) |
| `Btn` | Bit 24 (`0x800000`) gesetzt | `radio` |
| `Btn` | Mehrere Optionen via `getOptions()` | `radio` |
| `Btn` | sonst | `checkbox` |

Fallback-Kette in `resolvePdfFieldType()`:

1. pdf-lib `instanceof` (PDFTextField, PDFCheckBox, …)
2. AcroForm `/FT` + `/Ff` (`detectAcroFormFieldType`)
3. Konstruktor-Name-Heuristik (`includes('text')`, …)
4. `unsupported` → beim Import als `text` gespeichert

## Datum-Felder

Datumsfelder haben in PDF keinen eigenen Feldtyp. Sie erscheinen als Textfelder (`/FT /Tx`), ggf. mit Format-Hinweisen in `/AA` oder `/DV`. Eine **eindeutige** automatische Erkennung ist ohne zusätzliche Heuristiken (Feldname, Format-String) nicht zuverlässig möglich. Empfehlung: in Schritt 3 manuell auf Typ „Datum“ setzen.

## Fallback

Wenn weder `instanceof` noch AcroForm-Dictionary auslesbar sind, wird `text` als sicherer Standard verwendet. Der Nutzer kann den Typ in Schritt 3 korrigieren.

## Betroffene Dateien

- `src/native/bautagebuch/lib/setup-model.js` — `detectAcroFormFieldType`, `resolvePdfFieldType`
- Tests: `src/native/bautagebuch/lib/pdf-scan.test.ts`

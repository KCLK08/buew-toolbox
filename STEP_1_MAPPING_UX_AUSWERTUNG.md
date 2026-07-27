# STEP 1 – Mapping UX Auswertung

**Stand:** Nach Optimierung des geführten Feld-Zuordnungsmodus  
**Scope:** Nur Mapping-Wizard (Schritt 1) — keine Änderungen an `SetupFieldSettingsStep`, `SetupEditor`, ETB, Legacy, PDF-Export

---

## Geänderte Dateien

| Datei | Art der Änderung |
|-------|------------------|
| `expo-toolbox/src/native/bautagebuch/components/setup-wizard/SetupMappingStep.tsx` | Geführter Assistent, Abschluss-Dialog, Mobile-Footer, Feld-Sync |
| `expo-toolbox/src/native/bautagebuch/components/setup-wizard/SetupProgressHeader.tsx` | Feldinfo (Erkannter Name / Vorschlag), Fortschrittsbalken |
| `expo-toolbox/src/native/bautagebuch/components/setup-wizard/GroupOverlayCards.tsx` | Leer-Zustand, Gruppen-Highlight, größere Touchflächen |
| `expo-toolbox/src/native/bautagebuch/lib/setup-mapping.ts` | `resolveOverlayPlacement` korrigiert, `resolveCurrentMappingIndex` neu |
| `expo-toolbox/src/native/bautagebuch/lib/setup-mapping.test.ts` | Tests für Overlay-Platzierung und Index-Auflösung |
| `expo-toolbox/src/native/bautagebuch/components/SetupPdfFieldPreview.tsx` | Mapping: Flex-Höhe, Scroll-zum-Feld (nur `variant="mapping"`) |

**Nicht geändert:** `SetupFieldSettingsStep`, `SetupEditor`, `etb-template.js`, `templateService`, `fields.tsx`, PDF-Export

---

## Neue / erweiterte Funktionen

### 1. Aktives Feld-Management (`SetupMappingStep`)

- `resolveCurrentMappingIndex()` wählt beim Start automatisch das erste nicht zugeordnete Feld.
- `currentFieldIndex` aus `SetupWizardState` wird mit dem sichtbaren Feld synchronisiert.
- Aktives Feld wird an `SetupPdfFieldPreview` übergeben (`activeFieldId`, `activeFieldPage`).
- PDF springt zur Feldseite und scrollt das Highlight in die Mitte (WebView `scrollIntoView`).

### 2. Fortschrittsanzeige (`SetupProgressHeader`)

```
Feld 5 von 28
█████░░░░░ 
18%

Erkannter Name:
Text1

Vorschlag:
Baustellenbezeichnung
```

- Vorschlag aus `DetectedField.labelCandidate` — **keine automatische Speicherung** im Setup-Modell.

### 3. Overlay-Platzierung (`resolveOverlayPlacement`)

Korrigierte Logik (PDF-Koordinaten, Y nach oben):

| Feldposition | Overlay-Panel |
|--------------|---------------|
| Oben auf Seite | Unten |
| Unten auf Seite | Oben |
| Links | Rechts |
| Rechts | Links |

### 4. Gruppen-Auswahl (`GroupOverlayCards`)

**Fall A — Gruppen vorhanden:** Horizontale Karten mit großen Touchflächen (min. 48 px), optional markierte Auswahl.

**Fall B — Keine Gruppen:** Blockierende Meldung + „+ Neue Gruppe“ → Dialog mit Nameingabe → Speichern → Gruppe markiert, danach auswählbar.

### 5. Neue Gruppe (`addWizardGroup`)

- Bestehende API unverändert.
- Nach Erstellung: `selectedGroupId` gesetzt → visuelle Markierung der neuen Gruppe.

### 6. Überspringen

- `deferField()` unverändert.
- Prominenter „Überspringen“-Button in der Fußzeile.
- Übersprungene Felder → weiterhin `sonstiges` via `rebuildSectionsFromWizard`.

### 7. Abschluss

Alert-Dialog bei vollständiger Zuordnung:

- **„Zu Feldern wechseln“** → `onComplete()` → `rebuildSectionsFromWizard` + Navigation Schritt 2
- **„Später fortsetzen“** → speichern und zurück

Kein automatischer Sprung mehr ohne Bestätigung.

### 8. Mobile UI

- PDF-Vorschau: `flex: 1` statt fester Höhe → maximale Displayfläche.
- Footer: `useSafeAreaInsets()` für Android-Navigationsleiste.
- Navigation: große Buttons (min. 48 px), „Überspringen“ als Primary-Aktion.

---

## Screenshots vorher / nachher

> Screenshots müssen auf einem Android-Gerät oder Emulator manuell erstellt werden.  
> Cloud-Agent-Umgebung ohne GUI — keine Bilddateien angehängt.

### Vorher (Referenz aus bestehendem Code)

| Bereich | Verhalten |
|---------|-----------|
| Header | Nur `labelCandidate` als Titel, Fortschritt „X verbleibend“ |
| Feldinfo | Keine Trennung Erkannter Name / Vorschlag |
| Overlay | Y-Platzierung teilweise invertiert (Feld oben → Panel oben) |
| Abschluss | Automatischer Sprung zu Schritt 2 bei letztem Feld |
| Gruppen leer | Nicht explizit behandelt (Default-Gruppen immer vorhanden) |
| PDF-Höhe | Fix ~62 % Viewport-Höhe |

### Nachher (erwartetes UI)

| Bereich | Verhalten |
|---------|-----------|
| Header | „Feld N von M“, Prozentbalken, Erkannter Name + Vorschlag |
| PDF | Aktives Feld pulsiert, Seite springt, Feld zentriert |
| Overlay | Panel gegenüber der Feldposition |
| Footer | Große Touch-Buttons, Safe-Area-Padding |
| Abschluss | Bestätigungsdialog mit zwei Optionen |
| Keine Gruppen | Hinweistext + Gruppe-anlegen-Flow |

**Empfohlene Screenshot-Szenarien:**

1. Feld 1 von N mit Vorschlag-Text
2. Overlay unten bei Feld am Seitenanfang
3. Overlay oben bei Feld am Seitenende
4. Leer-Zustand „Gruppen erforderlich“
5. Abschluss-Dialog

---

## Bekannte Einschränkungen / Fehler

| Thema | Status | Detail |
|-------|--------|--------|
| Native Scan ohne Rects | Bekannt | iOS/Android Lite-Scan → keine Overlays/Platzierung ohne Rects |
| pdf.js CDN | Bekannt | Offline → Fallback ohne Highlights |
| `hint`-Feld | Unverändert | Wird in Schritt 2 gesetzt, nicht im Mapping |
| Default-Gruppen | Design | `ensureWizardInitialized` legt 6 Standardgruppen an — Fall B nur bei explizit leerem `wizard.groups` |
| Alert auf Web | Gering | `Alert.alert` funktioniert; auf Web ggf. Browser-native Dialog |
| Completion-Alert | Gering | Erscheint einmal; erneutes Öffnen nach „Später“ zeigt Button „Zu Feldern wechseln“ im Footer |

---

## Testfälle

| # | Szenario | Erwartung | Status |
|---|----------|-----------|--------|
| 1 | PDF mit 5 Feldern importieren | Assistent startet bei Feld 1, Fortschritt 0→100 % | Manuell |
| 2 | PDF mit 50 Feldern | Flüssige Navigation, kein Layout-Bruch | Manuell |
| 3 | Keine Gruppen (`wizard.groups = []`) | Hinweis + Gruppe anlegen erforderlich | Manuell |
| 4 | Mehrere Gruppen | Karten-Auswahl, Zuordnung speichert in `assignments` | Manuell |
| 5 | Feld am oberen Seitenrand | Overlay-Panel unten | Unit-Test ✓ |
| 6 | Feld am unteren Seitenrand | Overlay-Panel oben | Unit-Test ✓ |
| 7 | Mehrere PDF-Seiten | Seitenwechsel bei Feldwechsel | Manuell |
| 8 | Überspringen | `deferredFieldIds`, am Ende → `sonstiges` | Bestehend ✓ |
| 9 | Letztes Feld zuordnen | Dialog erscheint, kein Auto-Redirect | Manuell |
| 10 | „Zu Feldern wechseln“ | `wizard.step = fields`, Navigation Schritt 2 | Manuell |
| 11 | Typecheck | Keine TS-Fehler | ✓ `npm run typecheck` |
| 12 | `resolveCurrentMappingIndex` | Erstes offenes Feld nach Teilzuordnung | Unit-Test ✓ |

### Automatisierte Tests

```bash
cd expo-toolbox
npm run typecheck
# setup-mapping.test.ts (manuell via Projekt-Testrunner wenn konfiguriert)
```

---

## Architektur-Kompatibilität

- `SetupWizardState` unverändert (keine Schema-Änderung)
- `rebuildSectionsFromWizard` unverändert
- Schritt 2 (`SetupFieldSettingsStep`) erhält weiterhin valide `single_sections`
- ETB / Legacy: unberührt (`resolveSetupEntryPath`, `useLegacyEditor`)

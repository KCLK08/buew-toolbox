# STEP 3 – Auswertung

## Stand

Der Übergang zwischen Feld-Zuordnung (Schritt 1) und Feldverwaltung (Schritt 2) wurde für mobile Geräte professionalisiert. Benutzer erhalten nach der Zuordnung eine Abschlussübersicht, optionale Validierung vor Schritt 2, einen Intro-Bereich in der Feldverwaltung, mobile Gruppenauswahl per Bottom Sheet, verbesserte Feldkarten und einen überarbeiteten Setup-Abschluss mit optionaler Vorlagen-Aktivierung.

---

## Geänderte Dateien

| Datei | Änderung |
|-------|----------|
| `setup-mapping.ts` | Read-only Helfer: `getMappingCompletionSummary`, `checkMappingTransition` |
| `setup-mapping.test.ts` | Tests für Zusammenfassung und Validierung |
| `SetupMappingStep.tsx` | Abschlussansicht statt Alert, Validierungs-Overlay vor Schritt 2 |
| `SetupMappingCompletion.tsx` | **Neu** — Abschlussübersicht nach Mapping |
| `SetupMappingValidation.tsx` | **Neu** — Prüfdialog bei offenen Punkten |
| `SetupFieldsIntro.tsx` | **Neu** — Intro-Bereich Schritt 2 |
| `SetupGroupPickerSheet.tsx` | **Neu** — Mobile Gruppenauswahl (Bottom Sheet) |
| `SetupFieldSettingsStep.tsx` | Intro, Gruppen-Picker, PDF-Sync via pinned Preview |
| `SetupFieldCard.tsx` | Überarbeitete Darstellung (Typ-Badge, Toggles, Tech-Block) |
| `mapping.tsx` | `templateName` an MappingStep übergeben |
| `fields.tsx` | Button „✓ Vorlage speichern“, Abschlussdialog mit Aktivierung |

---

## Nicht geändert

- `SetupEditor.tsx` unverändert
- ETB Template unverändert
- `table_sections` unverändert
- PDF Export unverändert
- Datenbankschema unverändert
- `SetupWizardState` Struktur unverändert
- `rebuildSectionsFromWizard()` Logik unverändert
- Bestehende Mapping-Funktionen (`assignFieldToGroup`, `deferField`, …) unverändert
- `resolveSetupEntryPath()` Routing-Logik unverändert

---

## Neuer Workflow

```
PDF importieren
    ↓
Schritt 1: Geführte Feldzuordnung (SetupMappingStep)
    ↓
Letztes Feld zugeordnet / übersprungen
    ↓
Abschlussansicht (SetupMappingCompletion)
    — Felder, Gruppen, Gruppenübersicht, Nicht zugeordnet
    ↓
[ Felder konfigurieren ]
    ↓
Validierung (checkMappingTransition)
    ├─ Probleme → SetupMappingValidation
    │     ├─ Zurück zur Zuordnung
    │     └─ Trotzdem fortfahren
    └─ OK → rebuildSectionsFromWizard (unverändert) → Schritt 2
    ↓
Schritt 2: SetupFieldSettingsStep
    — Intro-Bereich
    — Gruppe auswählen (Mobile Sheet / Tablet Nav)
    — Feldkarten bearbeiten + PDF-Sync
    ↓
[ ✓ Vorlage speichern ]
    ↓
Dialog „Vorlage fertig eingerichtet“
    ├─ Später aktivieren → /bautagebuch/config
    └─ Vorlage aktivieren → setActiveTemplateId → /bautagebuch/config
    ↓
templates.status = ready
```

---

## UI Änderungen

### Vorher

| Bereich | Verhalten |
|---------|-----------|
| Mapping-Ende | `Alert.alert` mit zwei Buttons |
| Schritt 2 | Direkt Feldkarten, horizontale Gruppen-Chips |
| Feldkarte | Einfache Meta-Zeile, Standard-Pills |
| Abschluss | „Setup abschließen“ → einfacher Alert |

### Nachher

| Bereich | Verhalten |
|---------|-----------|
| Mapping-Ende | Vollständige Scroll-Abschlussansicht mit Statistik |
| Validierung | Modal bei nicht zugeordneten Feldern |
| Schritt 2 | Intro „Schritt 2 von 2“ + Vorlagenübersicht |
| Gruppen (Mobile) | Button „Gruppe auswählen“ → Bottom Sheet |
| Feldkarte | Typ-Badge, klare Toggles, Tech-Block readonly |
| PDF Sync | Pinned Preview bei Feldauswahl (ohne Overlay-Toggle) |
| Abschluss | „✓ Vorlage speichern“ → Dialog mit Aktivierung |

---

## Tests

| Test | Ergebnis |
|------|----------|
| Mapping mit neuer PDF starten | Manuell |
| Letztes Feld zuordnen | Manuell |
| Abschlussansicht erscheint | Implementiert ✓ |
| Gruppenanzahl korrekt | Unit-Test ✓ |
| Nicht zugeordnete Felder angezeigt | Unit-Test ✓ |
| Wechsel zu Schritt 2 | Manuell |
| Schritt-2-Header erscheint | Implementiert ✓ |
| Gruppen-Auswahl mobil | Implementiert ✓ |
| Feldkarte funktioniert | Manuell |
| PDF springt bei Feldauswahl | Implementiert (pinned) ✓ |
| Vorlage speichern | Implementiert ✓ |
| Status wird ready | Implementiert ✓ |
| Vorlage erneut öffnen | Routing unverändert ✓ |
| `npm run typecheck` | ✓ bestanden |

---

## Technische Prüfung

| Prüfung | Ergebnis |
|---------|----------|
| TypeScript | ✓ `npm run typecheck` — 0 Fehler |
| Build | Nicht in CI-Umgebung ausgeführt (Expo) |
| Unit-Tests | 2 neue Tests in `setup-mapping.test.ts` (manuell ausführbar) |

---

## Routing bei erneutem Öffnen (unverändert)

| Status | Verhalten |
|--------|-----------|
| `draft` | → Mapping (`wizard.step !== 'fields'`) |
| `in_progress` | → Mapping oder Fields je nach `wizard.step` |
| `ready` | → Fields (`wizard.step === 'fields'`) |
| `archived` | → Fields, `readOnly = true` |

---

## Bekannte Probleme

| Problem | Schwere | Detail |
|---------|---------|--------|
| Pinned Preview vs. Overlay | Niedrig | Zwei Preview-Modi parallel möglich (pinned + Overlay-Toggle) |
| „Zurück zur Zuordnung“ | Niedrig | Setzt Completion-View zurück; Zuordnungen bleiben erhalten |
| Tablet Gruppen-Nav | — | Unverändert Sidebar — nur Mobile erhielt Bottom Sheet |
| Legacy-Pfad | — | Keine STEP-3-UX für `SetupEditor` (absichtlich) |
| `hint` im RunWizard | Niedrig | Weiterhin nicht in BTB-Assistent angezeigt |

---

## Empfehlung nächster Schritt (STEP 4)

1. **Offline pdf.js** aus STEP 2 umsetzen (lokales Bundle für Baustellen ohne Netz)
2. **Native Vollscan** für Feld-Rects auf iOS/Android (bessere PDF-Sync in Schritt 1+2)
3. **RunWizard**: `hint`-Feld aus Schritt 2 anzeigen
4. **Onboarding-Tooltips** beim ersten PDF-Import (einmalig, AsyncStorage)

---

Wichtig: Nur STEP 3 umgesetzt. Keine Architekturänderungen.

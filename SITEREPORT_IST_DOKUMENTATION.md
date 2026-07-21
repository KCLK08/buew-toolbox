# SiteReport — Technische IST-Dokumentation (Expo APK)

> **Stand:** Juli 2026 (nach PWA-Paritäts-Refactoring)  
> **Zweck:** Vollständige Beschreibung des SiteReport-Moduls in der **BÜW-Toolbox Expo-App** für Weiterentwicklung ohne Quellcode-Zugriff.  
> **Scope:** Ausschließlich die native Android-APK (`expo-toolbox/`). Keine PWA, kein WebView, keine SvelteKit-Referenz.

---

## Inhaltsverzeichnis

1. [Projektübersicht](#1-projektübersicht)
2. [UI/UX](#2-uiux)
3. [Navigation](#3-navigation)
4. [Funktionen](#4-funktionen)
5. [Datenmodell](#5-datenmodell)
6. [Datenhaltung](#6-datenhaltung)
7. [Einstellungen](#7-einstellungen)
8. [Komponenten](#8-komponenten)
9. [Custom Hooks & Context](#9-custom-hooks--context)
10. [State Management](#10-state-management)
11. [Services](#11-services)
12. [Sicherheit](#12-sicherheit)
13. [Bekannte Probleme](#13-bekannte-probleme)
14. [Architektur](#14-architektur)
15. [Verbesserungspotential](#15-verbesserungspotential)
16. [Zusammenfassung](#16-zusammenfassung)

---

## 1. Projektübersicht

### 1.1 Zweck der Anwendung

**SiteReport** ist ein **offline-fähiges Baustellen-Protokoll-Tool** innerhalb der **BÜW-Toolbox Expo-App**. Nutzer erstellen Foto-basierte Protokolleinträge auf der Baustelle, erfassen strukturierte Zusatzdaten über einen **geführten Wizard** und exportieren das Ergebnis als **Excel (XLSX)** und/oder **PDF** — optional mit **Firmenlogo** im Export.

Kernworkflow:

1. Dashboard öffnen → „Neues Protokoll" (3-Schritt-Setup)
2. Tabellenformat wählen oder anlegen, Stammdaten + Logo erfassen
3. Protokoll-Detail → Einträge über **Guided Entry Wizard** (Foto → Felder → Zusammenfassung)
4. Protokoll abschließen (Speichern / PDF / Excel / beides) oder Export aus Export-Center

Die App wird als **standalone Release-APK** verteilt (eingebettetes JS-Bundle, kein Metro-Server nötig).

### 1.2 Zielgruppe

- Bauleitung, Polier, Projektleitung auf Baustellen
- Mitarbeitende, die **Foto-Protokolle** mit tabellarischer Struktur führen
- Nutzer der **BÜW-Toolbox** (SiteReport ist eines von zwei Werkzeugen neben Bautagebuch)

### 1.3 Aktueller Entwicklungsstand

| Bereich | Status |
|---------|--------|
| Dashboard mit Statistiken | ✅ |
| Geführter Protokoll-Start (3 Schritte) | ✅ |
| Guided Entry Wizard | ✅ |
| Format-Builder (Spaltenvorlagen) | ✅ inkl. Up/Down-Reorder |
| Firmenlogo | ✅ |
| Stammdaten bearbeiten | ✅ |
| PDF-Export | ✅ |
| XLSX-Export | ✅ |
| Export-Cache (Dateipfade + Re-Share) | ✅ |
| Export-Center (eigener Screen) | ✅ |
| Protokoll-Liste mit Bulk-Auswahl | ✅ |
| Protokoll löschen (Einzel + Bulk) | ✅ |
| Entry löschen | ✅ |
| Protokoll abschließen (Modal) | ✅ |
| Foto-Kompression + permanente Speicherung | ✅ |
| Toast-Feedback | ✅ |
| Bottom Sheets / Confirm Modals | ✅ |
| Offline-Betrieb | ✅ |
| Multi-DB-Backup (Toolbox-Ebene) | ✅ |
| Dark Mode | ⚠️ Tokens vorbereitet, nicht aktiv |
| Cloud-Sync | ❌ |

**Fazit:** Die Expo-App hat **funktionale Parität zur PWA** mit **nativer UX** (Karten, Wizard, Bulk, Modals). Verbleibende Lücken: Dark Mode, Cloud-Sync, Schema-Migrationen.

### 1.4 Verwendete Technologien

| Technologie | Version (ca.) | Verwendung |
|-------------|---------------|------------|
| Expo SDK | ~57 | App-Framework |
| React Native | 0.86.x | UI |
| React | 19.2.x | UI |
| TypeScript | 5.9.x | Typsicherheit |
| Expo Router | ~57 | File-based Navigation |
| expo-sqlite | ~57 | SQLite-Datenbank |
| expo-image-picker | ~57 | Kamera / Galerie |
| expo-image-manipulator | ~57 | Foto-Kompression (1600px, 0.75) |
| expo-file-system | ~57 | Dateispeicher, Exporte, Fotos |
| expo-sharing | ~57 | System-Share-Sheet |
| ExcelJS | 4.4.x | XLSX-Generierung |
| pdf-lib | 1.17.x | PDF-Generierung |
| Space Grotesk | Google Fonts | Typografie |

**Nicht verwendet in SiteReport:** Redux, Zustand, React Query, MobX, WebViews, Drawer-Navigation, Cloud-Sync.

### 1.5 APK-Build & Verteilung

| Aspekt | Detail |
|--------|--------|
| Projekt | `expo-toolbox/` |
| CI-Workflow | `.github/workflows/android-apk.yml` |
| Trigger | Push auf `main` bei Änderungen in `expo-toolbox/**` oder `shared/**` |
| Build | `npx expo prebuild --platform android` → `./gradlew assembleRelease` |
| Release-Asset | `buew-toolbox-<version>-<sha>.apk` |
| GitHub-Tag | `apk-v<version>.<run_number>` |
| PR-Checks | Nur Typecheck, **keine APK** |

Details: [`ANDROID_APK.md`](ANDROID_APK.md)

### 1.6 Projektstruktur

```
/workspace/expo-toolbox/
├── app/
│   ├── _layout.tsx                          # Root Stack, ToastProvider, DB-Init
│   ├── (tabs)/
│   │   ├── index.tsx                          # Toolbox-Home
│   │   └── sitereport/index.tsx               # SiteReport-Dashboard (Tab)
│   └── sitereport/
│       ├── new-protocol.tsx                   # 3-Schritt Protokoll-Setup
│       ├── format-builder.tsx                 # Format-Editor
│       ├── protocols/index.tsx                # Protokoll-Liste (Bulk)
│       ├── exports/index.tsx                  # Export-Center
│       └── protocol/
│           ├── [id].tsx                       # Protokoll-Detail
│           ├── [id]/edit.tsx                  # Stammdaten bearbeiten
│           └── [id]/wizard.tsx                # Guided Entry Wizard
├── src/
│   ├── components/
│   │   ├── mobile/                            # Shared UI (Screen, Button, …)
│   │   └── sitereport/                        # SiteReport-spezifische Komponenten
│   ├── contexts/ToastContext.tsx              # App-weite Toast-Meldungen
│   ├── constants/theme.ts                     # Design Tokens
│   ├── constants/tools.ts                     # Tab-Definition
│   ├── hooks/useOfflineBootstrap.ts           # Backup/Restore
│   ├── storage/backupService.ts               # Multi-DB-Backup
│   └── native/sitereport/
│       ├── db/database.ts                     # SQLite + Typen
│       ├── services/
│       │   ├── exportService.ts               # Export + Cache + Abschluss
│       │   ├── photoService.ts                # Kompression + Persistenz
│       │   └── protocolService.ts             # Entry/Protocol-Helfer, Bulk
│       └── lib/
│           ├── pdf.js
│           ├── xlsx-export.js
│           └── native-image.ts
├── app.config.js
├── eas.json
└── package.json
```

---

## 2. UI/UX

### 2.1 Design Tokens

**Datei:** `expo-toolbox/src/constants/theme.ts`

| Token | Wert | Verwendung |
|-------|------|------------|
| `colors.bg` | `#F2F0EB` | Hintergrund |
| `colors.ink` | `#1A1916` | Primärtext |
| `colors.muted` | `#6B6560` | Sekundärtext |
| `colors.accent` | `#C44B32` | Primär-Aktion, Tab aktiv |
| `colors.panel` | `#FFFCF7` | Karten, Header |
| `colors.border` | `#E4DDD2` | Rahmen |
| `colors.danger` | `#A12C24` | Fehler, Löschen |
| Font | Space Grotesk 400/600/700 | Gesamte App |

**Stil:** Professionell, minimalistisch, baustellen-tauglich, große Touchflächen, Karten-Layout.

---

### 2.2 Dashboard

**Datei:** `app/(tabs)/sitereport/index.tsx`  
**Route:** `/(tabs)/sitereport`

#### Aufbau

1. **StatCards** — Protokolle, Einträge, Fotos, Exporte (2×2 Grid)
2. **DashboardActionCards** — „Neues Protokoll" (Akzent), „Protokolle"
3. **Aktive Protokolle** — letzte 3 als `ListItem`
4. **Letzte Exporte** — letzte 3 als `ListItem`
5. **FAB** `+` → Neues Protokoll

#### Interaktion

- Pull-to-Refresh
- Navigation zu `/sitereport/new-protocol`, `/sitereport/protocols`, `/sitereport/exports`

---

### 2.3 Neues Protokoll (3-Schritt-Setup)

**Datei:** `app/sitereport/new-protocol.tsx`  
**Route:** `/sitereport/new-protocol`

| Schritt | Inhalt |
|---------|--------|
| 1 — Allgemeine Daten | Protokollname*, Projekt, Beschreibung, Teilnehmer, Datum (readonly) |
| 2 — Firmenlogo | Auswählen, Vorschau, ändern, entfernen |
| 3 — Tabellenformat | Template wählen, neu/bearbeiten, `TablePreview` |

**Abschluss:** „Protokoll erstellen" → `createProtocol()` → Navigate to Detail

**UI:** `WizardStep` mit Fortschrittsbalken, Zurück/Weiter-Footer

---

### 2.4 Protokoll-Detail

**Datei:** `app/sitereport/protocol/[id].tsx`  
**Route:** `/sitereport/protocol/:id`

#### Aufbau

1. **Summary StatCards** — Datum, Einträge, Fotos, Offen
2. **Aktionen** — Stammdaten bearbeiten, Export (Bottom Sheet), Protokoll abschließen
3. **EntryCards** — Foto groß oben, Status-Badge, Felder, Bearbeiten/Löschen
4. **Footer** — „+ Neuer Eintrag" → Wizard

#### Abschluss-Modal (`ConfirmModal`)

| Option | Aktion |
|--------|--------|
| Nur speichern | `updateProtocol()` → Protokoll-Liste |
| PDF erstellen | `closeProtocolWithExport('pdf')` |
| Excel erstellen | `closeProtocolWithExport('xlsx')` |
| PDF + Excel | `closeProtocolWithExport('both')` |

Export ohne Share-Sheet (`share: false`), Toast-Feedback.

#### Export Bottom Sheet

- PDF exportieren (mit Share-Sheet)
- Excel exportieren (mit Share-Sheet)

---

### 2.5 Stammdaten bearbeiten

**Datei:** `app/sitereport/protocol/[id]/edit.tsx`  
**Route:** `/sitereport/protocol/:id/edit`

Editierbar: Titel, Projekt, Beschreibung, Teilnehmer, Logo.  
Datum bleibt readonly (aus Protokoll-Record).

---

### 2.6 Guided Entry Wizard

**Datei:** `app/sitereport/protocol/[id]/wizard.tsx`  
**Route:** `/sitereport/protocol/:id/wizard?entryId=<id>` (optional für Bearbeiten)

#### Ablauf

1. Pro Spalte ein Screen (`WizardStep` mit `1/N` Anzeige)
2. **Foto-Spalte:** `PhotoCaptureStep` — große Kamera-Fläche, Vorschau, entfernen
3. **Status-Spalte:** `StatusFieldStep` — Offen / Bearbeitung / Erledigt (farbige Buttons)
4. **Andere Spalten:** `FieldStep` — ein Feld pro Screen
5. **Zusammenfassung:** `ProtocolSummary` → Speichern

**Foto:** `captureProtocolPhoto()` → Kompression → `sitereport/photos/{protocolId}/{entryId}.jpg`

**Query:** `entryId` gesetzt = Bearbeiten, sonst neuer Eintrag

---

### 2.7 Protokoll-Liste

**Datei:** `app/sitereport/protocols/index.tsx`  
**Route:** `/sitereport/protocols`

- **ProtocolCards** mit Titel, Projekt, Datum, Eintrags-/Fotoanzahl
- Einzel-Löschen (✕ + Alert)
- **Bulk-Modus:** Checkboxen, Alle auswählen, Bulk-Export (Bottom Sheet), Bulk-Löschen

---

### 2.8 Export-Center

**Datei:** `app/sitereport/exports/index.tsx`  
**Route:** `/sitereport/exports`

- **ExportCards** mit PDF/Excel-Badges
- PDF/Excel teilen, Löschen
- Bulk-Modus mit Mehrfach-Löschen

---

### 2.9 Format-Builder

**Datei:** `app/sitereport/format-builder.tsx`  
**Route:** `/sitereport/format-builder?mode=new|edit&templateId=...`

- `FormatColumnCard` pro Spalte
- Spalte hinzufügen/bearbeiten/entfernen, ↑/↓ Reorder
- Foto-Spalte fest, nicht löschbar
- Toast bei Speichern

---

### 2.10 Globale UI-Elemente

| Element | Verwendung |
|---------|------------|
| `ToastProvider` | „Gespeichert", „Exportiert", „Gelöscht" (2,2s, animiert) |
| `ConfirmModal` | Abschluss-Flow, kritische Bestätigungen |
| `BottomSheet` | Export-Auswahl, Bulk-Export |
| `Alert.alert` | Löschen-Bestätigung, Kamera-Berechtigung, Fehler |

---

## 3. Navigation

### 3.1 Expo Router Struktur

```
Root Stack (_layout.tsx) + ToastProvider
├── (tabs)
│   ├── index                    # Toolbox-Home
│   ├── sitereport/index         # Dashboard (Tab)
│   └── bautagebuch/
├── sitereport/new-protocol
├── sitereport/protocol/[id]
├── sitereport/protocol/[id]/edit
├── sitereport/protocol/[id]/wizard
├── sitereport/protocols/index
├── sitereport/exports/index
└── sitereport/format-builder
```

### 3.2 Screen-Hierarchie

```
Dashboard (Tab)
├── Neues Protokoll (3 Schritte)
│   └── → Protokoll-Detail
│       ├── Stammdaten bearbeiten
│       ├── Entry Wizard (neu/bearbeiten)
│       ├── Export (Bottom Sheet)
│       └── Abschließen (Modal) → Protokoll-Liste
├── Protokolle (Bulk)
├── Export-Center
└── Format-Builder
```

### 3.3 Deep Links

- `sitereport/new-protocol`
- `sitereport/protocol/<id>`
- `sitereport/protocol/<id>/edit`
- `sitereport/protocol/<id>/wizard?entryId=<id>`
- `sitereport/protocols`
- `sitereport/exports`
- `sitereport/format-builder?mode=new|edit&templateId=<id>`

---

## 4. Funktionen

### 4.1 Firmenlogo

| Aspekt | Detail |
|--------|--------|
| Speicher | Data-URL in SQLite `settings.id='logo'` |
| Eingabe | Galerie-Picker in Setup (Schritt 2) oder Stammdaten-Edit |
| Ausgabe | In PDF/XLSX Header eingebettet |

### 4.2 Tabellenformat

Wie zuvor: Templates in SQLite, Seed „Standard Baustelle", ↑/↓ Reorder, `FormatColumnCard`.

### 4.3 Protokoll erstellen

**Screen:** `new-protocol.tsx` (3 Schritte)  
**Pflicht:** Protokollname + Template  
**Ablauf:** `createProtocol()` mit vollständigen Stammdaten → Detail-Screen

### 4.4 Stammdaten bearbeiten

**Screen:** `protocol/[id]/edit.tsx`  
Editierbar: Titel, Projekt, Beschreibung, Teilnehmer, Logo. Datum readonly.

### 4.5 Eintrag hinzufügen (Guided Wizard)

| Aspekt | Detail |
|--------|--------|
| Ablauf | Wizard: Foto → je Spalte ein Screen → Zusammenfassung |
| Foto | `expo-image-manipulator` (1600px, JPEG 0.75) |
| Speicherort | `{documentDirectory}sitereport/photos/{protocolId}/{entryId}.jpg` |
| Status | `offen` / `bearbeitung` / `erledigt` |
| Reihenfolge | Neueste oben (prepend) |

### 4.6 Eintrag bearbeiten / löschen

| Aktion | Implementierung |
|--------|-----------------|
| Bearbeiten | Wizard mit `?entryId=` |
| Löschen | `removeProtocolEntry()` + Foto-Datei löschen, Alert-Bestätigung |

### 4.7 Protokoll abschließen

`closeProtocolWithExport(protocol, mode)` — Speichern, PDF, XLSX oder beides ohne Share-Sheet, dann Navigation zur Protokoll-Liste.

### 4.8 PDF / XLSX Export

Unverändert technisch (`pdf.js`, `xlsx-export.js`).  
**Neu:** `exportProtocolPdf/Xlsx(protocol, { share: true|false })`.

### 4.9 Export-Cache

Eigener Screen `/sitereport/exports`. `shareCachedExport`, `deleteCachedExport`, Bulk-Löschen.

### 4.10 Protokoll-Verwaltung

| Aktion | API |
|--------|-----|
| Einzel löschen | `deleteProtocolWithCleanup()` |
| Bulk löschen | `bulkDeleteProtocols()` |
| Bulk exportieren | `exportProtocolPdf/Xlsx` mit `share: false` |

### 4.11 Bulk-Auswahl

Protokolle und Exporte: Selection-Mode mit Checkboxen, Alle auswählen, Bulk-Aktionen.

---

## 5. Datenmodell

Unverändert — SQLite-Schema identisch, keine Migration nötig.

### 5.1 Kern-Typen

- `SiteReportColumn`, `SiteReportEntry`, `SiteReportProtocol`, `SiteReportTemplate`, `SiteReportSettings`, `SiteReportExport`

### 5.2 `photoPath`

Permanenter Pfad unter `sitereport/photos/{protocolId}/{entryId}.jpg` (nicht mehr Kamera-Cache-URI).

---

## 6. Datenhaltung

### 6.1 SQLite

` sitereport_native.db` — unverändert (protocols, settings, templates, exports).

### 6.2 Dateispeicherung

| Pfad | Inhalt |
|------|--------|
| `sitereport/exports/*.pdf` | PDF-Exporte |
| `sitereport/exports/*.xlsx` | XLSX-Exporte |
| `sitereport/photos/{protocolId}/*.jpg` | **Permanente Eintrags-Fotos** |

### 6.3 Foto-Pipeline

```
launchCameraAsync
  → ImageManipulator (resize 1600px, compress 0.75, JPEG)
  → copyAsync → sitereport/photos/{protocolId}/{entryId}.jpg
  → photoPath in entry
```

Beim Löschen: `deleteEntryPhoto()`. Beim Protokoll-Löschen: `deleteProtocolPhotos(protocolId)`.

### 6.4 Backup

Unverändert — `backupService.ts` sichert `sitereport_native.db` mit Toolbox-DBs.

---

## 7. Einstellungen

Unverändert: Template (`settings.current`), Logo (`settings.logo`). Kein Dark Mode aktiv.

---

## 8. Komponenten

### 8.1 SiteReport (`src/components/sitereport/`)

| Komponente | Zweck |
|------------|-------|
| `ProtocolCard` | Protokoll-Karte mit Bulk-Checkbox |
| `EntryCard` | Eintrag mit Foto, Status-Badge, Aktionen |
| `ExportCard` | Export-Karte mit PDF/Excel-Badges |
| `FormatColumnCard` | Spalten-Editor im Format-Builder |
| `WizardStep` | Schritt-Anzeige mit Fortschrittsbalken |
| `PhotoCaptureStep` | Kamera-UI im Wizard |
| `FieldStep` | Einzelfeld im Wizard |
| `StatusFieldStep` | Status-Auswahl (3 Optionen) |
| `ProtocolSummary` | Zusammenfassung vor Speichern |
| `TablePreview` | Tabellenvorschau im Setup |
| `DashboardActionCard` | Dashboard-Schnellaktionen |
| `ConfirmModal` | Native Modal für Bestätigungen |
| `BottomSheet` | Slide-up Sheet für Export-Auswahl |

### 8.2 Shared Mobile (`src/components/mobile/`)

`Screen`, `PrimaryButton`, `TextField`, `ListItem`, `Card`, `StatCard`, `StatusBadge`, `EmptyState`, `Fab`

---

## 9. Custom Hooks & Context

### `useToast()` — `src/contexts/ToastContext.tsx`

| API | Beschreibung |
|-----|--------------|
| `showToast(message)` | Kurze Meldung unten, auto-hide |

Eingebunden in `app/_layout.tsx` via `ToastProvider`.

### `useOfflineBootstrap()` — App-weit

Backup/Restore-Banner (unverändert).

**Keine SiteReport-spezifischen Hooks** — State in Screen-Komponenten.

---

## 10. State Management

| Ansatz | Verwendung |
|--------|------------|
| React `useState` | Alle Screens |
| `useCallback` / `useEffect` | Laden, Refresh |
| `ToastContext` | Globales Feedback |
| Kein Redux/Zustand/React Query | — |

---

## 11. Services

### 11.1 `database.ts`

Unverändert: CRUD für protocols, templates, settings, exports, `softDeleteProtocol`, `todayDe()`.

### 11.2 `photoService.ts`

| Funktion | Beschreibung |
|----------|--------------|
| `compressAndPersistPhoto` | Resize + JPEG + copy nach photos/ |
| `captureProtocolPhoto` | Kamera → persistieren |
| `deleteEntryPhoto` | Einzelnes Foto löschen |
| `deleteProtocolPhotos` | Gesamten Protokoll-Ordner löschen |

### 11.3 `protocolService.ts`

| Funktion | Beschreibung |
|----------|--------------|
| `createEntryId` | ID-Generierung |
| `emptyFieldsFromColumns` | Default-Felder inkl. Status `offen` |
| `protocolStats` | entryCount, photoCount, openCount |
| `removeProtocolEntry` | Entry + Foto löschen |
| `deleteProtocolWithCleanup` | Protokoll + Exporte + Fotos |
| `bulkDeleteProtocols` | Mehrere Protokolle |
| `upsertProtocolEntry` | Entry erstellen/aktualisieren |
| `getProtocolOrThrow` | Laden mit Fehler |

### 11.4 `exportService.ts`

| Funktion | Beschreibung |
|----------|--------------|
| `exportProtocolPdf(protocol, { share? })` | PDF generieren, optional teilen |
| `exportProtocolXlsx(protocol, { share? })` | XLSX generieren, optional teilen |
| `closeProtocolWithExport(protocol, mode)` | Abschluss: save/pdf/xlsx/both |
| `shareCachedExport` | Aus Cache teilen oder regenerieren |
| `deleteCachedExport` | Dateien + DB-Eintrag löschen |
| `listExports` | Re-export aus database |

---

## 12. Sicherheit

Unverändert: Kamera- und Galerie-Berechtigungen, keine DB-Verschlüsselung, lokale Datenhaltung.

---

## 13. Bekannte Probleme

| ID | Beschreibung | Schwere |
|----|--------------|---------|
| B1 | Kein Schema-Migrations-Framework | Mittel |
| B2 | Logo als große Data-URL in SQLite | Niedrig |
| B3 | Spalten umbenennen bricht `fields`-Keys in alten Einträgen | Mittel |
| B4 | `entriesJson` bei jedem Feld-Update voll serialisiert | Niedrig |
| B5 | Dark Mode Tokens vorhanden, nicht aktiv | Niedrig |
| B6 | Keine Cloud-Sync | Info |

**Behoben seit vorheriger Doku:** Foto-Persistenz, Wizard, Bulk, Protokoll-Löschen, Stammdaten-Edit, Abschluss-Flow.

---

## 14. Architektur

### 14.1 Schichten

| Schicht | Verantwortung |
|---------|---------------|
| Screens (`app/sitereport/`) | UI, lokaler State, Navigation |
| `components/sitereport/` | Wiederverwendbare SiteReport-UI |
| `protocolService.ts` | Business-Logik Entry/Protocol |
| `photoService.ts` | Foto-Pipeline |
| `exportService.ts` | Export-Pipeline |
| `database.ts` | Persistenz |

### 14.2 Diagramm

```mermaid
flowchart TB
  subgraph UI [Expo Screens]
    Dash[Dashboard]
    New[New Protocol]
    Detail[Protocol Detail]
    Wizard[Entry Wizard]
    Lists[Protocols / Exports]
  end

  subgraph Components [components/sitereport]
    Cards[ProtocolCard / EntryCard / ExportCard]
    WizardUI[WizardStep / PhotoCaptureStep]
    Modals[ConfirmModal / BottomSheet]
  end

  subgraph Services
    PS[protocolService]
    PhS[photoService]
    ES[exportService]
  end

  subgraph Data
    DB[(sitereport_native.db)]
    Photos[sitereport/photos/]
    Exports[sitereport/exports/]
  end

  Dash --> New --> Detail --> Wizard
  Detail --> Lists
  UI --> Components
  Wizard --> PhS --> Photos
  Detail --> PS --> DB
  ES --> Exports --> DB
  Dash --> ES
```

---

## 15. Verbesserungspotential

- **Schema-Migrationen** (`PRAGMA user_version`)
- **Dark Mode** aktivieren
- **Repository-Pattern** zwischen Screens und DB
- **Spalten-ID statt Name** als Field-Key (Umbenennung-sicher)
- **On-Device-Tests** der Export-Pipeline auf Release-APK
- Optional: Haptic Feedback, Swipe-to-Delete

---

## 16. Zusammenfassung

### Was funktioniert

| Feature | Status |
|---------|:------:|
| Dashboard mit Statistiken | ✅ |
| Geführter Protokoll-Start | ✅ |
| Guided Entry Wizard | ✅ |
| Stammdaten bearbeiten | ✅ |
| Protokoll abschließen | ✅ |
| Entry löschen | ✅ |
| Protokoll löschen (Einzel + Bulk) | ✅ |
| Bulk-Export | ✅ |
| Export-Center | ✅ |
| Foto-Kompression + Persistenz | ✅ |
| Toast / Modals / Bottom Sheets | ✅ |
| PDF/XLSX Export + Cache | ✅ |
| Offline + Backup | ✅ |

### Was fehlt noch

- Dark Mode
- Cloud-Sync
- SQLite-Migrations-Framework
- Template löschen (API fehlt)

### Nächste Schritte

1. PR #25 mergen → APK-Build auf `main`
2. On-Device-Test: Wizard → Export → Abschluss
3. Schema-Versionierung für zukünftige Updates

---

## Anhang A: Standard-Spalten

```typescript
[
  { id: 'col_photo', name: 'Bilder', type: 'text', isPhoto: true },
  { id: 'col_km', name: 'Kilometer', type: 'number', isPhoto: false },
  { id: 'col_desc', name: 'Beschreibung', type: 'text', isPhoto: false },
  { id: 'col_status', name: 'Status', type: 'text', isPhoto: false }
]
```

## Anhang B: ID-Generierung

| Entität | Format |
|---------|--------|
| Protocol | `protocol_{timestamp}_{random6}` |
| Template | `tpl_{timestamp}_{random6}` |
| Entry | `entry_{timestamp}_{random6}` |
| Column | `col_{timestamp}_{random6}` |
| Export | `export_{protocolId}` |

## Anhang C: Dateipfade

| Pfad (relativ zu documentDirectory) | Inhalt |
|-------------------------------------|--------|
| `SQLite/sitereport_native.db` | Hauptdatenbank |
| `sitereport/exports/` | PDF/XLSX |
| `sitereport/photos/{protocolId}/` | Eintrags-Fotos |
| `backups/sitereport_native_backup_{stamp}.db` | DB-Backup |

## Anhang D: Screen-Route-Übersicht

| Screen | Datei | Route |
|--------|-------|-------|
| Dashboard | `(tabs)/sitereport/index.tsx` | `/(tabs)/sitereport` |
| Neues Protokoll | `sitereport/new-protocol.tsx` | `/sitereport/new-protocol` |
| Protokoll-Detail | `sitereport/protocol/[id].tsx` | `/sitereport/protocol/:id` |
| Stammdaten | `sitereport/protocol/[id]/edit.tsx` | `/sitereport/protocol/:id/edit` |
| Entry Wizard | `sitereport/protocol/[id]/wizard.tsx` | `/sitereport/protocol/:id/wizard` |
| Protokolle | `sitereport/protocols/index.tsx` | `/sitereport/protocols` |
| Exporte | `sitereport/exports/index.tsx` | `/sitereport/exports` |
| Format-Builder | `sitereport/format-builder.tsx` | `/sitereport/format-builder` |

---

*Ende der IST-Dokumentation (Expo APK)*

# Bautagebuch Expo — PWA Gegenüberstellung (Phase 1)

> **Stand:** Juli 2026 · Branch `cursor/bautagebuch-expo-parity-fe37`

## Technische Gegenüberstellung

| Bereich | PWA Status | Expo Status (vor Parität) | Aufgabe / Umsetzung |
|---------|------------|---------------------------|---------------------|
| **Template-System** | Builtin `Vorlage-eBTB.pdf`, Setup v6, Dexie | Download + `scanTemplatePdfLite`, SQLite | ✅ Beibehalten; Lite-Scan bleibt (kein pdfjs) |
| **PDF Scan** | pdf-lib + pdfjs, rects, options, Canvas | Nur pdf-lib, keine rects/options | ⚠️ Teilweise; Dropdown-Matching über Setup-Optionen |
| **Setup Editor** | Canvas-Feld-Highlights | Text-Editor, Autosave 420 ms | ✅ Parität inhaltlich; visuelles PDF-Highlight offen |
| **Run Wizard** | 6 Sektionen, Autosave 450 ms, Pflichtfelder | Wizard ohne Export-Sperre, sofortiges Speichern | ✅ Pflichtfeld-Logik, Autosave 450 ms, Stepper |
| **Tabellen** | Spezial-UI Personal/Leistung | Generische Tabellen | ✅ Zeilenzähler, Uhrzeit, Multiline Leistungsblock |
| **Wetter** | Open-Meteo + Option-Matching | Open-Meteo, hardcoded Labels | ✅ `pickWeatherDropdownOption` portiert |
| **Fotos** | Galerie + Kamera, Cleanup | Nur Kamera, kein Cleanup | ✅ Galerie, Löschen inkl. Datei/DB |
| **PDF Export** | 3 Modi, Pflichtfeld-Sperre | 3 Modi, keine Sperre | ✅ Export blockiert bei fehlenden Pflichtfeldern |
| **PDF Vorschau** | Live Canvas während Eingabe | Nur Share-Preview | ✅ Live WebView-Vorschau (debounced, offline Base64) |
| **Backup** | IndexedDB-Snapshots, Event-Trigger | Toolbox SQLite-Backup bei App-Background | ✅ Backup bei Foto add/delete via `requestDatabaseBackup` |
| **Datenbank** | Dexie v1–v4, `deleteRunCascade` | SQLite, `softDeleteRun` ungenutzt | ✅ `deleteRunCascade`, `renameRun`, Photo-Sync |
| **Navigation** | Single-Page Views | Expo Router Tab + Stack | ✅ Home Cards, KW-Collapse, Auswahl-Modus |
| **Home CRUD** | Löschen, Umbenennen, Mehrfach-Löschen | Nur Öffnen/Erstellen | ✅ Swipe + Auswahl + Rename-Modal |
| **Mobile UX** | Browser-Layout | Native Tokens, FAB | ✅ Swipe, Toast, große Touchflächen, Status-Badges |

## Bewusst nicht umgesetzt (Scope)

- **WebView-PWA-Einbettung** — ausgeschlossen
- **Vollständiger pdfjs-Scan** — hoher Aufwand; separates Ticket
- **BTB_Backup.zip** — kein ZIP-Library im Projekt; Toolbox-DB-Backup bleibt
- **Komplette Ordner-Umstruktur** (`src/app/home.tsx` …) — unnötig; bestehende Expo-Router-Struktur erweitert
- **Dark Mode aktiv** — Tokens vorhanden, Aktivierung separat
- **Signatur-Erfassung** — in PWA ebenfalls übersprungen

## Neue / geänderte Dateien (Parität)

- `lib/run-validation.ts`, `lib/run-defaults.ts`
- `hooks/useRunAutosave.ts`
- `components/PdfPreviewPanel.tsx`, `components/BautagebuchRunCard.tsx`
- Erweitert: `RunWizard.tsx`, `database.ts`, `photoDocService.ts`, `weatherService.ts`, `exportService.ts`
- Screens: `(tabs)/bautagebuch/index.tsx`, `bautagebuch/run/[id].tsx`

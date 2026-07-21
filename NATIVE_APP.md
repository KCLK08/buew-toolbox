# Native App (ohne WebView)

Die Expo-App ist eine **eigenständige native Anwendung**, die dieselben Werkzeuge wie die PWA bereitstellt — implementiert mit React Native, SQLite und nativen APIs.

## Architektur

```
expo-toolbox/
├── app/                          # Expo Router Screens
│   ├── (tabs)/index.tsx          # Toolbox-Home
│   ├── (tabs)/bautagebuch/       # BTB-Übersicht
│   ├── (tabs)/sitereport/        # SiteReport-Übersicht
│   ├── bautagebuch/run/[id].tsx  # BTB Guided Flow
│   └── sitereport/protocol/[id].tsx
└── src/native/
    ├── bautagebuch/              # Port der Bautagebuch-PWA
    │   ├── lib/                  # setup-model, etb-template, pdf-export (aus PWA)
    │   ├── db/                   # SQLite
    │   └── components/RunWizard.tsx
    └── sitereport/               # Port der SiteReport-PWA
        └── db/
```

## Bautagebuch (nativ)

| PWA-Funktion | Status |
|--------------|--------|
| Vorlage-eBTB laden | ✅ (Download beim ersten Start) |
| BTB-Liste / neu erstellen | ✅ |
| Guided Flow (Kopfdaten, Wetter, Tabellen, Abschluss) | ✅ |
| Wetter-Sync (Open-Meteo + Standort) | ✅ |
| PDF-Export (pdf-lib) | ✅ |
| Fotodokumentation + PDF-Merge | 🔜 |
| Setup-Editor / Live-PDF-Vorschau | 🔜 |

## SiteReport (nativ)

| PWA-Funktion | Status |
|--------------|--------|
| Protokoll erstellen / Liste | ✅ |
| Foto-Einträge erfassen | ✅ |
| Format-Builder / Templates | 🔜 |
| XLSX-Export | 🔜 |
| PDF-Export | 🔜 |

## Keine WebViews

Die App lädt **keine** eingebetteten PWA-URLs. Daten liegen in separaten SQLite-Datenbanken auf dem Gerät (`bautagebuch_v2_native.db`, `sitereport_native.db`).

## Nächste Schritte

1. SiteReport: XLSX/PDF-Export (exceljs, pdf-lib) portieren
2. Bautagebuch: Fotodokumentation + Merge-Export
3. Format-Builder und Setup-Editor
4. Backup/Restore analog PWA

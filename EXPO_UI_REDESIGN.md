# Expo App — PWA-Shell (Korrektur)

Die Expo-App ist **keine separate native Daten-App**, sondern eine native Hülle um die bestehenden PWAs.

## Navigation

```
Bottom Tabs
├── Home        → Toolbox-Übersicht (SiteReport + Bautagebuch Karten)
├── SiteReport  → WebView: /sitereport/
└── Bautagebuch → WebView: /bautagebuch/
```

## Prinzip

| Aspekt | Umsetzung |
|--------|-----------|
| Funktionalität | 1:1 die PWA (SiteReport, Bautagebuch) |
| Daten | IndexedDB in der eingebetteten PWA (WebView) |
| Native Anteil | Tab-Bar, Safe Area, Ladeanzeige, Zurück in der PWA |
| URL-Basis | `EXPO_PUBLIC_TOOLBOX_WEB_BASE_URL` |

## Entfernt (fälschlich eingeführt)

Die vorherige „Native UI“-Phase hatte eigene Screens für Projekte, Tagebuch und Mängel gebaut. Diese **ersetzten** die PWA-Funktionen nicht und wurden entfernt:

- `app/(tabs)/projects.tsx`, `diary.tsx`, `defects.tsx`, `more.tsx`
- `app/project/*`, `app/diary/*`, `app/defect/*`

Repository-/SQLite-Code bleibt im Projekt (Offline-Architektur), wird aber nicht mehr in der App-UI genutzt.

## Komponenten

- `ToolCard`, `ToolboxBackground` — Home wie Web-Toolbox
- `ToolWebScreen` — WebView mit Tab- und Stack-Modus
- `src/constants/tools.ts` — zentrale Tool-Definitionen

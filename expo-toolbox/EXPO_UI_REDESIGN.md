# Expo UI Redesign — Native App Design

Datum: 2026-07-21  
Scope: **nur** `expo-toolbox` (keine PWA-/Datenmodell-/Repository-Änderungen)

---

## Ausgangslage (Analyse)

Die Expo-App war eine **webartige Landing-Shell**:

- Max-Width-Hero, Wrap-Grid mit Tool-Cards
- Keine Bottom-Tabs
- Keine nativen Listen/Formulare
- Produktworkflows nur als WebView (SiteReport / Bautagebuch)

Elemente, die „wie Website“ wirkten: fluid hero typography, Marketing-Copy, Card-Grid, fehlende Tab-Bar, kleine Toolbar-Links.

---

## Designentscheidungen

1. **Mobile First / Einhandbedienung** — Bottom Tabs, FAB für Create, große Touch-Targets (≥ 44px)
2. **Weniger gleichzeitig** — Dashboard mit wenigen Stat-Cards statt Tabellen
3. **Listen = Cards** — `ListItem` statt Tabellen
4. **Formulare baustellentauglich** — große Inputs, Footer-Primary-Buttons, KeyboardAvoiding
5. **Fotos** — `+ Foto hinzufügen` → Kamera (`expo-image-picker`) → Offline-Speichern → Galerie
6. **Web-Tools bleiben** — erreichbar unter Tab „Mehr“ (Funktionalität erhalten)
7. **Dark Mode** — Token `darkColors` vorbereitet, UI bleibt Light
8. **Keine Businesslogik in UI** — Screens rufen bestehende Repositories/Services auf

---

## Navigation

```
Root Stack
├── (tabs)
│   ├── Home (Dashboard)
│   ├── Projekte
│   ├── Tagebuch
│   ├── Mängel
│   └── Mehr (Dokumente · Web-Tools · Einstellungen)
├── project/new | project/[id]
├── diary/new | diary/[id]
├── defect/new | defect/[id]
├── sitereport (WebView)
└── bautagebuch (WebView)
```

---

## Geänderte / neue Screens

| Screen | Pfad | Änderung |
|--------|------|----------|
| Dashboard | `app/(tabs)/index.tsx` | Stat-Cards, Schnellzugriff, Offline-Banner |
| Projekte | `app/(tabs)/projects.tsx` | Card-Liste + FAB |
| Bautagebuch | `app/(tabs)/diary.tsx` | Card-Liste inkl. Fotoanzahl + FAB |
| Mängel | `app/(tabs)/defects.tsx` | Card-Liste + Priorität/Status + FAB |
| Mehr | `app/(tabs)/more.tsx` | Dokumente, Web-Tools, Settings |
| Projekt neu/detail | `app/project/*` | Formular + Foto-Galerie |
| Tagebuch neu/detail | `app/diary/*` | Abschnittsformular + Foto-Workflow |
| Mangel neu/detail | `app/defect/*` | Prioritäts-Chips + Foto-Workflow |
| Root Layout | `app/_layout.tsx` | Stack + Tabs, slide animation |
| Tabs Layout | `app/(tabs)/_layout.tsx` | Bottom Tabs |
| WebView Chrome | `ToolWebScreen.tsx` | Native Back, Touch-Targets |

Entfernt: altes Landing-`app/index.tsx`

---

## Neue Komponenten (`src/components/mobile/`)

| Komponente | Zweck |
|------------|--------|
| `Card` | Interaktions-/Listen-Container |
| `Section` | Abschnittskopf |
| `PrimaryButton` | Große Primary/Secondary/Ghost/Danger Buttons |
| `EmptyState` | Leere Listen mit CTA |
| `StatusBadge` | Status-/Prioritätsanzeige |
| `ListItem` | Card-Zeile für Listen |
| `Screen` | Safe Area, Header, Keyboard, Pull-to-Refresh, Footer |
| `Fab` | Floating Action Button |
| `TextField` | Große Formularfelder |
| `StatCard` | Dashboard-Kennzahl |

Design Tokens: `src/constants/theme.ts` (colors, darkColors, spacing, typography, shadows)

Hilfen: `src/lib/format.ts`, `src/lib/photoCapture.ts`

---

## Fotoworkflow

1. Button **+ Foto hinzufügen**
2. Kamera-Permission
3. `ImagePicker.launchCameraAsync`
4. `savePhotoOffline` (bestehender Service)
5. Große Vorschau in Detail-Screens

---

## Checks

- `npm run typecheck` (Expo)
- `npx expo export --platform ios`

---

## Nicht geändert

- PWA
- SQLite-Schema / Migrationen
- Repository Layer / Shared Types
- Offline-Architektur / Backup-Logik

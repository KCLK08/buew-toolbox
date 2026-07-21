# BÜW-Toolbox (Expo)

Native App-Hülle für die BÜW-Toolbox-PWAs: **SiteReport** und **Bautagebuch** laufen als eingebettete Web-Apps mit nativer Tab-Navigation.

```bash
cd expo-toolbox
cp .env.example .env
npm install
npm start
```

## App-Struktur

- **Home** — Toolbox-Übersicht (wie `index.html`)
- **SiteReport** — PWA `/sitereport/` im WebView
- **Bautagebuch** — PWA `/bautagebuch/` im WebView

Die PWAs behalten alle Funktionen (PDF-Export, Protokolle, Vorlagen, Fotos, Offline). Die App liefert native Navigation, Safe Area und Touch-optimierte Shell.

## Konfiguration

`EXPO_PUBLIC_TOOLBOX_WEB_BASE_URL` (Standard: `https://kclk08.github.io/buew-toolbox`)

## Android APK (GitHub Actions)

Bei Änderungen an `expo-toolbox/` oder `shared/` (Push auf `main`) erscheint eine **`.apk`** unter **Releases**.

- **Releases** → Datei `*.apk` herunterladen (installierbar)
- Actions-Artifacts sind bei GitHub immer ein **ZIP-Container** (nur für PRs)
- Workflow: `.github/workflows/android-apk.yml`
- Details: [`ANDROID_APK.md`](../ANDROID_APK.md)

Optional lokal/EAS: `eas.json` (Profile `preview` = APK).

PRs prüfen nur TypeScript; der APK-Build läuft einmal nach Merge auf `main`.

## Hinweis

Die Expo-App ist eine **PWA-Shell** — keine separate native CRUD-Oberfläche. Datenhaltung erfolgt in den eingebetteten PWAs (IndexedDB im WebView).

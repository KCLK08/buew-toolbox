# BÜW-Toolbox (Expo)

**Eigenständige native App** — dieselben Werkzeuge wie die PWA, ohne WebView.

```bash
cd expo-toolbox
cp .env.example .env
npm install
npm start
```

## Werkzeuge

| Tab | Beschreibung |
|-----|--------------|
| **Home** | Toolbox-Übersicht |
| **SiteReport** | Foto-Protokolle (nativ) |
| **Bautagebuch** | eBTB Guided Flow + PDF-Export (nativ) |

## Technik

- React Native / Expo Router
- SQLite (`bautagebuch_v2_native.db`, `sitereport_native.db`)
- PWA-Logik portiert nach `src/native/bautagebuch/lib/` (setup-model, etb-template, pdf-export)
- PDF: `pdf-lib` · Kamera: `expo-image-picker` · Wetter: `expo-location` + Open-Meteo

Details: [`NATIVE_APP.md`](../NATIVE_APP.md)

## Android APK

Siehe [`ANDROID_APK.md`](../ANDROID_APK.md).

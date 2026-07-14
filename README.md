# Bautagebuch

Expo-App für das elektronische Bautagebuch (eBTB) – portiert aus der Buew-Toolbox mit allen Kernfunktionen.

## Funktionen

- **Standard eBTB-Vorlage** (`Vorlage-eBTB.pdf`) mit fix definiertem Formular-Setup
- **Bautagebuch erstellen & bearbeiten** mit Abschnitten:
  - Kopfdaten
  - Witterung (inkl. automatischer Wetter-Sync per GPS)
  - Baustellenbesetzung (dynamische Zeilen)
  - Leistungsblock
  - Abschluss
  - Fotodokumentation
- **Offline-Speicherung** via SQLite + lokales Dateisystem
- **PDF-Export** als BTB, Fotodoku oder kombiniert
- **Eigene PDF-Vorlagen** hochladen (AcroForm)
- **Autosave** während der Bearbeitung

## Starten

```bash
npm install
npm start
```

Dann mit Expo Go (Android/iOS) oder `npm run web` im Browser testen.

## Technologie

- Expo SDK 57 + Expo Router
- TypeScript
- pdf-lib für PDF-Formulare
- expo-sqlite für lokale Daten
- expo-image-picker / expo-location

## Ursprung

Diese App basiert auf dem Bautagebuch-Tool aus [buew-toolbox](https://github.com/KCLK08/buew-toolbox) und erweitert es als native/mobile Expo-Anwendung.

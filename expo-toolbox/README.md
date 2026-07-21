# BÜW-Toolbox (Expo)

Native Offline-App der Toolbox mit Bottom-Tabs. SiteReport und Bautagebuch bleiben als eingebettete Web-Tools unter **Mehr** verfügbar. Persistenz: **SQLite + documentDirectory**.

```bash
cd expo-toolbox
cp .env.example .env
npm install
npm start
```

## UI

Siehe [`EXPO_UI_REDESIGN.md`](../EXPO_UI_REDESIGN.md).

- Bottom Tabs: Home · Projekte · Tagebuch · Mängel · Mehr
- Mobile Design-System unter `src/components/mobile/`
- Fotoaufnahme über Kamera (`expo-image-picker`) → Offline-Speicherung

## Offline-Architektur

Siehe `/OFFLINE.md`, `/OFFLINE_DATA_MODEL.md` und `/DATA_MODEL_PARITY_REPORT.md`.

Wichtige Module:

- `../shared/types/` – gemeinsame Domain-Typen (`DOMAIN_SCHEMA_VERSION = 2`)
- `src/database/` – SQLite, Migrationen (v2 Parität)
- `src/storage/` – Dateien (`documentDirectory`) + DB-Backups (max. 3)
- `src/repositories/` – Shared API (projects, diary, defects, photos, documents)
- `src/services/` – Integritätsprüfung, Foto-Speicherung
- `src/hooks/useAutosave.ts` – debounced Autosave
- `src/hooks/useOfflineBootstrap.ts` – Startcheck

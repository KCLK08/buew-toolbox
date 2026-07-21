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

Siehe `/OFFLINE.md` (inkl. Phase-2 Backup-Härtung), `/OFFLINE_DATA_MODEL.md` und `/DATA_MODEL_PARITY_REPORT.md`.

Wichtige Module:

- `../shared/types/` – gemeinsame Domain-Typen (`DOMAIN_SCHEMA_VERSION = 2`)
- `src/database/` – SQLite (WAL, `synchronous=NORMAL`), Migrationen, Write-Guards
- `src/storage/` – Dateien (`documentDirectory`) + sichere DB-Backups (max. 3, max. 1/min)
- `src/repositories/` – Shared API (projects, diary, defects, photos, documents)
- `src/services/` – Integrität, Orphan-Cleanup (nur melden), Soft-Delete-Purge (Dry-Run)
- `src/hooks/useAutosave.ts` – debounced Autosave (ohne Backup pro Save)
- `src/hooks/useOfflineBootstrap.ts` – Startcheck, Restore-Bestätigung, Background-Backup

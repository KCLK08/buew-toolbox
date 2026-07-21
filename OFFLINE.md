# Offline-Architektur BÜW-Toolbox

Die Toolbox läuft **vollständig offline**.

- Kein Login
- Keine Cloud / Supabase
- Keine Sync-Server
- Alle Geschäftsdaten lokal auf dem Gerät

## Clients

| Client | Persistenz | Dateien |
|--------|------------|---------|
| PWA (`bautagebuch-v2`, `sitereport`) | IndexedDB (Dexie) | Binärdaten in IndexedDB (`photo_assets`), nicht im Browser-Cache |
| Expo (`expo-toolbox`) | SQLite (`expo-sqlite`) | Fotos/Dokumente in `documentDirectory` (`expo-file-system`) |

## Gemeinsames Datenmodell (Phase 3)

- Spezifikation: [`OFFLINE_DATA_MODEL.md`](./OFFLINE_DATA_MODEL.md)
- Parity-Report: [`DATA_MODEL_PARITY_REPORT.md`](./DATA_MODEL_PARITY_REPORT.md)
- Shared Types: `shared/types/` (`DOMAIN_SCHEMA_VERSION = 2`)
- Repositories: gleiche Methodennamen auf Expo (SQLite) und PWA (IndexedDB)

## Expo-Struktur

```
expo-toolbox/src/
  database/       SQLite, Schema, Migrationen
  storage/        FileService, BackupService
  repositories/   Datenzugriff (shared API)
  services/       Integrity, Autosave-Koordination
  hooks/          useOfflineBootstrap, useAutosave
```

## Garantien

1. **Autosave** – Änderungen werden sofort persistiert (keine manuelle Speicherung).
2. **Transaktionen** – zusammenhängende Writes atomar (BEGIN/COMMIT/ROLLBACK).
3. **Backups** – bis zu 3 rotierende DB-Kopien nach wichtigen Änderungen.
4. **Startcheck** – DB/Tabellen/Dateien prüfen, ggf. Restore aus Backup.
5. **Soft Delete** – `deleted_at` statt sofortigem Hard-Delete.
6. **Migrationen** – Schema-Updates ohne Datenverlust.
7. **Datenmodell-Parität** – Expo und PWA nutzen dieselben fachlichen Typen und Repository-Methoden.

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
  services/       Integrity, Orphan-Cleanup, Soft-Delete-Purge, Autosave
  hooks/          useOfflineBootstrap, useAutosave
```

## Garantien

1. **Autosave** – Änderungen werden sofort persistiert (keine manuelle Speicherung).
2. **Transaktionen** – zusammenhängende Writes atomar (BEGIN/COMMIT/ROLLBACK).
3. **Backups** – bis zu 3 rotierende DB-Kopien nach wichtigen Ereignissen (nicht nach jedem Autosave).
4. **Startcheck** – DB/Tabellen/Dateien prüfen; Restore nur nach Benutzerbestätigung.
5. **Soft Delete** – `deleted_at` statt sofortigem Hard-Delete (Purge vorbereitet, nicht aktiv).
6. **Migrationen** – Schema-Updates ohne Datenverlust.
7. **Datenmodell-Parität** – Expo und PWA nutzen dieselben fachlichen Typen und Repository-Methoden.

## Phase 2 – Backup-Härtung

### SQLite (Expo)

- Vor jedem Backup: `PRAGMA wal_checkpoint(FULL)`
- `PRAGMA synchronous = NORMAL`
- Backup nur wenn kein Schreibvorgang aktiv ist; sonst verzögert
- Throttling: maximal 1 Backup pro Minute (außer `manual`)
- Trigger: `photo_added`, `record_deleted`, `status_change`, `app_background`

### PWA IndexedDB (Bautagebuch)

- Snapshots in `db_backups` enthalten **keine** Foto-Binärdaten
- Gesichert werden Metadaten/Referenzen; Binaries bleiben in `photo_assets`
- Restore merged Metadaten und erhält vorhandene Binaries
- Gleiche Backup-Trigger und 1/Minute-Throttling
- Restore: Hinweis mit Backup-Datum, explizite Bestätigung, kein stilles Überschreiben

### Orphan-Cleanup

- Expo: Dateien in `photos/` / `documents/` ohne aktiven DB-Eintrag melden
- PWA: `photo_assets` ohne Referenz in aktivem Run melden
- **Keine automatische Löschung**

### Soft-Delete-Purge (Vorbereitung)

- Dry-Run-Plan für Einträge mit `deleted_at` älter als Retention (Standard 30 Tage)
- `executeSoftDeletePurge` ist bewusst deaktiviert

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

## Expo-Struktur

```
expo-toolbox/src/
  database/       SQLite, Schema, Migrationen
  storage/        FileService, BackupService
  repositories/   Datenzugriff
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

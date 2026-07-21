# BÜW-Toolbox (Expo)

Offline Expo-Shell der Toolbox. SiteReport und Bautagebuch bleiben als eingebettete Web-Tools verfügbar. Persistenz der Shell: **SQLite + documentDirectory**.

```bash
cd expo-toolbox
cp .env.example .env
npm install
npm start
```

## Offline-Architektur

Siehe `/OFFLINE.md`.

Wichtige Module:

- `src/database/` – SQLite, Migrationen, Transaktionen
- `src/storage/` – Dateien (`documentDirectory`) + DB-Backups (max. 3)
- `src/repositories/` – Projekte, Bautagebuch-Läufe, Mängel, Notizen, Fotos
- `src/services/` – Integritätsprüfung, Foto-Speicherung
- `src/hooks/useAutosave.ts` – debounced Autosave
- `src/hooks/useOfflineBootstrap.ts` – Startcheck

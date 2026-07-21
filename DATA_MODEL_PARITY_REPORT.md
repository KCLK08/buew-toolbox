# Data Model Parity Report — Phase 3

Datum: 2026-07-21  
Domain-Schema-Version: **2**

---

## 1. Übereinstimmende Bereiche

Nach Phase 3 teilen Expo und PWA (Bautagebuch) dasselbe fachliche Modell:

| Bereich | Status |
|---------|--------|
| Shared Types (`shared/types`) | ✅ Einzelne Quelle |
| Domain-Schema-Version `2` | ✅ Expo `SCHEMA_VERSION` + PWA `app_meta` / `domain-schema.js` |
| Project inkl. description/location/date/status | ✅ |
| DiaryEntry inkl. entry_date + payload_json-Konvention | ✅ |
| Defect inkl. description/priority/diary_entry_id | ✅ |
| Note | ✅ |
| Photo mit parent_id/parent_type/filename/local_path/file_size | ✅ |
| Document | ✅ |
| Soft Delete über `deleted_at` | ✅ |
| Repository-Methodennamen (get/create/update/softDelete/add/delete) | ✅ |

---

## 2. Unterschiede (bewusst / historisch)

| Thema | Expo | PWA Bautagebuch | Bewertung |
|-------|------|-----------------|-----------|
| Persistenz | SQLite | IndexedDB (Dexie) | Erlaubt — gleiche Fach-API |
| Physischer Diary-Tabellenname | `diary_runs` | `diary_entries` | Dokumentiert; fachlich DiaryEntry |
| Legacy-Spalten Photos | `file_path`, `byte_size`, `diary_run_id`, … bleiben befüllt | nur Domain-Felder | Expo-Migration füllt beide |
| UI-Template-Layer | — | `templates` / `runs` / `photo_assets` | Zusätzlich; Domain parallel |
| SiteReport | — | Eigenes Protokoll-Modell | **Nicht** auf Project/Diary gemappt |
| Foto-Binaries | Dateisystem | `photo_assets` / Blobs | Metadaten-Shape identisch |

---

## 3. Durchgeführte Anpassungen

1. **`shared/types/`** angelegt (Project, DiaryEntry, Defect, Note, Photo, Document, common).
2. **`OFFLINE_DATA_MODEL.md`** als verbindliche Spezifikation.
3. **Expo Migration v2** `domain_parity_v2`: neue Felder + `documents`-Tabelle; Daten-Backfill.
4. **Expo Repositories** auf gemeinsame API umgestellt; Legacy-Aliase (`listActive`, `upsert`, `diaryRunRepository`, …) bleiben für Kompatibilität.
5. **Expo Types** re-exportieren Shared Types (keine abweichenden Domain-Typen).
6. **PWA Dexie v4** Domain-Stores + `bautagebuch-v2/src/lib/repositories/` mit gleicher API.
7. **`DOMAIN_SCHEMA_VERSION = 2`** zentral in Shared / Expo constants / PWA domain-schema.

---

## 4. Offene Punkte

1. **UI-Anbindung:** Bestehende Bautagebuch-UI nutzt weiter `runs`/`templates`; Domain-Repositories sind bereit, aber noch nicht in allen Screens verdrahtet (bewusst keine neue Funktion).
2. **SiteReport-Parität:** Protokolle ≠ Diary/Project — ggf. spätere Abbildung oder bewusst separate Domäne.
3. **Legacy-Spalten Expo:** `file_path`/`byte_size`/`notes`/`diary_run_id` können in einer späteren Migration entfernt werden, sobald keine Leser mehr darauf zugreifen.
4. **PWA `photo_assets` ↔ Domain `photos`:** Noch kein automatischer Sync zwischen Template-Foto-Flow und Domain-Photo-Store.
5. **Phase-2 Backup-Härtung** liegt auf separatem Branch/PR — dieser Report bezieht sich auf Domain-Parität.

---

## Validierungsfragen

| Frage | Antwort |
|-------|---------|
| Entitäten nur Expo? | Vorher: documents/domain-Felder fehlten in PWA → jetzt Domain-Stores vorhanden |
| Entitäten nur PWA? | UI-Layer (`templates`, `runs`, …) bleibt PWA-spezifisch |
| Unterschiedliche Feldnamen? | Domain: nein (snake_case). Legacy Expo-Spalten zusätzlich |
| Unterschiedliche Soft-Delete-Regeln? | Domain: nein (`deleted_at`) |
| Unterschiedliche Foto-Logik? | Metadaten vereinheitlicht; Binärspeicher plattformspezifisch |

# Offline-Datenmodell (gemeinsam)

**Domain-Schema-Version: `2`** (`shared/types` → `DOMAIN_SCHEMA_VERSION`)

Gültig für:

| Plattform | Persistenz | Physische Schema-Version |
|-----------|------------|--------------------------|
| Expo | SQLite (`buew_toolbox.db`) | `SCHEMA_VERSION = 2` |
| PWA Bautagebuch | IndexedDB Dexie (`BautagebuchV2`) | Domain `2` / Dexie store format `4` |

Keine Cloud, kein Login, keine Sync, kein Gerätewechsel. Expo und PWA bleiben getrennte lokale Apps mit **identischer fachlicher Struktur**.

---

## Konventionen

- IDs: UUID-Strings
- Zeitstempel: ISO-8601 Text (`created_at`, `updated_at`, `deleted_at`)
- Soft Delete: `deleted_at = null` → aktiv; gesetzt → gelöscht (kein Hard-Delete in der Domain-API)
- Feldnamen: **snake_case** in allen Domain-Objekten
- Status Entitäten: `draft | active | archived | completed`
- Foto-/Dokument-Status: `ready | pending | error | deleted`

---

## Entitäten

### Project

| Feld | Typ | Pflicht | Beschreibung |
|------|-----|---------|--------------|
| `id` | string (UUID) | ja | Primärschlüssel |
| `name` | string | ja | Projektname |
| `description` | string | nein (Default `''`) | Beschreibung |
| `location` | string | nein (Default `''`) | Ort |
| `date` | string \| null | nein | ISO-Datum `YYYY-MM-DD` |
| `status` | EntityStatus | ja | Default `active` |
| `created_at` | string | ja | |
| `updated_at` | string | ja | |
| `deleted_at` | string \| null | ja | Soft Delete |

**Beziehungen:** 1:n → DiaryEntry, Defect, Note, Photo, Document

---

### DiaryEntry (Bautagebuch)

| Feld | Typ | Pflicht | Beschreibung |
|------|-----|---------|--------------|
| `id` | string (UUID) | ja | |
| `project_id` | string \| null | nein | FK Project |
| `title` | string | ja | |
| `entry_date` | string \| null | nein | ISO-Datum |
| `status` | EntityStatus | ja | Default `draft` |
| `payload_json` | string | ja | JSON-String der DiaryPayload |
| `created_at` | string | ja | |
| `updated_at` | string | ja | |
| `deleted_at` | string \| null | ja | Soft Delete |

**Physischer Tabellenname Expo:** `diary_runs` (Legacy). Fachlicher Name überall: `DiaryEntry`.

#### DiaryPayload (Inhalt von `payload_json`)

| Schlüssel | Typ | Inhalt |
|-----------|-----|--------|
| `weather` | object | `category`, `temp_min`, `temp_max`, `notes` |
| `personnel` | array | `{ name, role, hours }` |
| `equipment` | array | `{ name, count, notes }` |
| `work` | array | `{ description, location, notes }` |
| `notes` | string | Freitext |

---

### Defect (Mangel)

| Feld | Typ | Pflicht | Beschreibung |
|------|-----|---------|--------------|
| `id` | string (UUID) | ja | |
| `project_id` | string \| null | nein | FK Project |
| `diary_entry_id` | string \| null | nein | FK DiaryEntry |
| `title` | string | ja | |
| `description` | string | nein (Default `''`) | Beschreibung |
| `priority` | `low \| normal \| high \| critical` | ja | Default `normal` |
| `status` | EntityStatus | ja | Default `draft` |
| `created_at` | string | ja | |
| `updated_at` | string | ja | |
| `deleted_at` | string \| null | ja | Soft Delete |

---

### Note

| Feld | Typ | Pflicht | Beschreibung |
|------|-----|---------|--------------|
| `id` | string (UUID) | ja | |
| `project_id` | string \| null | nein | |
| `diary_entry_id` | string \| null | nein | |
| `defect_id` | string \| null | nein | |
| `body` | string | ja | |
| `created_at` / `updated_at` / `deleted_at` | string / string / string\|null | ja | Soft Delete |

---

### Photo

Einheitliche Metadaten (Binärdaten **nicht** in der Domain-Zeile):

| Feld | Typ | Pflicht | Beschreibung |
|------|-----|---------|--------------|
| `id` | string (UUID) | ja | |
| `parent_id` | string \| null | nein | ID des Elternobjekts |
| `parent_type` | `project \| diary_entry \| defect \| note` \| null | nein | |
| `filename` | string | ja | Dateiname |
| `local_path` | string | ja | Relativer Pfad (Expo) oder logischer Key (PWA) |
| `mime_type` | string | ja | |
| `file_size` | number | ja | Bytes |
| `status` | PhotoStatus | ja | Default `ready` |
| `created_at` / `updated_at` / `deleted_at` | … | ja | Soft Delete setzt auch `status=deleted` |

**Binärschutz:**

- Expo: Datei unter `documentDirectory/photos/`
- PWA: Binaries in `photo_assets` (UI-Template-Flow) bzw. separat referenziert; Domain-`photos` nur Metadaten

---

### Document

| Feld | Typ | Pflicht | Beschreibung |
|------|-----|---------|--------------|
| `id` | string (UUID) | ja | |
| `parent_id` | string \| null | nein | |
| `parent_type` | ParentType \| null | nein | |
| `filename` | string | ja | |
| `local_path` | string | ja | |
| `mime_type` | string | ja | |
| `file_size` | number | ja | |
| `status` | DocumentStatus | ja | Default `ready` |
| `created_at` / `updated_at` / `deleted_at` | … | ja | Soft Delete |

---

## Soft-Delete-Regeln

1. Alle Domain-Entitäten haben `deleted_at`.
2. Listen-APIs liefern nur `deleted_at IS NULL` / unset.
3. Cascade bei Soft-Delete von DiaryEntry: zugehörige Photos, Defects, Notes (soft).
4. Cascade bei Soft-Delete von Defect: zugehörige Photos (soft).
5. Hard-Purge ist **nicht** aktiv (nur vorbereitet in späteren Phasen).

---

## Repository-API (fachlich identisch)

### projectRepository
- `getProjects()`
- `getProjectById(id)`
- `createProject(input)`
- `updateProject(input)`
- `softDeleteProject(id)`

### diaryRepository
- `getDiaryEntries(projectId?)`
- `getDiaryEntryById(id)`
- `createDiaryEntry(input)`
- `updateDiaryEntry(input)`
- `softDeleteDiaryEntry(id)`

### defectRepository
- `getDefects(projectId?)`
- `getDefectById(id)`
- `createDefect(input)`
- `updateDefect(input)`
- `softDeleteDefect(id)`

### photoRepository
- `addPhoto(input)`
- `getPhotos(filter?)`
- `deletePhoto(id)`  *(Soft Delete)*

### documentRepository
- `addDocument(input)`
- `getDocuments(filter?)`
- `softDeleteDocument(id)`

### noteRepository
- `getNotes()`
- `createNote(input)`
- `updateNote(input)`
- `softDeleteNote(id)`

Implementierungen:

- Expo: `expo-toolbox/src/repositories/index.ts` (SQLite)
- PWA: `bautagebuch-v2/src/lib/repositories/index.js` (IndexedDB)

---

## TypeScript Types

Pfad: `shared/types/`

- `common.ts` — Status, SoftDelete, `DOMAIN_SCHEMA_VERSION`
- `Project.ts`, `DiaryEntry.ts`, `Defect.ts`, `Note.ts`, `Photo.ts`, `Document.ts`
- `index.ts` — Re-Exports

Expo importiert ausschließlich diese Domain-Typen (plus lokale Integrity-Typen).

---

## Migrationen

### Expo (SQLite)

| Version | Name | Inhalt |
|---------|------|--------|
| 1 | `initial_offline_schema` | Basis-Tabellen |
| 2 | `domain_parity_v2` | Project-Felder, `entry_date`, Defect description/priority, Photo parent_*, Documents |

### PWA Bautagebuch (Dexie)

| Dexie | Domain | Inhalt |
|-------|--------|--------|
| 1–3 | — | Template/Run/photo_assets/db_backups (UI-Layer) |
| 4 | **2** | Domain-Stores: projects, diary_entries, defects, notes, photos, documents, app_meta |

`app_meta.domain_schema_version` = `"2"`.

---

## Plattform-Zusatz (nicht Domain)

Diese Stores gehören zur bestehenden UI und sind **kein** Ersatz für das Domain-Modell:

- PWA: `templates`, `runs`, `setup_models`, `detected_fields`, `exports`, `photo_assets`, `db_backups`
- SiteReport: eigenes Protokoll-Modell (`protocols`, `entries`, …) — siehe Parity-Report

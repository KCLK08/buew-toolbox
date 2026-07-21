# BÜW-Toolbox

Zwei Clients derselben Anwendung:

| Client | Ordner | Stack |
|--------|--------|-------|
| PWA (Web) | `web/` | React + Vite + React Router |
| Mobile | `mobile/` | Expo React Native |
| Shared | `shared/` | Auth, Validierung, Typen, Services |

## Architekturregel

Neue Features, Bugfixes, Datenmodelle, Navigation und Supabase-Änderungen werden **immer für beide Clients** umgesetzt. Businesslogik liegt in `shared/`.

## Schnellstart

```bash
npm install
cp mobile/.env.example mobile/.env
cp web/.env.example web/.env

# Web
npm run dev:web

# Mobile
npm run dev:mobile
```

## Supabase

1. Env-Werte aus `mobile/.env.example` / `web/.env.example`
2. Migrationen in `supabase/migrations/` im SQL-Editor ausführen
3. Auth ausschließlich über Supabase Auth

## Auth

Gemeinsam:

- Login / Registrierung / Passwort vergessen
- Session-Handling + Auto-Refresh
- Geschützte Routen
- Rollen: `admin`, `benutzer` (Default: `benutzer`)

Nur Mobile:

- Biometrische Anmeldung (Face ID / Fingerabdruck) via SecureStore

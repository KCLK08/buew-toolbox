# BÜW-Toolbox (Expo)

Native Expo-Kopie der BÜW-Toolbox-Startseite. Enthält SiteReport und Bautagebuch (als eingebettete Web-Tools) und ist für die Anbindung an Supabase vorbereitet.

## Start

```bash
cd expo-toolbox
cp .env.example .env
npm install
npm start
```

## Tools

| Tool | Route | Inhalt |
|------|-------|--------|
| SiteReport | `/sitereport` | WebView auf die bestehende Web-App |
| Bautagebuch | `/bautagebuch` | WebView auf die bestehende Web-App |

Die Web-Basis-URL steuert `EXPO_PUBLIC_TOOLBOX_WEB_BASE_URL` (Standard: `https://kclk08.github.io/buew-toolbox`).

## Supabase vorbereiten

1. Projekt in Supabase anlegen
2. `.env` mit `EXPO_PUBLIC_SUPABASE_URL` und `EXPO_PUBLIC_SUPABASE_ANON_KEY` füllen
3. SQL aus `supabase/migrations/20260721000000_profiles.sql` im SQL-Editor ausführen
4. Optional: Typen aktualisieren mit `supabase gen types typescript`

Der Client liegt in `src/lib/supabase.ts`, Session-Handling in `src/providers/AuthProvider.tsx`.
Ohne gesetzte Env-Variablen startet die App weiterhin – Auth bleibt dann deaktiviert.

# BÜW-Toolbox (Expo)

Native Expo-Kopie der BÜW-Toolbox-Startseite. Enthält SiteReport und Bautagebuch (als eingebettete Web-Tools) und ist für die Anbindung an Supabase vorbereitet.

## Start

```bash
cd expo-toolbox
cp .env.example .env
npm install
npm start
```

Die Supabase-URL und der Publishable Key liegen bereits in `.env.example`. Nach dem Kopieren nach `.env` ist die Client-Anbindung aktiv.

## Tools

| Tool | Route | Inhalt |
|------|-------|--------|
| SiteReport | `/sitereport` | WebView auf die bestehende Web-App |
| Bautagebuch | `/bautagebuch` | WebView auf die bestehende Web-App |

Die Web-Basis-URL steuert `EXPO_PUBLIC_TOOLBOX_WEB_BASE_URL` (Standard: `https://kclk08.github.io/buew-toolbox`).

## Supabase

Bereits vorbereitet:

1. `.env.example` enthält `EXPO_PUBLIC_SUPABASE_URL` und `EXPO_PUBLIC_SUPABASE_ANON_KEY`
2. Client: `src/lib/supabase.ts`
3. Session-Handling: `src/providers/AuthProvider.tsx`
4. SQL für Profile + RLS: `supabase/migrations/20260721000000_profiles.sql` (einmal im Supabase SQL-Editor ausführen)

Optional Typen aktualisieren mit `supabase gen types typescript`.

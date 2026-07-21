# BÜW-Toolbox Web (PWA)

React/Vite-Client der Toolbox mit Supabase-Auth.

```bash
cp .env.example .env
npm install
npm run dev
```

## Routen

| Pfad | Zugriff |
|------|---------|
| `/login` | öffentlich |
| `/register` | öffentlich |
| `/forgot-password` | öffentlich |
| `/` | geschützt (Dashboard) |
| `/settings` | geschützt |

Businesslogik liegt in `@buew/shared`.

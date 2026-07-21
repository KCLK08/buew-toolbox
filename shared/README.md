# @buew/shared

Gemeinsame Businesslogik für Web-PWA und Expo-App.

## Inhalt

| Pfad | Zweck |
|------|-------|
| `src/types/` | Auth-, Rollen- und Database-Typen |
| `src/validation/` | Zod-Schemas (Login, Register, Forgot) |
| `src/services/authService.ts` | Supabase Auth Service |
| `src/auth/AuthProvider.tsx` | AuthContext + `useAuth` + Session-Listener |
| `src/constants/` | Design-Tokens, Auth-Copy |
| `src/lib/` | Fehler-Mapping, Network-Monitor |

UI bleibt plattformspezifisch in `web/` und `mobile/`.

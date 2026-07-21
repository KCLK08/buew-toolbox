# BÜW-Toolbox Mobile (Expo)

Expo-Client der Toolbox. Nutzt `@buew/shared` für Auth-Businesslogik.

```bash
cp .env.example .env
npm install
npm start
```

## Auth

- Login / Registrierung / Passwort vergessen
- Session in SecureStore (chunked)
- Geschützte Screens
- Biometrie (nur Mobile): Face ID / Fingerabdruck

## Routen

| Route | Zugriff |
|-------|---------|
| `/login`, `/register`, `/forgot-password` | öffentlich |
| `/biometric-unlock` | Session + Biometrie aktiv |
| `/dashboard`, `/settings`, `/sitereport`, `/bautagebuch` | geschützt |

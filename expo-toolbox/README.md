# BÜW-Toolbox (Expo)

Offline-fähige Expo-Kopie der BÜW-Toolbox-Startseite. Enthält SiteReport und Bautagebuch als eingebettete Web-Tools. Kein Login, kein Backend.

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

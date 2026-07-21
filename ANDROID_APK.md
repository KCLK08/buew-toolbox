# Android APK via GitHub Actions

Bei relevanten Änderungen erzeugt GitHub automatisch eine **standalone Release-APK** mit eingebettetem JavaScript-Bundle.

## Wichtig: `.apk` vs `.zip`

| Quelle | Dateityp | Hinweis |
|--------|----------|---------|
| **Releases** | `.apk` | Richtiger Download für die Installation |
| **Actions → Artifacts** | immer `.zip` | GitHub packt Artifacts immer als ZIP (darin liegt die APK) |

Für die Installation immer unter **Releases** die `.apk` herunterladen.

## Warum Release statt Debug?

Eine **Debug-APK** erwartet einen laufenden Metro-Bundler (`localhost:8081`) und startet auf dem Gerät nicht eigenständig. Die Fehlermeldung *„Unable to load script“* bedeutet genau das.

Die CI baut deshalb `assembleRelease`. Dabei wird das JS-Bundle per `export:embed` in die APK gepackt — die App läuft ohne Entwicklungsrechner.

## Trigger

- Push auf `main` mit Änderungen in `expo-toolbox/**` oder `shared/**`
- Pull Requests gegen `main` (nur Artifact-ZIP, kein Release)
- Manuell: **Actions → Build Android APK → Run workflow** (Release nur auf `main`)

## Ergebnis auf `main`

- GitHub Release mit Tag `apk-v<version>.<run_number>`
- Asset: `buew-toolbox-<version>-<sha>.apk`
- Pre-Release / Latest

## Download

1. Repo → **Releases**
2. Neuestes Pre-Release öffnen
3. Unter Assets die **`.apk`**-Datei laden
4. Auf dem Android-Gerät installieren (ggf. „Unbekannte Quellen“ erlauben)

## App-Icon

Icons werden bei `expo prebuild` aus `app.config.js` erzeugt (`icon`, `android.adaptiveIcon`). Nach Installation erscheint das BÜW-Logo im Launcher.

## Technik

1. `npm ci` in `expo-toolbox`
2. `npx expo prebuild --platform android`
3. `./gradlew assembleRelease` (JS-Bundle + Assets eingebettet)
4. CI prüft, dass `assets/index.android.bundle` (oder Hermes `.hbc`) in der APK liegt
5. Auf `main`: `gh release create … datei.apk`

Native Ordner `android/` werden nicht committed (CNG / `.gitignore`).

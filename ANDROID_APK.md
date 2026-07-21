# Android APK via GitHub Actions

Bei relevanten Änderungen erzeugt GitHub automatisch eine installierbare **Debug-APK**.

## Wichtig: `.apk` vs `.zip`

| Quelle | Dateityp | Hinweis |
|--------|----------|---------|
| **Releases** | `.apk` | Richtiger Download für die Installation |
| **Actions → Artifacts** | immer `.zip` | GitHub packt Artifacts immer als ZIP (darin liegt die APK) |

Für die Installation immer unter **Releases** die `.apk` herunterladen.

## Trigger

- Push auf `main` mit Änderungen in `expo-toolbox/**` oder `shared/**`
- Pull Requests gegen `main` (nur Artifact-ZIP, kein Release)
- Manuell: **Actions → Build Android APK → Run workflow** (Release nur auf `main`)

## Ergebnis auf `main`

- GitHub Release mit Tag `apk-v<version>.<run_number>`
- Asset: `buew-toolbox-<version>-<sha>-debug.apk`
- Pre-Release / Latest

## Download

1. Repo → **Releases**
2. Neuestes Pre-Release öffnen
3. Unter Assets die **`.apk`**-Datei laden
4. Auf dem Android-Gerät installieren

## Technik

1. `npm ci` in `expo-toolbox`
2. `npx expo prebuild --platform android`
3. `./gradlew assembleDebug`
4. Auf `main`: `gh release create … datei.apk`

Native Ordner `android/` werden nicht committed (CNG / `.gitignore`).

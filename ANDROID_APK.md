# Android APK via GitHub Actions

Bei relevanten Änderungen erzeugt GitHub automatisch eine installierbare **Debug-APK** und veröffentlicht sie unter **Releases**.

## Trigger

- Push auf `main` mit Änderungen in `expo-toolbox/**` oder `shared/**`
- Pull Requests gegen `main` (nur Artifact, kein Release)
- Manuell: **Actions → Build Android APK → Run workflow** (Release nur auf `main`)

## Ergebnis

### Releases (sichtbar unter GitHub → Releases)

- Tag: `apk-v<version>.<run_number>` (z. B. `apk-v1.0.0.42`)
- Titel: `BÜW-Toolbox Android <version> (#<run>)`
- Asset: `buew-toolbox-<version>-<sha>-debug.apk`
- Als Pre-Release markiert, jeweils als **Latest**

### Workflow-Artifact (zusätzlich)

- Name: `buew-toolbox-android-apk`
- Aufbewahrung: 30 Tage
- Pfad: Actions → Lauf → Artifacts

## Download

1. Repo → **Releases**
2. Neuestes Pre-Release öffnen
3. APK unter Assets herunterladen

## Technik

1. `npm ci` in `expo-toolbox`
2. `npx expo prebuild --platform android`
3. `./gradlew assembleDebug`
4. Upload als Artifact
5. Auf `main`: GitHub Release mit APK-Asset

Native Ordner `android/` werden nicht committed (CNG / `.gitignore`).

## Optional: EAS

`expo-toolbox/eas.json` enthält Profile (`preview` = APK). Für Cloud-Builds zusätzlich Expo-Account + `EXPO_TOKEN` nötig.

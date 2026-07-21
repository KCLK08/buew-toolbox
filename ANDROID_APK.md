# Android APK via GitHub Actions

Bei relevanten Änderungen erzeugt GitHub automatisch eine installierbare **Debug-APK**.

## Trigger

- Push auf `main` mit Änderungen in `expo-toolbox/**` oder `shared/**`
- Pull Requests gegen `main` mit denselben Pfaden
- Manuell: **Actions → Build Android APK → Run workflow**

## Ergebnis

- Workflow-Artifact: `buew-toolbox-android-apk`
- Dateiname: `buew-toolbox-<version>-<sha>-debug.apk`
- Aufbewahrung: 30 Tage
- Debug-signiert (kein Play-Store-Keystore nötig)

Download: GitHub → Actions → erfolgreicher Lauf → Artifacts

## Technik

1. `npm ci` in `expo-toolbox`
2. `npx expo prebuild --platform android`
3. `./gradlew assembleDebug`
4. Upload als Artifact

Native Ordner `android/` werden nicht committed (CNG / `.gitignore`).

## Optional: EAS

`expo-toolbox/eas.json` enthält Profile (`preview` = APK). Für Cloud-Builds zusätzlich Expo-Account + `EXPO_TOKEN` nötig.

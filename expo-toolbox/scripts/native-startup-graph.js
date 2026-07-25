#!/usr/bin/env node
'use strict';

/**
 * Maps the Expo/RN startup path before first React render.
 * Run: node scripts/native-startup-graph.js
 */

const fs = require('fs');
const path = require('path');

const ROOT = path.join(__dirname, '..');

const NATIVE_ANDROID = {
  phase: 'Native (before JS bundle executes)',
  steps: [
    { order: 1, file: 'MainApplication.kt', action: 'Application.onCreate → loadReactNative()' },
    { order: 2, file: 'PackageList (autolinked)', action: 'Register native modules (lazy until first JS call)' },
    { order: 3, file: 'MainActivity.kt', action: 'setTheme(AppTheme) for splash, ReactActivityDelegateWrapper' },
    { order: 4, file: 'Hermes runtime', action: 'Load embedded .hbc bundle from APK assets' }
  ]
};

const JS_SYNC = {
  phase: 'JS synchronous (before React renderRootComponent paints)',
  chain: [
    { id: 1, module: 'expo-router/entry.js', loads: ['expo-router/entry-classic.js'] },
    { id: 2, module: 'entry-classic.js', loads: ['@expo/metro-runtime', 'qualified-entry', 'renderRootComponent'] },
    { id: 3, module: 'renderRootComponent.js', loads: ['expo', 'react-native', 'utils/splash'] },
    { id: 4, module: 'expo (Expo.fx)', loads: ['./winter', './async-require', 'expo-asset'] },
    { id: 5, module: 'expo/src/winter/index.ts', loads: ['./runtime'] },
    { id: 6, module: 'expo/src/winter/runtime.native.ts', loads: [
      'react-native/Libraries/Core/InitializeCore',
      'expo/src/winter/AbortSignal ← CRASH if AbortSignal undefined',
      'expo/src/winter/FormData',
      'expo fetch polyfill'
    ]},
    { id: 7, module: 'InitializeCore.js', loads: ['setUpXHR → abort-controller polyfill for AbortSignal global'] },
    { id: 8, module: 'qualified-entry → ExpoRoot', loads: ['expo-router/_ctx (all routes lazy)'] },
    { id: 9, module: 'app/_layout.tsx', loads: ['Only when ExpoRoot resolves root layout — AFTER winter init'] }
  ]
};

const APP_CONFIG_PLUGINS = [
  { plugin: 'expo-router', nativeAtStartup: false, note: 'JS routing only' },
  { plugin: 'expo-asset', nativeAtStartup: false, note: 'On first asset load' },
  { plugin: 'expo-font', nativeAtStartup: false, note: 'useFonts() in _layout useEffect/render' },
  { plugin: 'expo-web-browser', nativeAtStartup: false, note: 'On first openAuthSession' },
  { plugin: 'expo-sqlite', nativeAtStartup: false, note: 'openDatabaseAsync in useEffect, not import' },
  { plugin: 'expo-image-picker', nativeAtStartup: false, note: 'On first launchCameraAsync' },
  { plugin: 'expo-system-ui', nativeAtStartup: false, note: 'Background color at runtime' }
];

const NOT_USED = ['expo-secure-store', 'react-native-mmkv', 'react-native-vision-camera', 'expo-splash-screen (explicit)'];

function readPackageJson() {
  return JSON.parse(fs.readFileSync(path.join(ROOT, 'package.json'), 'utf8'));
}

function listAutolinkedNativeModules() {
  const pkg = readPackageJson();
  const native = [];
  for (const [name, ver] of Object.entries(pkg.dependencies || {})) {
    if (!name.startsWith('expo-') && !name.startsWith('react-native')) continue;
    const pkgPath = path.join(ROOT, 'node_modules', name, 'package.json');
    if (!fs.existsSync(pkgPath)) continue;
    const meta = JSON.parse(fs.readFileSync(pkgPath, 'utf8'));
    if (meta.plugin || name.startsWith('expo-') || name.includes('gesture') || name.includes('reanimated') || name.includes('screens') || name.includes('safe-area') || name.includes('webview')) {
      native.push({ name, version: ver, hasExpoModule: fs.existsSync(path.join(ROOT, 'node_modules', name, 'android')) });
    }
  }
  return native;
}

console.log('# Expo Native + JS Startup Graph\n');
console.log('## 1. Native Android path\n');
for (const s of NATIVE_ANDROID.steps) {
  console.log(`${s.order}. \`${s.file}\` — ${s.action}`);
}

console.log('\n## 2. JS sync path (before first paint)\n');
console.log('```mermaid');
console.log('flowchart TD');
console.log('  A[expo-router/entry.js] --> B[entry-classic.js]');
console.log('  B --> C[@expo/metro-runtime]');
console.log('  B --> D[renderRootComponent]');
console.log('  D --> E[expo / Expo.fx]');
console.log('  E --> F[expo/src/winter/runtime.native.ts]');
console.log('  F --> G[InitializeCore / setUpXHR]');
console.log('  F --> H[installAbortSignalPatch]');
console.log('  H -->|AbortSignal undefined| X[FATAL JS CRASH]');
console.log('  D --> I[ExpoRoot / qualified-entry]');
console.log('  I --> J[app/_layout.tsx]');
console.log('  J --> K[SQLite / Fonts / Offline - useEffect]');
console.log('```\n');

for (const step of JS_SYNC.chain) {
  console.log(`${step.id}. **${step.module}**`);
  console.log(`   → ${Array.isArray(step.loads) ? step.loads.join(', ') : step.loads}`);
}

console.log('\n## 3. app.config.js plugins vs startup\n');
console.log('| Plugin | Native at cold start |');
console.log('|--------|---------------------|');
for (const p of APP_CONFIG_PLUGINS) {
  console.log(`| ${p.plugin} | ${p.nativeAtStartup ? 'YES' : 'NO'} — ${p.note} |`);
}

console.log('\n## 4. Not in project\n');
console.log(NOT_USED.map((x) => `- ${x}`).join('\n'));

console.log('\n## 5. Autolinked native-capable packages\n');
for (const m of listAutolinkedNativeModules()) {
  console.log(`- ${m.name}@${m.version}${m.hasExpoModule ? ' (android/)' : ''}`);
}

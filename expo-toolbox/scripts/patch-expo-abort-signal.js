'use strict';

/**
 * Patches Expo's winter runtime for Android release builds where bare `AbortSignal`
 * is undefined when installAbortSignalPatch runs (startup crash).
 */
const fs = require('fs');
const path = require('path');

function patchExpoWinterRuntime() {
  const expoRoot = path.dirname(require.resolve('expo/package.json'));
  const runtimeFile = path.join(expoRoot, 'src/winter/runtime.native.ts');
  const abortSignalFile = path.join(expoRoot, 'src/winter/AbortSignal.ts');

  let runtimeSource = fs.readFileSync(runtimeFile, 'utf8');
  const runtimeNeedle = 'installAbortSignalPatch(AbortSignal);';
  const runtimeReplacement = `installAbortSignalPatch(
  (globalThis as typeof globalThis & { AbortSignal?: typeof AbortSignal }).AbortSignal ??
    (global as typeof global & { AbortSignal?: typeof AbortSignal }).AbortSignal ??
    require('abort-controller/dist/abort-controller.js').AbortSignal
);`;

  if (runtimeSource.includes(runtimeNeedle)) {
    runtimeSource = runtimeSource.replace(runtimeNeedle, runtimeReplacement);
    fs.writeFileSync(runtimeFile, runtimeSource);
    console.log('[patch-expo-abort-signal] Patched expo/src/winter/runtime.native.ts');
  } else if (runtimeSource.includes('abort-controller/dist/abort-controller.js')) {
    console.log('[patch-expo-abort-signal] runtime.native.ts already patched');
  } else {
    console.warn('[patch-expo-abort-signal] runtime.native.ts patch target not found');
  }

  let abortSignalSource = fs.readFileSync(abortSignalFile, 'utf8');
  const guardNeedle = 'export function installAbortSignalPatch(';
  const guardReplacement = `export function installAbortSignalPatch(
  abortSignal: AbortSignalConstructor | null | undefined
): AbortSignalConstructor {
  if (abortSignal == null) {
    abortSignal =
      (globalThis as typeof globalThis & { AbortSignal?: AbortSignalConstructor }).AbortSignal ??
      (global as typeof global & { AbortSignal?: AbortSignalConstructor }).AbortSignal ??
      require('abort-controller/dist/abort-controller.js').AbortSignal;
    (globalThis as typeof globalThis & { AbortSignal?: AbortSignalConstructor }).AbortSignal =
      abortSignal;
  }`;

  if (
    abortSignalSource.includes(guardNeedle) &&
    !abortSignalSource.includes("abortSignal == null")
  ) {
    abortSignalSource = abortSignalSource.replace(
      `export function installAbortSignalPatch(
  abortSignal: AbortSignalConstructor
): AbortSignalConstructor {
  if (abortSignal.timeout == null) {`,
      `${guardReplacement}
  if (abortSignal.timeout == null) {`
    );
    fs.writeFileSync(abortSignalFile, abortSignalSource);
    console.log('[patch-expo-abort-signal] Patched expo/src/winter/AbortSignal.ts');
  } else if (abortSignalSource.includes("abortSignal == null")) {
    console.log('[patch-expo-abort-signal] AbortSignal.ts already patched');
  } else {
    console.warn('[patch-expo-abort-signal] AbortSignal.ts patch target not found');
  }
}

try {
  patchExpoWinterRuntime();
} catch (error) {
  console.error('[patch-expo-abort-signal] Failed:', error);
  process.exit(1);
}

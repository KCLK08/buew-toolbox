#!/usr/bin/env node
'use strict';

/**
 * Simulates the exact JS crash chain when abort-controller is bundled as Metro asset.
 * Run: node scripts/simulate-abortsignal-crash.js
 */

function polyfillGlobal(name, getValue) {
  let value;
  let valueSet = false;
  Object.defineProperty(globalThis, name, {
    get() {
      if (!valueSet) {
        valueSet = true;
        value = getValue();
        Object.defineProperty(globalThis, name, {
          value,
          configurable: true,
          enumerable: true,
          writable: true
        });
      }
      return value;
    },
    configurable: true,
    enumerable: true
  });
}

function registerAsset(meta) {
  return meta.hash;
}

function installAbortSignalPatch(abortSignal) {
  if (abortSignal.timeout == null) {
    Object.defineProperty(abortSignal, 'timeout', { value: () => {}, configurable: true });
  }
  return abortSignal;
}

const abortControllerJs = require('../node_modules/abort-controller/dist/abort-controller.js');

console.log('=== A: abort-controller as JS (workflow #24) ===');
delete globalThis.AbortSignal;
polyfillGlobal('AbortSignal', () => abortControllerJs.AbortSignal);
try {
  installAbortSignalPatch(globalThis.AbortSignal);
  console.log('OK');
} catch (e) {
  console.log('CRASH:', e.message);
}

console.log('\n=== B: abort-controller as Metro asset (PR #8b0866a) ===');
delete globalThis.AbortSignal;
polyfillGlobal('AbortSignal', () => {
  const assetId = registerAsset({ name: 'abort-controller', type: 'mjs' });
  return typeof assetId === 'object' ? assetId.AbortSignal : undefined;
});
try {
  installAbortSignalPatch(globalThis.AbortSignal);
  console.log('OK');
} catch (e) {
  console.log('CRASH:', e.message);
}

console.log('\n=== C: PR #55 null-guard patch ===');
delete globalThis.AbortSignal;
polyfillGlobal('AbortSignal', () => undefined);

function resolveAbortSignalCtor(abortSignal) {
  if (abortSignal != null) return abortSignal;
  const fromGlobal = globalThis.AbortSignal ?? global.AbortSignal;
  if (fromGlobal != null) return fromGlobal;
  const { AbortSignal } = abortControllerJs;
  globalThis.AbortSignal = AbortSignal;
  return AbortSignal;
}

try {
  installAbortSignalPatch(resolveAbortSignalCtor(undefined));
  console.log('OK');
} catch (e) {
  console.log('CRASH:', e.message);
}

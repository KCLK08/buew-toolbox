'use strict';

/**
 * Must run after InitializeCore and BEFORE expo/src/winter (see metro.config.js).
 * Expo winter calls installAbortSignalPatch(AbortSignal) — if AbortSignal is still
 * undefined, Android release builds crash with:
 *   Cannot read property 'timeout' of undefined
 */
const { AbortController, AbortSignal } = require('abort-controller/dist/abort-controller');

if (typeof global.AbortController === 'undefined') {
  global.AbortController = AbortController;
}

if (typeof global.AbortSignal === 'undefined') {
  global.AbortSignal = AbortSignal;
}

globalThis.AbortController = global.AbortController;
globalThis.AbortSignal = global.AbortSignal;

// Materialize lazy globals from InitializeCore if present.
void global.AbortController;
void global.AbortSignal;

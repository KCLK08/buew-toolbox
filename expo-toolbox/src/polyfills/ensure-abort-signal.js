/**
 * Ensures AbortController/AbortSignal exist before Expo's winter runtime boots.
 * Without this, installAbortSignalPatch(AbortSignal) can crash with:
 * "Cannot read property 'timeout' of undefined"
 */
require('react-native/Libraries/Core/InitializeCore');

const { AbortController, AbortSignal } = require('abort-controller/dist/abort-controller');

if (typeof globalThis.AbortController === 'undefined') {
  globalThis.AbortController = AbortController;
}

if (typeof globalThis.AbortSignal === 'undefined') {
  globalThis.AbortSignal = AbortSignal;
}

// Touch globals so lazy polyfills from InitializeCore are materialized.
void globalThis.AbortController;
void globalThis.AbortSignal;

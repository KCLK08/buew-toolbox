'use strict';

// Runs before expo-router entry via Metro getModulesRunBeforeMainModule.
// Ensures AbortSignal exists before expo/src/winter/runtime.native.ts executes.
const { AbortController, AbortSignal } = require('abort-controller/dist/abort-controller.js');

if (globalThis.AbortSignal == null) {
  globalThis.AbortController = AbortController;
  globalThis.AbortSignal = AbortSignal;
}

if (global.AbortSignal == null) {
  global.AbortController = AbortController;
  global.AbortSignal = AbortSignal;
}

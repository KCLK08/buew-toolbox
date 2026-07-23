'use strict';

const { installAbortSignalPatch: installOriginal } = require('expo/src/winter/AbortSignal');

function resolveAbortSignalCtor(passed) {
  if (passed != null) {
    return passed;
  }

  const fromGlobal = globalThis.AbortSignal ?? global.AbortSignal;
  if (fromGlobal != null) {
    return fromGlobal;
  }

  const { AbortController, AbortSignal } = require('abort-controller/dist/abort-controller');
  global.AbortController = AbortController;
  global.AbortSignal = AbortSignal;
  globalThis.AbortController = AbortController;
  globalThis.AbortSignal = AbortSignal;
  return AbortSignal;
}

function installAbortSignalPatch(abortSignal) {
  const resolved = resolveAbortSignalCtor(abortSignal);
  return installOriginal(resolved);
}

module.exports = {
  installAbortSignalPatch
};

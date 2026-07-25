'use strict';

// Patched expo/src/winter/runtime.native.ts — explicit AbortSignal before installAbortSignalPatch.
require('react-native/Libraries/Core/InitializeCore');

require('expo/types');

const { AbortController, AbortSignal } = require('abort-controller/dist/abort-controller.js');
globalThis.AbortController = AbortController;
globalThis.AbortSignal = AbortSignal;
global.AbortController = AbortController;
global.AbortSignal = AbortSignal;

const { installAbortSignalPatch } = require('expo/src/winter/AbortSignal');
const { installFormDataPatch } = require('expo/src/winter/FormData');
const { installGlobal: install } = require('expo/src/winter/installGlobal');

install('TextDecoder', () => require('expo/src/winter/TextDecoder').TextDecoder);
install('TextDecoderStream', () => require('expo/src/winter/TextDecoderStream').TextDecoderStream);
install('TextEncoderStream', () => require('expo/src/winter/TextDecoderStream').TextEncoderStream);
install('URL', () => require('expo/src/winter/url').URL);
install('URLSearchParams', () => require('expo/src/winter/url').URLSearchParams);
install('DOMException', () => require('expo/src/winter/DOMException').DOMException);
install('__ExpoImportMetaRegistry', () => require('expo/src/winter/ImportMetaRegistry').ImportMetaRegistry);
install('structuredClone', () => require('@ungap/structured-clone').default);

installFormDataPatch(FormData);
installAbortSignalPatch(AbortSignal);

Symbol.asyncIterator ??= Symbol.for('Symbol.asyncIterator');

const useRnFetch =
  process.env.EXPO_PUBLIC_USE_RN_FETCH === '1' || process.env.EXPO_PUBLIC_USE_RN_FETCH === 'true';

if (!useRnFetch) {
  if (!globalThis.Headers) {
    throw new Error(
      "expo/fetch expected `globalThis.Headers` to be installed by React Native's fetch polyfill."
    );
  }
  install('fetch', () => require('expo/src/winter/fetch').fetch);
}

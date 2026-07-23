'use strict';

const { AbortController, AbortSignal } = require('abort-controller/dist/abort-controller');

// Always install concrete implementations before Expo winter boots.
global.AbortController = AbortController;
global.AbortSignal = AbortSignal;
globalThis.AbortController = AbortController;
globalThis.AbortSignal = AbortSignal;

// Materialize lazy globals from InitializeCore if present.
void global.AbortController;
void global.AbortSignal;

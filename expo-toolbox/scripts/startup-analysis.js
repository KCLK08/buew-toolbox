#!/usr/bin/env node
'use strict';

const Metro = require('metro');
const fs = require('fs');
const path = require('path');

const PROJECT_ROOT = path.join(__dirname, '..');
const ENTRY = 'node_modules/expo-router/entry.js';

const STARTUP_CHAIN = [
  { id: 'entry', file: 'expo-router/entry.js', role: 'package.json main' },
  { id: 'entry-classic', file: 'expo-router/entry-classic.js', role: 'Metro runtime + renderRootComponent' },
  { id: 'metro-runtime', file: '@expo/metro-runtime', role: 'Fast Refresh / runtime hooks' },
  { id: 'qualified-entry', file: 'expo-router/build/qualified-entry.js', role: 'App shell + _ctx' },
  { id: 'expo-router-entry-side-effects', file: 'expo-router/build/entry.js side effects', role: 'Loads expo / react-native' },
  { id: 'expo-fx', file: 'expo/src/Expo.fx.tsx', role: 'import ./winter (CRASH ZONE)' },
  { id: 'winter-index', file: 'expo/src/winter/index.ts', role: 'import ./runtime' },
  { id: 'winter-runtime', file: 'expo/src/winter/runtime.native.ts', role: 'InitializeCore + installAbortSignalPatch' },
  { id: 'initialize-core', file: 'react-native/Libraries/Core/InitializeCore.js', role: 'setUpXHR → AbortSignal polyfill' },
  { id: 'abort-signal-patch', file: 'expo/src/winter/AbortSignal.ts', role: 'installAbortSignalPatch(AbortSignal)' },
  { id: 'app-layout', file: 'app/_layout.tsx', role: 'First React root layout (after entry tree)' }
];

async function buildBundle(configPath, outFile) {
  process.chdir(PROJECT_ROOT);
  const absConfig = path.resolve(PROJECT_ROOT, configPath);
  delete require.cache[absConfig];
  const config = require(absConfig);
  await Metro.runBuild(config, {
    entry: ENTRY,
    out: outFile,
    platform: 'android',
    dev: false,
    minify: false
  });
  return fs.readFileSync(outFile, 'utf8');
}

function analyze(content, label) {
  const hasWinterCall = /installAbortSignalPatch\)\(AbortSignal\)/.test(content);
  const hasResolvedCall = content.includes('installAbortSignalPatch(resolveAbortSignalCtor())');
  const hasNullGuard = content.includes('resolveAbortSignalCtor(abortSignal)');
  const registerAssetAbort = /registerAsset\(\{[^}]*abort-controller/i.test(content);

  const acIsJs = /AbortController = \/\*#__PURE__\*\/function \(\)/.test(content);
  const acIsAsset = /registerAsset\(\{[^}]*name: ['"]abort-controller/i.test(content);

  let acType = 'unknown';
  if (acIsJs && !registerAssetAbort) acType = 'JS module';
  if (registerAssetAbort || acIsAsset) acType = 'ASSET (broken)';

  const originalCrash = (() => {
    try {
      // eslint-disable-next-line no-new-func
      const fn = new Function(`
        function installAbortSignalPatch(abortSignal) {
          if (abortSignal.timeout == null) {}
          return abortSignal;
        }
        installAbortSignalPatch(undefined);
      `);
      fn();
      return false;
    } catch {
      return true;
    }
  })();

  return { label, hasWinterCall, hasResolvedCall, hasNullGuard, acType, originalCrash };
}

async function main() {
  const configs = [
    ['metro.config.w24-test.js', 'Workflow #24'],
    ['metro.config.broken-test.js', 'PR #33 broken (mjs assetExts only)'],
    ['metro.config.main-test.js', 'Main PR #54'],
    ['metro.config.js', 'Fix PR #55']
  ];

  console.log('# Startup chain (sync, before first React render)\n');
  for (const step of STARTUP_CHAIN) {
    console.log(`- ${step.id}: \`${step.file}\` — ${step.role}`);
  }

  console.log('\n# Bundle verification\n');
  console.log('| Config | abort-controller | installAbortSignalPatch(AbortSignal) | null guard in patch | original patch crashes on undefined |');
  console.log('|--------|------------------|--------------------------------------|---------------------|-------------------------------------|');

  for (const [cfg, label] of configs) {
    const out = path.join('/tmp', `startup-${path.basename(cfg, '.js')}.js`);
    try {
      const content = await buildBundle(cfg, out);
      const r = analyze(content, label);
      console.log(
        `| ${label} | ${r.acType} | ${r.hasWinterCall ? 'yes' : 'no'} | ${r.hasNullGuard ? 'yes' : 'no'} | ${r.originalCrash ? 'yes' : 'no'} |`
      );
    } catch (err) {
      console.log(`| ${label} | BUILD FAILED | | | ${err.message.slice(0, 60)} |`);
    }
  }
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});

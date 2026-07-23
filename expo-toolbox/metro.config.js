const path = require('path');
const { getDefaultConfig } = require('expo/metro-config');

const projectRoot = __dirname;
const workspaceRoot = path.resolve(projectRoot, '..');
const sharedRoot = path.resolve(workspaceRoot, 'shared');
const abortSignalPolyfill = path.resolve(projectRoot, 'src/polyfills/ensure-abort-signal.js');

/** @type {import('expo/metro-config').MetroConfig} */
const config = getDefaultConfig(projectRoot);

config.watchFolders = [...(config.watchFolders || []), workspaceRoot, sharedRoot];

config.resolver = {
  ...config.resolver,
  assetExts: [...(config.resolver.assetExts || []), 'mjs'],
  sourceExts: [...(config.resolver.sourceExts || []), 'mjs'],
  nodeModulesPaths: [
    path.resolve(projectRoot, 'node_modules'),
    path.resolve(workspaceRoot, 'node_modules')
  ],
  extraNodeModules: {
    ...(config.resolver?.extraNodeModules || {}),
    '@buew/shared': sharedRoot
  }
};

const defaultGetModulesRunBeforeMainModule = config.serializer?.getModulesRunBeforeMainModule;

config.serializer = {
  ...config.serializer,
  getModulesRunBeforeMainModule: () => {
    const defaults = defaultGetModulesRunBeforeMainModule?.() ?? [];
    const winterIndex = defaults.findIndex((modulePath) =>
      modulePath.includes(`${path.sep}expo${path.sep}src${path.sep}winter`)
    );

    if (winterIndex >= 0) {
      const ordered = [...defaults];
      ordered.splice(winterIndex, 0, abortSignalPolyfill);
      return ordered;
    }

    return [abortSignalPolyfill, ...defaults];
  }
};

module.exports = config;

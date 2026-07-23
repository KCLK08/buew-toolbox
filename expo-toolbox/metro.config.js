const path = require('path');
const { resolve } = require('metro-resolver');
const { getDefaultConfig } = require('expo/metro-config');

const projectRoot = __dirname;
const workspaceRoot = path.resolve(projectRoot, '..');
const sharedRoot = path.resolve(workspaceRoot, 'shared');
const patchedWinterRuntime = path.resolve(projectRoot, 'src/polyfills/expo-winter-runtime.native.js');
const patchedWinterIndex = path.resolve(projectRoot, 'src/polyfills/expo-winter-index.js');
const abortControllerJs = path.resolve(
  projectRoot,
  'node_modules/abort-controller/dist/abort-controller.js'
);

function targetsAbortController(moduleName) {
  return (
    moduleName === 'abort-controller' ||
    moduleName.startsWith('abort-controller/') ||
    moduleName.endsWith('/abort-controller/dist/abort-controller') ||
    moduleName.endsWith('/abort-controller/dist/abort-controller.js') ||
    moduleName.endsWith('/abort-controller/dist/abort-controller.mjs')
  );
}

/** @type {import('expo/metro-config').MetroConfig} */
const config = getDefaultConfig(projectRoot);

config.watchFolders = [...(config.watchFolders || []), workspaceRoot, sharedRoot];

const defaultResolveRequest = config.resolver.resolveRequest;

function isExpoWinterModule(modulePath) {
  return modulePath.includes(`${path.sep}expo${path.sep}src${path.sep}winter`);
}

config.resolver = {
  ...config.resolver,
  // .mjs must stay an asset so pdfjs-dist is not fully bundled (Hermes rejects its Node imports).
  assetExts: [...(config.resolver.assetExts || []), 'mjs'],
  sourceExts: [...(config.resolver.sourceExts || []), 'mjs'],
  nodeModulesPaths: [
    path.resolve(projectRoot, 'node_modules'),
    path.resolve(workspaceRoot, 'node_modules')
  ],
  extraNodeModules: {
    ...(config.resolver?.extraNodeModules || {}),
    '@buew/shared': sharedRoot
  },
  resolveRequest: (context, moduleName, platform) => {
    const fromExpoFx = context.originModulePath?.includes(`${path.sep}expo${path.sep}src${path.sep}Expo.fx`);
    const targetsRelativeWinter = fromExpoFx && moduleName === './winter';

    const targetsWinterIndex =
      moduleName === 'expo/src/winter' ||
      moduleName === 'expo/src/winter/index' ||
      moduleName.endsWith('/expo/src/winter/index.ts') ||
      moduleName.endsWith('/expo/src/winter/index.js');

    const fromWinterIndex = context.originModulePath?.includes(
      `${path.sep}expo${path.sep}src${path.sep}winter${path.sep}index`
    );
    const targetsWinterRuntimeImport = fromWinterIndex && moduleName === './runtime' && platform !== 'web';

    const targetsWinterRuntime =
      moduleName.endsWith('/expo/src/winter/runtime.native') ||
      moduleName.endsWith('/expo/src/winter/runtime.native.ts') ||
      moduleName.endsWith('/expo/src/winter/runtime.native.js');

    if (targetsWinterIndex || targetsRelativeWinter || targetsWinterRuntimeImport || targetsWinterRuntime) {
      return {
        filePath: targetsWinterIndex ? patchedWinterIndex : patchedWinterRuntime,
        type: 'sourceFile'
      };
    }

    if (targetsAbortController(moduleName)) {
      return {
        filePath: abortControllerJs,
        type: 'sourceFile'
      };
    }

    if (defaultResolveRequest) {
      return defaultResolveRequest(context, moduleName, platform);
    }

    return resolve(context, moduleName, platform);
  }
};

const defaultGetModulesRunBeforeMainModule = config.serializer?.getModulesRunBeforeMainModule;

config.serializer = {
  ...config.serializer,
  getModulesRunBeforeMainModule: () => {
    const defaults = defaultGetModulesRunBeforeMainModule?.() ?? [];
    const replaced = defaults.map((modulePath) =>
      isExpoWinterModule(modulePath) ? patchedWinterRuntime : modulePath
    );
    return [...new Set(replaced)];
  }
};

module.exports = config;

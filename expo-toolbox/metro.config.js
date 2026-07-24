const path = require('path');
const { resolve } = require('metro-resolver');
const { getDefaultConfig } = require('expo/metro-config');

const projectRoot = __dirname;
const workspaceRoot = path.resolve(projectRoot, '..');
const sharedRoot = path.resolve(workspaceRoot, 'shared');
const abortControllerJs = path.resolve(
  projectRoot,
  'node_modules/abort-controller/dist/abort-controller.js'
);
const patchedAbortSignal = path.resolve(projectRoot, 'src/polyfills/expo-winter-AbortSignal.js');

/** @type {import('expo/metro-config').MetroConfig} */
const config = getDefaultConfig(projectRoot);

config.watchFolders = [...(config.watchFolders || []), workspaceRoot, sharedRoot];

const defaultResolveRequest = config.resolver.resolveRequest;

function targetsAbortController(moduleName) {
  return (
    moduleName === 'abort-controller' ||
    moduleName.startsWith('abort-controller/') ||
    moduleName.endsWith('/abort-controller/dist/abort-controller') ||
    moduleName.endsWith('/abort-controller/dist/abort-controller.js') ||
    moduleName.endsWith('/abort-controller/dist/abort-controller.mjs')
  );
}

function targetsExpoAbortSignal(context, moduleName) {
  const fromExpoWinter = context.originModulePath?.includes(`${path.sep}expo${path.sep}src${path.sep}winter`);
  return (
    moduleName === 'expo/src/winter/AbortSignal' ||
    moduleName.endsWith('/expo/src/winter/AbortSignal.ts') ||
    moduleName.endsWith('/expo/src/winter/AbortSignal.js') ||
    (fromExpoWinter && (moduleName === './AbortSignal' || moduleName === '../AbortSignal'))
  );
}

function targetsPdfJsMjs(moduleName) {
  return moduleName.includes('pdfjs-dist') && moduleName.endsWith('.mjs');
}

config.resolver = {
  ...config.resolver,
  // Keep local pdf.js worker assets addressable; do not treat every .mjs as source.
  assetExts: [...(config.resolver.assetExts || []), 'mjs'],
  nodeModulesPaths: [
    path.resolve(projectRoot, 'node_modules'),
    path.resolve(workspaceRoot, 'node_modules')
  ],
  extraNodeModules: {
    ...(config.resolver?.extraNodeModules || {}),
    '@buew/shared': sharedRoot
  },
  resolveRequest: (context, moduleName, platform) => {
    if (targetsAbortController(moduleName)) {
      return {
        filePath: abortControllerJs,
        type: 'sourceFile'
      };
    }

    if (targetsExpoAbortSignal(context, moduleName)) {
      return {
        filePath: patchedAbortSignal,
        type: 'sourceFile'
      };
    }

    if (targetsPdfJsMjs(moduleName)) {
      const resolved = defaultResolveRequest
        ? defaultResolveRequest(context, moduleName, platform)
        : resolve(context, moduleName, platform);

      if (resolved?.type === 'sourceFile') {
        return {
          type: 'assetFiles',
          filePaths: [resolved.filePath]
        };
      }

      return resolved;
    }

    if (defaultResolveRequest) {
      return defaultResolveRequest(context, moduleName, platform);
    }

    return resolve(context, moduleName, platform);
  }
};

module.exports = config;

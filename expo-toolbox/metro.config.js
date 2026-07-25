const path = require('path');
const { resolve } = require('metro-resolver');
const { getDefaultConfig } = require('expo/metro-config');

const projectRoot = __dirname;
const sharedRoot = path.resolve(projectRoot, '../shared');
const abortControllerJs = path.resolve(
  projectRoot,
  'node_modules/abort-controller/dist/abort-controller.js'
);
const edgeToEdgeJs = path.resolve(
  projectRoot,
  'node_modules/react-native-is-edge-to-edge/dist/index.js'
);

/** @type {import('expo/metro-config').MetroConfig} */
const config = getDefaultConfig(projectRoot);

config.watchFolders = [...(config.watchFolders || []), sharedRoot];

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

function targetsPdfJsMjs(moduleName) {
  return moduleName.includes('pdfjs-dist') && moduleName.endsWith('.mjs');
}

function targetsEdgeToEdge(moduleName) {
  return (
    moduleName === 'react-native-is-edge-to-edge' ||
    moduleName.startsWith('react-native-is-edge-to-edge/') ||
    moduleName.endsWith('/react-native-is-edge-to-edge/dist/index') ||
    moduleName.endsWith('/react-native-is-edge-to-edge/dist/index.js') ||
    moduleName.endsWith('/react-native-is-edge-to-edge/dist/index.mjs')
  );
}

config.resolver = {
  ...config.resolver,
  assetExts: [...(config.resolver.assetExts || []), 'mjs'],
  nodeModulesPaths: [path.resolve(projectRoot, 'node_modules')],
  extraNodeModules: {
    ...(config.resolver?.extraNodeModules || {}),
    '@buew/shared': sharedRoot,
    expo: path.resolve(projectRoot, 'node_modules/expo'),
    react: path.resolve(projectRoot, 'node_modules/react'),
    'react-native': path.resolve(projectRoot, 'node_modules/react-native')
  },
  resolveRequest: (context, moduleName, platform) => {
    if (targetsAbortController(moduleName)) {
      return {
        filePath: abortControllerJs,
        type: 'sourceFile'
      };
    }

    if (targetsEdgeToEdge(moduleName)) {
      return {
        filePath: edgeToEdgeJs,
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

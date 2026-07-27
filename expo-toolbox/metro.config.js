const path = require('path');
const { resolve } = require('metro-resolver');
const { getDefaultConfig } = require('expo/metro-config');

const projectRoot = __dirname;
const sharedRoot = path.resolve(projectRoot, '../shared');
const edgeToEdgeJs = path.resolve(
  projectRoot,
  'node_modules/react-native-is-edge-to-edge/dist/index.js'
);
const pdfWorkerAsset = path.resolve(projectRoot, 'assets/pdf.worker.min.mjs');
const pdfPreviewCoreAsset = path.resolve(projectRoot, 'assets/pdfjs/pdf.min.js');
const pdfPreviewWorkerAsset = path.resolve(projectRoot, 'assets/pdfjs/pdf.worker.min.js');

/** @type {import('expo/metro-config').MetroConfig} */
const config = getDefaultConfig(projectRoot);

config.watchFolders = [...(config.watchFolders || []), sharedRoot];

const defaultResolveRequest = config.resolver.resolveRequest;

function targetsPdfWorkerAsset(moduleName) {
  return (
    moduleName.endsWith('/assets/pdf.worker.min.mjs') ||
    moduleName.endsWith('assets/pdf.worker.min.mjs') ||
    (moduleName.includes('pdfjs-dist') && moduleName.includes('worker') && moduleName.endsWith('.mjs'))
  );
}

function targetsPdfPreviewAssets(moduleName) {
  return (
    moduleName.endsWith('/assets/pdfjs/pdf.min.js') ||
    moduleName.endsWith('assets/pdfjs/pdf.min.js') ||
    moduleName.endsWith('/assets/pdfjs/pdf.worker.min.js') ||
    moduleName.endsWith('assets/pdfjs/pdf.worker.min.js')
  );
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
  nodeModulesPaths: [path.resolve(projectRoot, 'node_modules')],
  extraNodeModules: {
    ...(config.resolver?.extraNodeModules || {}),
    '@buew/shared': sharedRoot,
    expo: path.resolve(projectRoot, 'node_modules/expo'),
    react: path.resolve(projectRoot, 'node_modules/react'),
    'react-native': path.resolve(projectRoot, 'node_modules/react-native')
  },
  resolveRequest: (context, moduleName, platform) => {
    if (targetsEdgeToEdge(moduleName)) {
      return {
        filePath: edgeToEdgeJs,
        type: 'sourceFile'
      };
    }

    if (targetsPdfWorkerAsset(moduleName)) {
      const resolved = defaultResolveRequest
        ? defaultResolveRequest(context, moduleName, platform)
        : resolve(context, moduleName, platform);

      const filePath =
        resolved?.type === 'sourceFile'
          ? resolved.filePath
          : moduleName.includes('pdf.worker.min.mjs')
            ? pdfWorkerAsset
            : null;

      if (filePath) {
        return {
          type: 'assetFiles',
          filePaths: [filePath]
        };
      }

      return resolved;
    }

    if (targetsPdfPreviewAssets(moduleName)) {
      const filePath = moduleName.includes('pdf.worker.min.js')
        ? pdfPreviewWorkerAsset
        : pdfPreviewCoreAsset;

      return {
        type: 'assetFiles',
        filePaths: [filePath]
      };
    }

    if (defaultResolveRequest) {
      return defaultResolveRequest(context, moduleName, platform);
    }

    return resolve(context, moduleName, platform);
  }
};

module.exports = config;

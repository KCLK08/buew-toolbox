const path = require('path');
const { getDefaultConfig } = require('expo/metro-config');

/** Broken PR #33 config: mjs in assetExts, no abort-controller alias */
const projectRoot = __dirname;
const workspaceRoot = path.resolve(projectRoot, '..');
const sharedRoot = path.resolve(workspaceRoot, 'shared');
const config = getDefaultConfig(projectRoot);
config.watchFolders = [...(config.watchFolders || []), workspaceRoot, sharedRoot];
config.resolver = {
  ...config.resolver,
  assetExts: [...(config.resolver.assetExts || []), 'mjs'],
  nodeModulesPaths: [
    path.resolve(projectRoot, 'node_modules'),
    path.resolve(workspaceRoot, 'node_modules')
  ],
  extraNodeModules: { ...(config.resolver?.extraNodeModules || {}), '@buew/shared': sharedRoot }
};
module.exports = config;

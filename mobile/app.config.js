/** @type {import('expo/config').ConfigContext} */
module.exports = ({ config }) => ({
  ...config,
  name: 'BÜW-Toolbox',
  slug: 'buew-toolbox',
  version: '1.0.0',
  orientation: 'portrait',
  icon: './assets/icon.png',
  userInterfaceStyle: 'light',
  scheme: 'buew-toolbox',
  backgroundColor: '#f3efe7',
  splash: {
    image: './assets/splash-icon.png',
    resizeMode: 'contain',
    backgroundColor: '#f3efe7'
  },
  ios: {
    supportsTablet: true,
    bundleIdentifier: 'de.buew.toolbox'
  },
  android: {
    adaptiveIcon: {
      backgroundColor: '#f3efe7',
      foregroundImage: './assets/android-icon-foreground.png',
      backgroundImage: './assets/android-icon-background.png',
      monochromeImage: './assets/android-icon-monochrome.png'
    },
    package: 'de.buew.toolbox',
    predictiveBackGestureEnabled: false
  },
  web: {
    favicon: './assets/favicon.png',
    bundler: 'metro'
  },
  plugins: [
    'expo-router',
    'expo-font',
    'expo-web-browser',
    'expo-secure-store',
    'expo-local-authentication',
    [
      'expo-system-ui',
      {
        backgroundColor: '#f3efe7'
      }
    ]
  ],
  experiments: {
    typedRoutes: true
  },
  extra: {
    supabaseUrl: process.env.EXPO_PUBLIC_SUPABASE_URL ?? '',
    supabaseAnonKey: process.env.EXPO_PUBLIC_SUPABASE_ANON_KEY ?? '',
    toolboxWebBaseUrl:
      process.env.EXPO_PUBLIC_TOOLBOX_WEB_BASE_URL ?? 'https://kclk08.github.io/buew-toolbox',
    eas: {
      projectId: process.env.EAS_PROJECT_ID
    }
  }
});

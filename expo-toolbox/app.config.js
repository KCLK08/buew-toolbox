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
  backgroundColor: '#F2F0EB',
  splash: {
    image: './assets/splash-icon.png',
    resizeMode: 'contain',
    backgroundColor: '#F2F0EB'
  },
  ios: {
    supportsTablet: true,
    bundleIdentifier: 'de.buew.toolbox',
    infoPlist: {
      NSCameraUsageDescription: 'Kamera wird für die Fotodokumentation auf der Baustelle benötigt.'
    }
  },
  android: {
    adaptiveIcon: {
      backgroundColor: '#F2F0EB',
      foregroundImage: './assets/android-icon-foreground.png',
      backgroundImage: './assets/android-icon-background.png',
      monochromeImage: './assets/android-icon-monochrome.png'
    },
    package: 'de.buew.toolbox',
    predictiveBackGestureEnabled: false,
    permissions: ['CAMERA']
  },
  web: {
    favicon: './assets/favicon.png',
    bundler: 'metro'
  },
  plugins: [
    'expo-router',
    'expo-font',
    'expo-web-browser',
    'expo-sqlite',
    [
      'expo-image-picker',
      {
        cameraPermission: 'Erlaube BÜW-Toolbox den Kamerazugriff für die Fotodokumentation.'
      }
    ],
    [
      'expo-system-ui',
      {
        backgroundColor: '#F2F0EB'
      }
    ]
  ],
  experiments: {
    typedRoutes: true
  },
  extra: {
    toolboxWebBaseUrl:
      process.env.EXPO_PUBLIC_TOOLBOX_WEB_BASE_URL ?? 'https://kclk08.github.io/buew-toolbox',
    eas: {
      projectId: process.env.EAS_PROJECT_ID
    }
  }
});

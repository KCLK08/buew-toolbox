import Constants from 'expo-constants';

type ExtraConfig = {
  toolboxWebBaseUrl?: string;
};

const extra = (Constants.expoConfig?.extra ?? {}) as ExtraConfig;

export const appConfig = {
  toolboxWebBaseUrl:
    process.env.EXPO_PUBLIC_TOOLBOX_WEB_BASE_URL ||
    extra.toolboxWebBaseUrl ||
    'https://kclk08.github.io/buew-toolbox'
};

export function toolWebUrl(webPath: string): string {
  const base = appConfig.toolboxWebBaseUrl.replace(/\/$/, '');
  const path = webPath.startsWith('/') ? webPath : `/${webPath}`;
  return `${base}${path}`;
}

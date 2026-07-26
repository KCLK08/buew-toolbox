import type { ImageSourcePropType } from 'react-native';

export type ToolId = 'sitereport' | 'bautagebuch';

export type ToolboxTool = {
  id: ToolId;
  title: string;
  description: string;
  href: '/sitereport' | '/bautagebuch';
  tabHref: '/sitereport' | '/bautagebuch';
  webPath: string;
  icon: ImageSourcePropType;
  iconAlt: string;
};

export const TOOLBOX_TOOLS: ToolboxTool[] = [
  {
    id: 'sitereport',
    title: 'SiteReport',
    description: 'Foto‑basierte Protokolle mit XLSX‑Export.',
    href: '/sitereport',
    tabHref: '/sitereport',
    webPath: '/sitereport/',
    icon: require('../../assets/icons/sitereport.png'),
    iconAlt: 'SiteReport Icon'
  },
  {
    id: 'bautagebuch',
    title: 'Bautagebuch',
    description: 'AcroForm-Tool mit Auto-Erkennung, Guided Flow und sauberem PDF-Export.',
    href: '/bautagebuch',
    tabHref: '/bautagebuch',
    webPath: '/bautagebuch/',
    icon: require('../../assets/icons/bautagebuch.png'),
    iconAlt: 'Bautagebuch Icon'
  }
];

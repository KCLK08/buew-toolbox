import type { ImageSourcePropType } from 'react-native';

export type ToolId = 'sitereport' | 'bautagebuch';

export type ToolboxTool = {
  id: ToolId;
  title: string;
  description: string;
  href: `/${ToolId}`;
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
    webPath: '/sitereport/',
    icon: require('../../assets/icons/sitereport.png'),
    iconAlt: 'SiteReport Icon'
  },
  {
    id: 'bautagebuch',
    title: 'Bautagebuch',
    description: 'AcroForm-Tool mit Auto-Erkennung, Guided Flow und sauberem PDF-Export.',
    href: '/bautagebuch',
    webPath: '/bautagebuch/',
    icon: require('../../assets/icons/bautagebuch.png'),
    iconAlt: 'Bautagebuch Icon'
  }
];

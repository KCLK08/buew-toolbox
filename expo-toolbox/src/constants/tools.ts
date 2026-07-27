import type { ImageSourcePropType } from 'react-native';

export type ToolId = 'sitereport' | 'bautagebuch';

export type ToolboxTool = {
  id: ToolId;
  title: string;
  description: string;
  features: string[];
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
    description:
      'Baustellenprotokolle mit Fotos, strukturierten Einträgen und professionellem Export für die Bauüberwachung.',
    features: ['Protokolle erstellen', 'Fotos dokumentieren', 'PDF- und Excel-Export'],
    href: '/sitereport',
    tabHref: '/sitereport',
    webPath: '/sitereport/',
    icon: require('../../assets/icons/sitereport.png'),
    iconAlt: 'SiteReport Icon'
  },
  {
    id: 'bautagebuch',
    title: 'Bautagebuch',
    description:
      'Digitale Erfassung, Dokumentation und Verwaltung von Baustellenereignissen, Fotos und Leistungen.',
    features: ['Protokolle erstellen', 'Fotos dokumentieren', 'PDF/XLSX Export'],
    href: '/bautagebuch',
    tabHref: '/bautagebuch',
    webPath: '/bautagebuch/',
    icon: require('../../assets/icons/bautagebuch.png'),
    iconAlt: 'Bautagebuch Icon'
  }
];

export const colors = {
  bg: '#f3efe7',
  ink: '#161614',
  muted: '#5f5a52',
  accent: '#d3543c',
  accent2: '#1d2b36',
  panel: '#fffaf2',
  border: '#e6ddcf',
  glow: 'rgba(211, 84, 60, 0.22)',
  white: '#ffffff',
  danger: '#b42318',
  success: '#0f7b4a',
  gradientWarm: '#fbe3c6',
  gradientCool: '#cfe2f5',
  gradientTop: '#f9f2e7',
  gradientBottom: '#f3efe7'
} as const;

export const AUTH_COPY = {
  appName: 'BÜW-Toolbox',
  loginTitle: 'Anmelden',
  registerTitle: 'Registrieren',
  forgotTitle: 'Passwort vergessen',
  loginSubtitle: 'Melde dich an, um auf die Toolbox zuzugreifen.',
  registerSubtitle: 'Erstelle ein Konto für die BÜW-Toolbox.',
  forgotSubtitle: 'Wir senden dir einen Link zum Zurücksetzen deines Passworts.',
  offline: 'Keine Internetverbindung. Bitte später erneut versuchen.',
  biometricPrompt: 'Möchtest du die biometrische Anmeldung aktivieren?',
  biometricEnable: 'Aktivieren',
  biometricSkip: 'Später',
  logout: 'Abmelden'
} as const;

export const PUBLIC_AUTH_PATHS = ['/login', '/register', '/forgot-password'] as const;

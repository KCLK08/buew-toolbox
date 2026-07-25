/**
 * Mobile design tokens for the Expo shell.
 * Light mode is default; `dark` tokens prepare future dark mode without enabling it.
 */

export const colors = {
  bg: '#F2F0EB',
  ink: '#1A1916',
  muted: '#6B6560',
  accent: '#C44B32',
  accentPressed: '#A33D29',
  accent2: '#243240',
  panel: '#FFFCF7',
  panelElevated: '#FFFFFF',
  border: '#E4DDD2',
  borderStrong: '#D0C7B8',
  success: '#2F6B45',
  warning: '#9A6B12',
  danger: '#A12C24',
  info: '#2A5F8F',
  white: '#FFFFFF',
  overlay: 'rgba(26, 25, 22, 0.45)',
  tabInactive: '#8A837A',
  tabActive: '#C44B32',
  badgeBg: 'rgba(196, 75, 50, 0.12)',
  glow: 'rgba(196, 75, 50, 0.18)',
  gradientWarm: '#F7E4C8',
  gradientCool: '#D5E4F2',
  gradientTop: '#F7F2E9',
  gradientBottom: '#F2F0EB'
} as const;

/** Prepared for dark mode — not applied yet. */
export const darkColors = {
  bg: '#121417',
  ink: '#F4F1EA',
  muted: '#A39E95',
  accent: '#E06A52',
  accentPressed: '#C44B32',
  accent2: '#C9D3DC',
  panel: '#1C2026',
  panelElevated: '#252A32',
  border: '#2F3640',
  borderStrong: '#3D4652',
  success: '#5FBF84',
  warning: '#E0B34A',
  danger: '#F07167',
  info: '#6FA8D6',
  white: '#FFFFFF',
  overlay: 'rgba(0, 0, 0, 0.55)',
  tabInactive: '#8B939E',
  tabActive: '#E06A52',
  badgeBg: 'rgba(224, 106, 82, 0.18)',
  glow: 'rgba(224, 106, 82, 0.22)',
  gradientWarm: '#2A2218',
  gradientCool: '#18222C',
  gradientTop: '#171A1F',
  gradientBottom: '#121417'
} as const;

export const spacing = {
  xxs: 4,
  xs: 8,
  sm: 12,
  md: 16,
  lg: 20,
  xl: 24,
  xxl: 32,
  pageX: 16,
  pageTop: 12,
  pageBottom: 28,
  sectionGap: 20,
  cardGap: 12,
  cardPadding: 16,
  cardRadius: 16,
  inputRadius: 14,
  buttonRadius: 14,
  touchMin: 44,
  fabSize: 56,
  tabBarBody: 56,
  iconSize: 28,
  heroGap: 16,
  iconRadius: 12
} as const;

export const typography = {
  display: {
    fontFamily: 'SpaceGrotesk_700Bold',
    fontSize: 28,
    lineHeight: 34,
    letterSpacing: -0.4
  },
  title: {
    fontFamily: 'SpaceGrotesk_700Bold',
    fontSize: 22,
    lineHeight: 28,
    letterSpacing: -0.3
  },
  subtitle: {
    fontFamily: 'SpaceGrotesk_600SemiBold',
    fontSize: 17,
    lineHeight: 24
  },
  body: {
    fontFamily: 'SpaceGrotesk_400Regular',
    fontSize: 16,
    lineHeight: 22
  },
  bodyStrong: {
    fontFamily: 'SpaceGrotesk_600SemiBold',
    fontSize: 16,
    lineHeight: 22
  },
  caption: {
    fontFamily: 'SpaceGrotesk_400Regular',
    fontSize: 13,
    lineHeight: 18
  },
  label: {
    fontFamily: 'SpaceGrotesk_600SemiBold',
    fontSize: 13,
    lineHeight: 18
  },
  button: {
    fontFamily: 'SpaceGrotesk_700Bold',
    fontSize: 16,
    lineHeight: 20
  }
} as const;

export const shadows = {
  card: {
    shadowColor: '#1A1916',
    shadowOpacity: 0.08,
    shadowRadius: 10,
    shadowOffset: { width: 0, height: 4 },
    elevation: 2
  },
  fab: {
    shadowColor: '#1A1916',
    shadowOpacity: 0.2,
    shadowRadius: 12,
    shadowOffset: { width: 0, height: 6 },
    elevation: 5
  }
} as const;

import { spacing } from '../../constants/theme';

export type HomeLayoutTier = 'relaxed' | 'compact' | 'dense';

const TIER_THRESHOLDS: Record<Exclude<HomeLayoutTier, 'dense'>, number> = {
  relaxed: 700,
  compact: 560
};

export function resolveHomeLayoutTier(availableHeight: number): HomeLayoutTier {
  if (availableHeight >= TIER_THRESHOLDS.relaxed) return 'relaxed';
  if (availableHeight >= TIER_THRESHOLDS.compact) return 'compact';
  return 'dense';
}

export function homeCardsGap(tier: HomeLayoutTier): number {
  if (tier === 'dense') return spacing.xs;
  if (tier === 'compact') return spacing.sm;
  return spacing.md;
}

export function homeHeaderBottomGap(tier: HomeLayoutTier): number {
  if (tier === 'dense') return spacing.xs;
  if (tier === 'compact') return spacing.sm;
  return spacing.md;
}

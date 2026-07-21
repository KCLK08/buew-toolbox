import { colors } from '@buew/shared';

export { colors };

export const spacing = {
  pageX: 20,
  pageTop: 48,
  pageBottom: 96,
  heroGap: 24,
  cardGap: 18,
  cardPadding: 18,
  iconSize: 46,
  iconRadius: 14,
  cardRadius: 18
} as const;

export { TOOLBOX_TOOLS, type ToolId, type ToolboxTool } from './tools';

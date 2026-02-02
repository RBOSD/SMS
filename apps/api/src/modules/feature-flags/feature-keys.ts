export const FEATURE_KEYS = [
  'module_issues',
  'module_plans',
  'module_ai_review',
] as const;

export type FeatureKey = (typeof FEATURE_KEYS)[number];

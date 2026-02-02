export const FeatureFlagKeys = [
  'module_issues',
  'module_plans',
  'module_ai_review',
] as const;

export type FeatureFlagKey = (typeof FeatureFlagKeys)[number];

export type Role = 'admin' | 'manager' | 'viewer';

export interface FeatureFlagsEffective {
  module_issues: boolean;
  module_plans: boolean;
  module_ai_review: boolean;
}


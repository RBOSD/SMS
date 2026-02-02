import { SetMetadata } from '@nestjs/common';

export const FEATURE_FLAG_KEY = 'featureFlagKey';
export const RequireFeature = (key: string) =>
  SetMetadata(FEATURE_FLAG_KEY, key);

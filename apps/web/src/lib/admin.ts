import { apiFetch, type FeatureFlagsEffective } from './api';

export async function getAdminFeatureFlags(): Promise<FeatureFlagsEffective> {
  const r = await apiFetch<{ data: FeatureFlagsEffective }>('/api/admin/feature-flags', {
    method: 'GET',
  });
  return r.data;
}

export async function updateAdminFeatureFlags(
  patch: Partial<FeatureFlagsEffective>,
): Promise<FeatureFlagsEffective> {
  const r = await apiFetch<{ data: FeatureFlagsEffective }>('/api/admin/feature-flags', {
    method: 'PUT',
    json: { flags: patch },
  });
  return r.data;
}


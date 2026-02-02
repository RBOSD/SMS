export type FeatureFlagsEffective = {
  module_issues: boolean;
  module_plans: boolean;
  module_ai_review: boolean;
};

export type MeResponse = {
  isLogin: boolean;
  id: number;
  username: string;
  role: 'MANAGER' | 'VIEWER';
  isAdmin: boolean;
  features: FeatureFlagsEffective;
};

export async function apiFetch<T>(
  url: string,
  options: RequestInit & { json?: unknown } = {},
): Promise<T> {
  const { json, headers, ...rest } = options;
  const res = await fetch(url, {
    ...rest,
    credentials: 'include',
    headers: {
      ...(json ? { 'Content-Type': 'application/json' } : {}),
      ...(headers || {}),
    },
    body: json ? JSON.stringify(json) : rest.body,
  });

  const data = (await res.json().catch(() => ({}))) as any;
  if (!res.ok) {
    const message =
      typeof data?.message === 'string'
        ? data.message
        : typeof data?.error === 'string'
          ? data.error
          : `HTTP ${res.status}`;
    const err = new Error(message);
    (err as any).status = res.status;
    (err as any).data = data;
    throw err;
  }
  return data as T;
}

export async function getMe(): Promise<MeResponse> {
  return await apiFetch<MeResponse>('/api/auth/me', { method: 'GET' });
}

export async function login(username: string, password: string) {
  return await apiFetch('/api/auth/login', {
    method: 'POST',
    json: { username, password },
  });
}

export async function logout() {
  return await apiFetch('/api/auth/logout', { method: 'POST' });
}


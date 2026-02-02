import { apiFetch } from './api';

export type Plan = {
  id: number;
  name: string;
  year: string;
  status: string | null;
  createdAt: string;
  updatedAt: string;
};

export async function listPlans() {
  return await apiFetch<{
    page: number;
    pageSize: number;
    total: number;
    data: Plan[];
  }>('/api/plans?page=1&pageSize=20', { method: 'GET' });
}


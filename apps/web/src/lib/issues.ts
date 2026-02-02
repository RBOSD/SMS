import { apiFetch } from './api';

export type IssueListItem = {
  id: number;
  number: string;
  year: string | null;
  unit: string | null;
  status: string | null;
  content: string | null;
  latestRound: null | {
    round: number;
    handling: string | null;
    review: string | null;
    replyDate: string | null;
    responseDate: string | null;
  };
};

export async function listIssues() {
  return await apiFetch<{
    page: number;
    pageSize: number;
    total: number;
    data: IssueListItem[];
  }>('/api/issues?page=1&pageSize=20', { method: 'GET' });
}


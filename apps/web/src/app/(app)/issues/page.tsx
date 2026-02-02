'use client';

import { useEffect, useState } from 'react';
import { listIssues, type IssueListItem } from '../../../lib/issues';
import { useMe } from '../../../lib/useMe';

export default function IssuesPage() {
  const { me } = useMe();
  const [items, setItems] = useState<IssueListItem[] | null>(null);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    (async () => {
      setError(null);
      try {
        const r = await listIssues();
        setItems(r.data);
      } catch (e: any) {
        setItems(null);
        setError(e?.message || '載入失敗');
      }
    })();
  }, []);
  if (!me) return null;

  return (
    <div style={{ display: 'grid', gap: 12 }}>
      <div style={{ fontSize: 18, fontWeight: 900, color: '#0f172a' }}>開立事項</div>
      <div style={{ color: '#64748b', fontSize: 13 }}>
        這裡將會是 v2 的開立事項查詢/編輯介面（下一步會接上 API 與資料模型）。
      </div>
      <div
        style={{
          padding: 12,
          borderRadius: 12,
          border: '1px solid #e2e8f0',
          background: '#fff',
        }}
      >
        <div style={{ fontWeight: 800, marginBottom: 6 }}>列表（示範）</div>
        {error ? (
          <div style={{ color: '#b91c1c', fontSize: 13, fontWeight: 700 }}>{error}</div>
        ) : items == null ? (
          <div style={{ color: '#64748b', fontSize: 13 }}>載入中…</div>
        ) : items.length === 0 ? (
          <div style={{ color: '#64748b', fontSize: 13 }}>目前沒有資料</div>
        ) : (
          <div style={{ display: 'grid', gap: 8 }}>
            {items.map((it) => (
              <div
                key={it.id}
                style={{
                  padding: '10px 12px',
                  borderRadius: 12,
                  border: '1px solid #e2e8f0',
                  background: '#f8fafc',
                }}
              >
                <div style={{ fontWeight: 900, color: '#0f172a' }}>{it.number}</div>
                <div style={{ color: '#64748b', fontSize: 13 }}>
                  {it.unit || '-'} · {it.year || '-'}
                </div>
              </div>
            ))}
          </div>
        )}
      </div>
      <div
        style={{
          padding: 12,
          borderRadius: 12,
          border: '1px solid #e2e8f0',
          background: '#fff',
        }}
      >
        <div style={{ fontWeight: 800, marginBottom: 6 }}>AI 審查</div>
        <div style={{ color: '#64748b', fontSize: 13 }}>
          {me.features.module_ai_review
            ? 'AI 審查模組已啟用（會在編輯頁提供 AI 建議）。'
            : 'AI 審查模組已關閉（非管理員將看不到入口）。'}
        </div>
      </div>
    </div>
  );
}


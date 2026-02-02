'use client';

import { useEffect, useState } from 'react';
import { listPlans, type Plan } from '../../../lib/plans';

export default function PlansPage() {
  const [items, setItems] = useState<Plan[] | null>(null);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    (async () => {
      setError(null);
      try {
        const r = await listPlans();
        setItems(r.data);
      } catch (e: any) {
        setItems(null);
        setError(e?.message || '載入失敗');
      }
    })();
  }, []);

  return (
    <div style={{ display: 'grid', gap: 12 }}>
      <div style={{ fontSize: 18, fontWeight: 900, color: '#0f172a' }}>檢查計畫</div>
      <div style={{ color: '#64748b', fontSize: 13 }}>
        這裡將會是 v2 的檢查計畫/行程檢索介面（下一步會接上 Plans/Schedule API）。
      </div>
      <div
        style={{
          padding: 12,
          borderRadius: 12,
          border: '1px solid #e2e8f0',
          background: '#fff',
          maxWidth: 640,
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
            {items.map((p) => (
              <div
                key={p.id}
                style={{
                  padding: '10px 12px',
                  borderRadius: 12,
                  border: '1px solid #e2e8f0',
                  background: '#f8fafc',
                }}
              >
                <div style={{ fontWeight: 900, color: '#0f172a' }}>
                  {p.year} · {p.name}
                </div>
                <div style={{ color: '#64748b', fontSize: 13 }}>{p.status || '-'}</div>
              </div>
            ))}
          </div>
        )}
      </div>
    </div>
  );
}


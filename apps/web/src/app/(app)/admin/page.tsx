'use client';

import { useEffect, useState } from 'react';
import type { FeatureFlagsEffective } from '../../../lib/api';
import { getAdminFeatureFlags, updateAdminFeatureFlags } from '../../../lib/admin';
import { useMe } from '../../../lib/useMe';

export default function AdminPage() {
  const { me } = useMe();
  const [flags, setFlags] = useState<FeatureFlagsEffective | null>(null);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [importMsg, setImportMsg] = useState<string | null>(null);

  useEffect(() => {
    (async () => {
      setError(null);
      try {
        const f = await getAdminFeatureFlags();
        setFlags(f);
      } catch (e: any) {
        setError(e?.message || '載入失敗');
      }
    })();
  }, []);

  async function onSave() {
    if (!flags) return;
    setSaving(true);
    setError(null);
    try {
      const updated = await updateAdminFeatureFlags(flags);
      setFlags(updated);
    } catch (e: any) {
      setError(e?.message || '儲存失敗');
    } finally {
      setSaving(false);
    }
  }

  if (!me) return null;

  async function upload(kind: 'issues' | 'plans', file: File) {
    setImportMsg(null);
    try {
      const fd = new FormData();
      fd.append('file', file);
      const res = await fetch(`/api/admin/import/${kind}`, {
        method: 'POST',
        credentials: 'include',
        body: fd,
      });
      const j = await res.json().catch(() => ({}));
      if (!res.ok) throw new Error(j?.message || j?.error || `HTTP ${res.status}`);
      setImportMsg(`${kind} 匯入完成：新增 ${j.created}、更新 ${j.updated}、跳過 ${j.skipped}`);
    } catch (e: any) {
      setImportMsg(e?.message || '匯入失敗');
    }
  }

  return (
    <div style={{ display: 'grid', gap: 12 }}>
      <div style={{ fontSize: 18, fontWeight: 900, color: '#0f172a' }}>後台管理</div>
      <div style={{ color: '#64748b', fontSize: 13 }}>
        模組開關只會影響非管理員：入口會隱藏，且後端 API 會回 404（像不存在）。
      </div>

      <div
        style={{
          padding: 14,
          borderRadius: 12,
          border: '1px solid #e2e8f0',
          background: '#fff',
          display: 'grid',
          gap: 10,
          maxWidth: 560,
        }}
      >
        <div style={{ fontWeight: 900 }}>模組開關</div>

        {!flags ? (
          <div style={{ color: '#64748b', fontSize: 13 }}>載入中…</div>
        ) : (
          <>
            <label style={{ display: 'flex', gap: 10, alignItems: 'center' }}>
              <input
                type="checkbox"
                checked={flags.module_issues}
                onChange={(e) =>
                  setFlags({ ...flags, module_issues: e.target.checked })
                }
              />
              <span style={{ fontWeight: 700 }}>開立事項</span>
            </label>

            <label style={{ display: 'flex', gap: 10, alignItems: 'center' }}>
              <input
                type="checkbox"
                checked={flags.module_plans}
                onChange={(e) =>
                  setFlags({ ...flags, module_plans: e.target.checked })
                }
              />
              <span style={{ fontWeight: 700 }}>檢查計畫</span>
            </label>

            <label style={{ display: 'flex', gap: 10, alignItems: 'center' }}>
              <input
                type="checkbox"
                checked={flags.module_ai_review}
                onChange={(e) =>
                  setFlags({ ...flags, module_ai_review: e.target.checked })
                }
              />
              <span style={{ fontWeight: 700 }}>AI 審查</span>
            </label>

            {error ? (
              <div style={{ color: '#b91c1c', fontSize: 13, fontWeight: 700 }}>
                {error}
              </div>
            ) : null}

            <button
              onClick={onSave}
              disabled={saving || !flags}
              style={{
                marginTop: 6,
                padding: '10px 12px',
                borderRadius: 12,
                border: '1px solid #2563eb',
                background: '#2563eb',
                color: '#fff',
                fontWeight: 900,
                cursor: saving ? 'not-allowed' : 'pointer',
              }}
            >
              {saving ? '儲存中…' : '儲存設定'}
            </button>
          </>
        )}
      </div>

      <div
        style={{
          padding: 14,
          borderRadius: 12,
          border: '1px solid #e2e8f0',
          background: '#fff',
          display: 'grid',
          gap: 10,
          maxWidth: 560,
        }}
      >
        <div style={{ fontWeight: 900 }}>匯入/匯出（CSV/XLSX）</div>
        <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap' }}>
          <a
            href="/api/admin/export/issues.csv"
            style={{
              padding: '8px 10px',
              borderRadius: 12,
              border: '1px solid #e2e8f0',
              background: '#fff',
              textDecoration: 'none',
              fontWeight: 800,
              color: '#0f172a',
            }}
          >
            下載開立事項 CSV
          </a>
          <a
            href="/api/admin/export/plans.csv"
            style={{
              padding: '8px 10px',
              borderRadius: 12,
              border: '1px solid #e2e8f0',
              background: '#fff',
              textDecoration: 'none',
              fontWeight: 800,
              color: '#0f172a',
            }}
          >
            下載檢查計畫 CSV
          </a>
        </div>

        <div style={{ display: 'grid', gap: 8 }}>
          <label style={{ display: 'grid', gap: 6 }}>
            <span style={{ fontWeight: 800, fontSize: 13 }}>匯入開立事項（.csv / .xlsx）</span>
            <input
              type="file"
              accept=".csv,.xlsx"
              onChange={(e) => {
                const f = e.target.files?.[0];
                if (f) upload('issues', f);
              }}
            />
          </label>

          <label style={{ display: 'grid', gap: 6 }}>
            <span style={{ fontWeight: 800, fontSize: 13 }}>匯入檢查計畫（.csv / .xlsx）</span>
            <input
              type="file"
              accept=".csv,.xlsx"
              onChange={(e) => {
                const f = e.target.files?.[0];
                if (f) upload('plans', f);
              }}
            />
          </label>
        </div>

        {importMsg ? (
          <div style={{ fontSize: 13, fontWeight: 700, color: '#0f172a' }}>{importMsg}</div>
        ) : (
          <div style={{ fontSize: 12, color: '#64748b' }}>
            匯入規則：至少需要欄位 number（事項）或 name+year（計畫）。可使用中文欄名（例如：編號、計畫名稱、年度）。
          </div>
        )}
      </div>
    </div>
  );
}


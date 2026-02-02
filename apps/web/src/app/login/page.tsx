'use client';

import { useState } from 'react';
import { useRouter } from 'next/navigation';
import { login } from '../../lib/api';

export default function LoginPage() {
  const router = useRouter();
  const [username, setUsername] = useState('');
  const [password, setPassword] = useState('');
  const [error, setError] = useState<string | null>(null);
  const [loading, setLoading] = useState(false);

  async function onSubmit(e: React.FormEvent) {
    e.preventDefault();
    setError(null);
    setLoading(true);
    try {
      await login(username, password);
      router.replace('/');
    } catch (e: any) {
      setError(e?.message || '登入失敗');
    } finally {
      setLoading(false);
    }
  }

  return (
    <div
      style={{
        minHeight: '100vh',
        display: 'grid',
        placeItems: 'center',
        background: '#0b1220',
        padding: 24,
      }}
    >
      <form
        onSubmit={onSubmit}
        style={{
          width: '100%',
          maxWidth: 420,
          background: '#ffffff',
          borderRadius: 16,
          padding: 20,
          border: '1px solid #e2e8f0',
        }}
      >
        <div style={{ fontWeight: 900, fontSize: 20, color: '#0f172a' }}>
          SMS v2 登入
        </div>
        <div style={{ marginTop: 6, color: '#64748b', fontSize: 13 }}>
          請使用後台建立的帳號登入。
        </div>

        <div style={{ marginTop: 16 }}>
          <label style={{ display: 'block', fontSize: 13, fontWeight: 700, color: '#334155' }}>
            帳號
          </label>
          <input
            value={username}
            onChange={(e) => setUsername(e.target.value)}
            autoComplete="username"
            style={{
              marginTop: 6,
              width: '100%',
              padding: '10px 12px',
              borderRadius: 12,
              border: '1px solid #cbd5e1',
              outline: 'none',
            }}
          />
        </div>

        <div style={{ marginTop: 12 }}>
          <label style={{ display: 'block', fontSize: 13, fontWeight: 700, color: '#334155' }}>
            密碼
          </label>
          <input
            value={password}
            onChange={(e) => setPassword(e.target.value)}
            type="password"
            autoComplete="current-password"
            style={{
              marginTop: 6,
              width: '100%',
              padding: '10px 12px',
              borderRadius: 12,
              border: '1px solid #cbd5e1',
              outline: 'none',
            }}
          />
        </div>

        {error ? (
          <div style={{ marginTop: 12, color: '#b91c1c', fontSize: 13, fontWeight: 600 }}>
            {error}
          </div>
        ) : null}

        <button
          disabled={loading}
          type="submit"
          style={{
            marginTop: 16,
            width: '100%',
            padding: '10px 12px',
            borderRadius: 12,
            border: '1px solid #2563eb',
            background: '#2563eb',
            color: '#fff',
            fontWeight: 800,
            cursor: loading ? 'not-allowed' : 'pointer',
          }}
        >
          {loading ? '登入中…' : '登入'}
        </button>
      </form>
    </div>
  );
}


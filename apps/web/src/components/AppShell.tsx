'use client';

import Link from 'next/link';
import { usePathname, useRouter } from 'next/navigation';
import type { ReactNode } from 'react';
import { logout, type MeResponse } from '../lib/api';

export function AppShell(props: { me: MeResponse; children: ReactNode }) {
  const pathname = usePathname();
  const router = useRouter();

  const links: Array<{ href: string; label: string; show: boolean }> = [
    { href: '/issues', label: '開立事項', show: props.me.features.module_issues },
    { href: '/plans', label: '檢查計畫', show: props.me.features.module_plans },
    { href: '/admin', label: '後台管理', show: props.me.isAdmin },
  ];

  async function onLogout() {
    try {
      await logout();
    } finally {
      router.replace('/login');
    }
  }

  return (
    <div
      style={{
        minHeight: '100vh',
        display: 'grid',
        gridTemplateColumns: '240px 1fr',
        background: '#f8fafc',
      }}
    >
      <aside
        style={{
          borderRight: '1px solid #e2e8f0',
          background: '#fff',
          padding: 16,
        }}
      >
        <div style={{ fontWeight: 800, marginBottom: 12 }}>SMS v2</div>
        <nav style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
          {links
            .filter((l) => l.show)
            .map((l) => {
              const active = pathname === l.href;
              return (
                <Link
                  key={l.href}
                  href={l.href}
                  style={{
                    padding: '10px 12px',
                    borderRadius: 10,
                    textDecoration: 'none',
                    color: active ? '#0f172a' : '#334155',
                    background: active ? '#eef2ff' : 'transparent',
                    border: active ? '1px solid #c7d2fe' : '1px solid transparent',
                    fontWeight: active ? 700 : 600,
                  }}
                >
                  {l.label}
                </Link>
              );
            })}
        </nav>
      </aside>

      <div style={{ display: 'flex', flexDirection: 'column' }}>
        <header
          style={{
            height: 56,
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'space-between',
            padding: '0 16px',
            borderBottom: '1px solid #e2e8f0',
            background: '#fff',
          }}
        >
          <div style={{ color: '#475569', fontWeight: 600 }}>
            {props.me.isAdmin ? '系統管理員' : props.me.role === 'MANAGER' ? '管理者' : '檢視者'} ·{' '}
            {props.me.username}
          </div>
          <button
            onClick={onLogout}
            style={{
              padding: '8px 10px',
              borderRadius: 10,
              border: '1px solid #e2e8f0',
              background: '#fff',
              cursor: 'pointer',
              fontWeight: 600,
            }}
          >
            登出
          </button>
        </header>

        <main style={{ padding: 16 }}>{props.children}</main>
      </div>
    </div>
  );
}


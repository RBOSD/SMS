'use client';

import { useEffect } from 'react';
import { usePathname, useRouter } from 'next/navigation';
import { AppShell } from '../../components/AppShell';
import { useMe } from '../../lib/useMe';

export default function AppLayout({ children }: { children: React.ReactNode }) {
  const { me, loading } = useMe();
  const router = useRouter();
  const pathname = usePathname();

  useEffect(() => {
    if (!loading && !me) router.replace('/login');
  }, [loading, me, router]);

  useEffect(() => {
    if (!me) return;
    if (pathname.startsWith('/issues') && !me.features.module_issues) {
      router.replace('/');
    }
    if (pathname.startsWith('/plans') && !me.features.module_plans) {
      router.replace('/');
    }
    if (pathname.startsWith('/admin') && !me.isAdmin) {
      router.replace('/');
    }
  }, [me, pathname, router]);

  if (loading) {
    return (
      <div style={{ padding: 24, fontFamily: 'system-ui', color: '#475569' }}>
        載入中…
      </div>
    );
  }
  if (!me) return null;

  return <AppShell me={me}>{children}</AppShell>;
}


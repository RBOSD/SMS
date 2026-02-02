'use client';

import { useEffect, useState } from 'react';
import { useRouter } from 'next/navigation';
import { getMe } from '../lib/api';

export default function Home() {
  const router = useRouter();
  const [message, setMessage] = useState('載入中…');

  useEffect(() => {
    (async () => {
      try {
        const me = await getMe();
        if (!me?.isLogin) {
          router.replace('/login');
          return;
        }
        if (me.features.module_issues) {
          router.replace('/issues');
          return;
        }
        if (me.features.module_plans) {
          router.replace('/plans');
          return;
        }
        if (me.isAdmin) {
          router.replace('/admin');
          return;
        }
        setMessage('目前無可用模組，請聯絡系統管理員。');
      } catch {
        router.replace('/login');
      }
    })();
  }, [router]);

  return (
    <div style={{ padding: 24, fontFamily: 'system-ui', color: '#475569' }}>
      {message}
    </div>
  );
}

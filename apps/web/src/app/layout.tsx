import type { Metadata } from 'next';
import './globals.css';

export const metadata: Metadata = {
  title: "SMS v2",
  description: "SMS 開立事項系統（v2）",
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="zh-TW">
      <body>{children}</body>
    </html>
  );
}

import type { Metadata } from 'next';
import './globals.css';

export const metadata: Metadata = {
  title: 'Fair Price PPT Builder',
  description: 'Generate Visual Merch presentations from Perigee survey data',
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="en">
      <body>{children}</body>
    </html>
  );
}

import React from 'react';
import { TopBar } from './TopBar';
import { cn } from '@/utils/cn';

export interface PageLayoutProps {
  title: string;
  subtitle?: string;
  children: React.ReactNode;
  className?: string;
  noPadding?: boolean;
}

export function PageLayout({
  title,
  subtitle,
  children,
  className,
  noPadding = false,
}: PageLayoutProps) {
  return (
    <div className="flex flex-col h-full w-full overflow-hidden">
      <TopBar title={title} subtitle={subtitle} />
      <main
        className={cn(
          'flex-1 overflow-y-auto bg-somus-gray-50',
          !noPadding && 'p-6',
          className
        )}
      >
        {children}
      </main>
    </div>
  );
}

export default PageLayout;

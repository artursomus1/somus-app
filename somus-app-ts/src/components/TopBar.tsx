import React from 'react';
import { Minus, Square, X, User, Bell } from 'lucide-react';
import { useAppStore, MODULE_LABELS } from '@/stores/appStore';
import { cn } from '@/utils/cn';

export interface TopBarProps {
  title: string;
  subtitle?: string;
  className?: string;
}

export function TopBar({ title, subtitle, className }: TopBarProps) {
  const { userName, currentModule } = useAppStore();

  const displaySubtitle = subtitle ?? MODULE_LABELS[currentModule];

  const handleMinimize = () => {
    window.electron?.app.minimize();
  };
  const handleMaximize = () => {
    window.electron?.app.maximize();
  };
  const handleClose = () => {
    window.electron?.app.close();
  };

  return (
    <header
      className={cn(
        'flex items-center justify-between h-12 px-5 bg-somus-bg-secondary border-b border-somus-border select-none',
        className
      )}
      style={{ WebkitAppRegion: 'drag' } as React.CSSProperties}
    >
      {/* Left: Breadcrumb - Module > Page */}
      <div className="flex items-baseline gap-2 min-w-0">
        <h1 className="text-base font-semibold text-somus-text-primary truncate">
          {title}
        </h1>
        {displaySubtitle && (
          <>
            <span className="text-somus-text-tertiary">/</span>
            <span className="text-sm text-somus-text-secondary truncate">
              {displaySubtitle}
            </span>
          </>
        )}
      </div>

      {/* Right: Notification, User, Version, Window Controls */}
      <div
        className="flex items-center gap-4 shrink-0"
        style={{ WebkitAppRegion: 'no-drag' } as React.CSSProperties}
      >
        {/* Notification placeholder */}
        <button className="relative p-1.5 rounded-lg text-somus-text-tertiary hover:text-somus-text-secondary hover:bg-somus-bg-hover transition-colors">
          <Bell className="h-4 w-4" />
        </button>

        {/* User */}
        {userName && (
          <div className="flex items-center gap-2 text-sm text-somus-text-secondary">
            <div className="h-6 w-6 rounded-full bg-somus-bg-tertiary border border-somus-border flex items-center justify-center">
              <User className="h-3.5 w-3.5 text-somus-text-accent" />
            </div>
            <span className="hidden sm:inline">{userName}</span>
          </div>
        )}

        {/* Version */}
        <span className="text-xs text-somus-text-tertiary font-mono">v2.0.0</span>

        {/* Window Controls */}
        <div className="flex items-center -mr-2">
          <button
            onClick={handleMinimize}
            className="p-2 text-somus-text-tertiary hover:text-somus-text-secondary hover:bg-somus-bg-hover transition-colors rounded"
            title="Minimizar"
          >
            <Minus className="h-4 w-4" />
          </button>
          <button
            onClick={handleMaximize}
            className="p-2 text-somus-text-tertiary hover:text-somus-text-secondary hover:bg-somus-bg-hover transition-colors rounded"
            title="Maximizar"
          >
            <Square className="h-3.5 w-3.5" />
          </button>
          <button
            onClick={handleClose}
            className="p-2 text-somus-text-tertiary hover:text-white hover:bg-red-600/80 transition-colors rounded"
            title="Fechar"
          >
            <X className="h-4 w-4" />
          </button>
        </div>
      </div>
    </header>
  );
}

export default TopBar;

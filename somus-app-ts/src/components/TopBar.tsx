import React from 'react';
import { Minus, Square, X, User } from 'lucide-react';
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
        'flex items-center justify-between h-12 px-5 bg-white border-b border-somus-gray-200 select-none',
        className
      )}
      style={{ WebkitAppRegion: 'drag' } as React.CSSProperties}
    >
      {/* Left: Title */}
      <div className="flex items-baseline gap-2 min-w-0">
        <h1 className="text-base font-semibold text-somus-gray-900 truncate">
          {title}
        </h1>
        {displaySubtitle && (
          <>
            <span className="text-somus-gray-300">/</span>
            <span className="text-sm text-somus-gray-500 truncate">
              {displaySubtitle}
            </span>
          </>
        )}
      </div>

      {/* Right: User, Version, Window Controls */}
      <div
        className="flex items-center gap-4 shrink-0"
        style={{ WebkitAppRegion: 'no-drag' } as React.CSSProperties}
      >
        {/* User */}
        {userName && (
          <div className="flex items-center gap-2 text-sm text-somus-gray-600">
            <User className="h-4 w-4" />
            <span className="hidden sm:inline">{userName}</span>
          </div>
        )}

        {/* Version */}
        <span className="text-xs text-somus-gray-400 font-mono">v2.0.0</span>

        {/* Window Controls */}
        <div className="flex items-center -mr-2">
          <button
            onClick={handleMinimize}
            className="p-2 text-somus-gray-400 hover:text-somus-gray-600 hover:bg-somus-gray-100 transition-colors rounded"
            title="Minimizar"
          >
            <Minus className="h-4 w-4" />
          </button>
          <button
            onClick={handleMaximize}
            className="p-2 text-somus-gray-400 hover:text-somus-gray-600 hover:bg-somus-gray-100 transition-colors rounded"
            title="Maximizar"
          >
            <Square className="h-3.5 w-3.5" />
          </button>
          <button
            onClick={handleClose}
            className="p-2 text-somus-gray-400 hover:text-white hover:bg-red-500 transition-colors rounded"
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

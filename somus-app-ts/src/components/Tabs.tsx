import React, { useState } from 'react';
import { cn } from '@/utils/cn';

export interface Tab {
  id: string;
  label: string;
  icon?: React.ReactNode;
  disabled?: boolean;
}

export interface TabsProps {
  tabs: Tab[];
  defaultTab?: string;
  activeTab?: string;
  onChange?: (tabId: string) => void;
  children?: (activeTabId: string) => React.ReactNode;
  className?: string;
}

export function Tabs({
  tabs,
  defaultTab,
  activeTab: controlledActive,
  onChange,
  children,
  className,
}: TabsProps) {
  const [internalActive, setInternalActive] = useState(
    defaultTab ?? tabs[0]?.id ?? ''
  );

  const activeTab = controlledActive ?? internalActive;

  const handleSelect = (tabId: string) => {
    setInternalActive(tabId);
    onChange?.(tabId);
  };

  return (
    <div className={cn('w-full', className)}>
      {/* Tab Headers */}
      <div className="flex border-b border-somus-border">
        {tabs.map((tab) => (
          <button
            key={tab.id}
            onClick={() => !tab.disabled && handleSelect(tab.id)}
            disabled={tab.disabled}
            className={cn(
              'flex items-center gap-2 px-4 py-2.5 text-sm font-medium transition-colors duration-150 border-b-2 -mb-px',
              'focus:outline-none',
              activeTab === tab.id
                ? 'border-somus-green-500 text-somus-text-accent'
                : 'border-transparent text-somus-text-tertiary hover:text-somus-text-secondary hover:border-somus-border-light',
              tab.disabled && 'opacity-40 cursor-not-allowed'
            )}
          >
            {tab.icon && <span className="shrink-0">{tab.icon}</span>}
            {tab.label}
          </button>
        ))}
      </div>

      {/* Tab Content */}
      {children && <div className="pt-4">{children(activeTab)}</div>}
    </div>
  );
}

export default Tabs;

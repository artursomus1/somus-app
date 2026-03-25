import React from 'react';
import { cn } from '@/utils/cn';

export interface CardProps {
  title?: string;
  subtitle?: string;
  headerRight?: React.ReactNode;
  padding?: 'none' | 'sm' | 'md' | 'lg';
  accent?: 'none' | 'top' | 'left' | 'green' | 'purple' | 'navy' | 'gold';
  className?: string;
  children: React.ReactNode;
}

const paddingStyles = {
  none: '',
  sm: 'p-3',
  md: 'p-5',
  lg: 'p-6',
};

const accentStyles: Record<NonNullable<CardProps['accent']>, string> = {
  none: '',
  top: 'border-t-2 border-t-somus-green-500',
  left: 'border-l-3 border-l-somus-green-500',
  green: 'border-l-3 border-l-somus-green-500',
  purple: 'border-l-3 border-l-[#7030A0]',
  navy: 'border-l-3 border-l-[#002060]',
  gold: 'border-l-3 border-l-[#D4A017]',
};

export function Card({
  title,
  subtitle,
  headerRight,
  padding = 'md',
  accent = 'none',
  className,
  children,
}: CardProps) {
  return (
    <div
      className={cn(
        'bg-somus-bg-secondary/80 backdrop-blur-xl rounded-xl border border-somus-border/50',
        'transition-all duration-200 hover:border-somus-border-light/60',
        accentStyles[accent],
        className
      )}
    >
      {(title || headerRight) && (
        <div className="flex items-center justify-between px-5 pt-5 pb-0">
          <div>
            {title && (
              <h3 className="text-base font-semibold text-somus-text-primary">
                {title}
              </h3>
            )}
            {subtitle && (
              <p className="text-sm text-somus-text-secondary mt-0.5">{subtitle}</p>
            )}
          </div>
          {headerRight && <div>{headerRight}</div>}
        </div>
      )}
      <div className={cn(paddingStyles[padding])}>{children}</div>
    </div>
  );
}

export default Card;

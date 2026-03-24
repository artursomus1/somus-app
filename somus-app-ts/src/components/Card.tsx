import React from 'react';
import { cn } from '@/utils/cn';

export interface CardProps {
  title?: string;
  subtitle?: string;
  headerRight?: React.ReactNode;
  padding?: 'none' | 'sm' | 'md' | 'lg';
  className?: string;
  children: React.ReactNode;
}

const paddingStyles = {
  none: '',
  sm: 'p-3',
  md: 'p-5',
  lg: 'p-6',
};

export function Card({
  title,
  subtitle,
  headerRight,
  padding = 'md',
  className,
  children,
}: CardProps) {
  return (
    <div
      className={cn(
        'bg-white rounded-lg shadow-sm border border-somus-gray-200',
        className
      )}
    >
      {(title || headerRight) && (
        <div className="flex items-center justify-between px-5 pt-5 pb-0">
          <div>
            {title && (
              <h3 className="text-base font-semibold text-somus-gray-900">
                {title}
              </h3>
            )}
            {subtitle && (
              <p className="text-sm text-somus-gray-500 mt-0.5">{subtitle}</p>
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

import React from 'react';
import { TrendingUp, TrendingDown } from 'lucide-react';
import { cn } from '@/utils/cn';

export interface KPICardProps {
  title: string;
  value: string;
  subtitle?: string;
  icon?: React.ReactNode;
  trend?: { value: number; isPositive: boolean };
  variant?: 'default' | 'green' | 'blue' | 'orange' | 'red';
  onClick?: () => void;
  className?: string;
}

const variantStyles = {
  default: {
    bg: 'bg-white',
    iconBg: 'bg-somus-gray-100',
    iconText: 'text-somus-gray-600',
  },
  green: {
    bg: 'bg-white',
    iconBg: 'bg-emerald-50',
    iconText: 'text-emerald-600',
  },
  blue: {
    bg: 'bg-white',
    iconBg: 'bg-blue-50',
    iconText: 'text-blue-600',
  },
  orange: {
    bg: 'bg-white',
    iconBg: 'bg-orange-50',
    iconText: 'text-orange-600',
  },
  red: {
    bg: 'bg-white',
    iconBg: 'bg-red-50',
    iconText: 'text-red-600',
  },
};

export function KPICard({
  title,
  value,
  subtitle,
  icon,
  trend,
  variant = 'default',
  onClick,
  className,
}: KPICardProps) {
  const styles = variantStyles[variant];

  return (
    <div
      onClick={onClick}
      className={cn(
        'rounded-lg border border-somus-gray-200 shadow-sm p-5 transition-all duration-150',
        styles.bg,
        onClick && 'cursor-pointer hover:shadow-md hover:border-somus-gray-300',
        className
      )}
    >
      <div className="flex items-start justify-between">
        <div className="flex-1 min-w-0">
          <p className="text-sm font-medium text-somus-gray-500 truncate">
            {title}
          </p>
          <p className="mt-2 text-2xl font-bold text-somus-gray-900 tracking-tight">
            {value}
          </p>
          <div className="mt-2 flex items-center gap-2">
            {trend && (
              <span
                className={cn(
                  'inline-flex items-center gap-0.5 text-xs font-semibold',
                  trend.isPositive ? 'text-emerald-600' : 'text-red-600'
                )}
              >
                {trend.isPositive ? (
                  <TrendingUp className="h-3.5 w-3.5" />
                ) : (
                  <TrendingDown className="h-3.5 w-3.5" />
                )}
                {trend.value.toLocaleString('pt-BR', { minimumFractionDigits: 1 })}%
              </span>
            )}
            {subtitle && (
              <span className="text-xs text-somus-gray-400">{subtitle}</span>
            )}
          </div>
        </div>

        {icon && (
          <div
            className={cn(
              'shrink-0 p-2.5 rounded-lg',
              styles.iconBg,
              styles.iconText
            )}
          >
            {icon}
          </div>
        )}
      </div>
    </div>
  );
}

export default KPICard;

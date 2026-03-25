import React from 'react';
import { TrendingUp, TrendingDown } from 'lucide-react';
import { cn } from '@/utils/cn';

export interface KPICardProps {
  title: string;
  value: string;
  subtitle?: string;
  icon?: React.ReactNode;
  trend?: { value: number; isPositive: boolean };
  variant?: 'default' | 'green' | 'purple' | 'navy' | 'gold' | 'orange' | 'red';
  onClick?: () => void;
  className?: string;
}

const variantConfig: Record<NonNullable<KPICardProps['variant']>, { borderColor: string; iconBg: string; iconText: string; glowColor: string }> = {
  default: {
    borderColor: '#1E2A3A',
    iconBg: 'bg-somus-bg-tertiary',
    iconText: 'text-somus-text-secondary',
    glowColor: 'rgba(26, 122, 62, 0.08)',
  },
  green: {
    borderColor: '#1A7A3E',
    iconBg: 'bg-somus-green-700/20',
    iconText: 'text-somus-text-accent',
    glowColor: 'rgba(26, 122, 62, 0.12)',
  },
  purple: {
    borderColor: '#7030A0',
    iconBg: 'bg-purple-900/20',
    iconText: 'text-purple-400',
    glowColor: 'rgba(112, 48, 160, 0.12)',
  },
  navy: {
    borderColor: '#002060',
    iconBg: 'bg-blue-900/20',
    iconText: 'text-blue-400',
    glowColor: 'rgba(0, 32, 96, 0.12)',
  },
  gold: {
    borderColor: '#D4A017',
    iconBg: 'bg-yellow-900/20',
    iconText: 'text-yellow-400',
    glowColor: 'rgba(212, 160, 23, 0.12)',
  },
  orange: {
    borderColor: '#ED7D31',
    iconBg: 'bg-orange-900/20',
    iconText: 'text-orange-400',
    glowColor: 'rgba(237, 125, 49, 0.12)',
  },
  red: {
    borderColor: '#C00000',
    iconBg: 'bg-red-900/20',
    iconText: 'text-red-400',
    glowColor: 'rgba(192, 0, 0, 0.12)',
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
  const config = variantConfig[variant];

  return (
    <div
      onClick={onClick}
      className={cn(
        'relative overflow-hidden rounded-xl p-5 transition-all duration-200',
        'bg-somus-bg-secondary/80 backdrop-blur-xl border border-somus-border/50',
        onClick && 'cursor-pointer hover:border-somus-border-light',
        className
      )}
      style={{
        borderLeftWidth: '3px',
        borderLeftColor: config.borderColor,
      }}
      onMouseEnter={(e) => {
        if (onClick) {
          (e.currentTarget as HTMLElement).style.boxShadow = `0 0 20px ${config.glowColor}`;
        }
      }}
      onMouseLeave={(e) => {
        (e.currentTarget as HTMLElement).style.boxShadow = 'none';
      }}
    >
      {/* Subtle top gradient line */}
      <div
        className="absolute top-0 left-0 right-0 h-px"
        style={{ background: `linear-gradient(90deg, transparent, ${config.borderColor}40, transparent)` }}
      />

      <div className="flex items-start justify-between">
        <div className="flex-1 min-w-0">
          <p className="text-xs font-medium text-somus-text-tertiary uppercase tracking-wider truncate">
            {title}
          </p>
          <p className="mt-2 text-2xl font-bold text-somus-text-primary tracking-tight">
            {value}
          </p>
          <div className="mt-2 flex items-center gap-2">
            {trend && (
              <span
                className={cn(
                  'inline-flex items-center gap-0.5 text-xs font-semibold',
                  trend.isPositive ? 'text-somus-text-accent' : 'text-somus-text-danger'
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
              <span className="text-xs text-somus-text-tertiary">{subtitle}</span>
            )}
          </div>
        </div>

        {icon && (
          <div
            className={cn(
              'shrink-0 p-2.5 rounded-lg',
              config.iconBg,
              config.iconText
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

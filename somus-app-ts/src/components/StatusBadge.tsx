import React from 'react';
import { cn } from '@/utils/cn';

export type StatusVariant = 'ativa' | 'pendente' | 'encerrada' | 'atrasado';

export interface StatusBadgeProps {
  status: StatusVariant;
  label?: string;
  className?: string;
}

const statusConfig: Record<StatusVariant, { bg: string; text: string; dot: string; defaultLabel: string }> = {
  ativa: {
    bg: 'bg-emerald-900/30',
    text: 'text-emerald-400',
    dot: 'bg-emerald-400',
    defaultLabel: 'Ativa',
  },
  pendente: {
    bg: 'bg-yellow-900/30',
    text: 'text-yellow-400',
    dot: 'bg-yellow-400',
    defaultLabel: 'Pendente',
  },
  encerrada: {
    bg: 'bg-somus-bg-tertiary',
    text: 'text-somus-text-tertiary',
    dot: 'bg-somus-text-tertiary',
    defaultLabel: 'Encerrada',
  },
  atrasado: {
    bg: 'bg-red-900/30',
    text: 'text-red-400',
    dot: 'bg-red-400',
    defaultLabel: 'Atrasado',
  },
};

export function StatusBadge({ status, label, className }: StatusBadgeProps) {
  const config = statusConfig[status];
  return (
    <span
      className={cn(
        'inline-flex items-center gap-1.5 px-2.5 py-0.5 rounded-full text-xs font-medium',
        config.bg,
        config.text,
        className
      )}
    >
      <span className={cn('h-1.5 w-1.5 rounded-full', config.dot)} />
      {label ?? config.defaultLabel}
    </span>
  );
}

export default StatusBadge;

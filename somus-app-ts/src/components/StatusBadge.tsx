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
    bg: 'bg-emerald-50',
    text: 'text-emerald-700',
    dot: 'bg-emerald-500',
    defaultLabel: 'Ativa',
  },
  pendente: {
    bg: 'bg-yellow-50',
    text: 'text-yellow-700',
    dot: 'bg-yellow-500',
    defaultLabel: 'Pendente',
  },
  encerrada: {
    bg: 'bg-somus-gray-100',
    text: 'text-somus-gray-600',
    dot: 'bg-somus-gray-400',
    defaultLabel: 'Encerrada',
  },
  atrasado: {
    bg: 'bg-red-50',
    text: 'text-red-700',
    dot: 'bg-red-500',
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

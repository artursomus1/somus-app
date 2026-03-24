import { clsx, type ClassValue } from 'clsx';
import { twMerge } from 'tailwind-merge';

export function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs));
}

/**
 * Formata numero para padrão brasileiro: 1.234.567,89
 */
export function formatBRL(value: number): string {
  return value.toLocaleString('pt-BR', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
}

/**
 * Formata numero como moeda BRL: R$ 1.234,56
 */
export function formatCurrency(value: number): string {
  return value.toLocaleString('pt-BR', {
    style: 'currency',
    currency: 'BRL',
  });
}

/**
 * Formata numero como percentual: 12,34%
 */
export function formatPercent(value: number, decimals = 2): string {
  return `${value.toLocaleString('pt-BR', {
    minimumFractionDigits: decimals,
    maximumFractionDigits: decimals,
  })}%`;
}

/**
 * Remove formatacao e retorna numero
 */
export function parseBRLNumber(value: string): number {
  const cleaned = value.replace(/[^\d,-]/g, '').replace(',', '.');
  return parseFloat(cleaned) || 0;
}

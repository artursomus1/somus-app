import React, { useState, useCallback } from 'react';
import { cn } from '@/utils/cn';

export interface CurrencyInputProps {
  value?: number;
  onChange?: (value: number) => void;
  label?: string;
  error?: string;
  placeholder?: string;
  disabled?: boolean;
  name?: string;
  className?: string;
}

/**
 * Converte centavos inteiros para string formatada "1.234,56"
 */
function formatCents(cents: number): string {
  const abs = Math.abs(cents);
  const intPart = Math.floor(abs / 100);
  const decPart = abs % 100;
  const intFormatted = intPart.toLocaleString('pt-BR');
  const decStr = decPart.toString().padStart(2, '0');
  const sign = cents < 0 ? '-' : '';
  return `${sign}${intFormatted},${decStr}`;
}

/**
 * Remove tudo exceto digitos e retorna centavos
 */
function parseToCents(raw: string): number {
  const digits = raw.replace(/\D/g, '');
  return parseInt(digits || '0', 10);
}

export const CurrencyInput = React.forwardRef<HTMLInputElement, CurrencyInputProps>(
  ({ value, onChange, label, error, placeholder = '0,00', disabled, name, className }, ref) => {
    // Controle interno: armazenamos centavos
    const [internalCents, setInternalCents] = useState<number>(
      value !== undefined ? Math.round(value * 100) : 0
    );

    const cents = value !== undefined ? Math.round(value * 100) : internalCents;
    const displayValue = cents === 0 ? '' : formatCents(cents);

    const handleChange = useCallback(
      (e: React.ChangeEvent<HTMLInputElement>) => {
        const newCents = parseToCents(e.target.value);
        setInternalCents(newCents);
        onChange?.(newCents / 100);
      },
      [onChange]
    );

    return (
      <div className={cn('flex flex-col gap-1.5', className)}>
        {label && (
          <label className="text-sm font-medium text-somus-gray-700">
            {label}
          </label>
        )}
        <div className="relative">
          <span className="absolute left-3 top-1/2 -translate-y-1/2 text-sm font-medium text-somus-gray-500 select-none">
            R$
          </span>
          <input
            ref={ref}
            type="text"
            inputMode="numeric"
            name={name}
            value={displayValue}
            onChange={handleChange}
            placeholder={placeholder}
            disabled={disabled}
            className={cn(
              'w-full rounded-lg border bg-white py-2 pl-10 pr-3 text-sm text-right font-medium transition-colors duration-150',
              'focus:outline-none focus:ring-2 focus:ring-somus-green/40 focus:border-somus-green',
              'disabled:opacity-50 disabled:cursor-not-allowed',
              error
                ? 'border-red-300 text-red-900'
                : 'border-somus-gray-300 text-somus-gray-900'
            )}
          />
        </div>
        {error && <p className="text-xs text-red-600">{error}</p>}
      </div>
    );
  }
);

CurrencyInput.displayName = 'CurrencyInput';

export default CurrencyInput;

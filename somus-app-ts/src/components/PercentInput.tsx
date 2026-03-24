import React, { useState, useCallback } from 'react';
import { cn } from '@/utils/cn';

export interface PercentInputProps {
  value?: number;
  onChange?: (value: number) => void;
  decimals?: number;
  label?: string;
  error?: string;
  placeholder?: string;
  disabled?: boolean;
  name?: string;
  className?: string;
}

function formatPercentValue(val: number, decimals: number): string {
  if (val === 0) return '';
  return val.toLocaleString('pt-BR', {
    minimumFractionDigits: decimals,
    maximumFractionDigits: decimals,
  });
}

function parsePercentInput(raw: string, decimals: number): number {
  const digits = raw.replace(/\D/g, '');
  if (!digits) return 0;
  const divisor = Math.pow(10, decimals);
  return parseInt(digits, 10) / divisor;
}

export const PercentInput = React.forwardRef<HTMLInputElement, PercentInputProps>(
  (
    {
      value,
      onChange,
      decimals = 2,
      label,
      error,
      placeholder = '0,00',
      disabled,
      name,
      className,
    },
    ref
  ) => {
    const [internalValue, setInternalValue] = useState<number>(value ?? 0);

    const current = value ?? internalValue;
    const displayValue = formatPercentValue(current, decimals);

    const handleChange = useCallback(
      (e: React.ChangeEvent<HTMLInputElement>) => {
        const parsed = parsePercentInput(e.target.value, decimals);
        setInternalValue(parsed);
        onChange?.(parsed);
      },
      [onChange, decimals]
    );

    return (
      <div className={cn('flex flex-col gap-1.5', className)}>
        {label && (
          <label className="text-sm font-medium text-somus-gray-700">
            {label}
          </label>
        )}
        <div className="relative">
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
              'w-full rounded-lg border bg-white py-2 pl-3 pr-8 text-sm text-right font-medium transition-colors duration-150',
              'focus:outline-none focus:ring-2 focus:ring-somus-green/40 focus:border-somus-green',
              'disabled:opacity-50 disabled:cursor-not-allowed',
              error
                ? 'border-red-300 text-red-900'
                : 'border-somus-gray-300 text-somus-gray-900'
            )}
          />
          <span className="absolute right-3 top-1/2 -translate-y-1/2 text-sm font-medium text-somus-gray-500 select-none">
            %
          </span>
        </div>
        {error && <p className="text-xs text-red-600">{error}</p>}
      </div>
    );
  }
);

PercentInput.displayName = 'PercentInput';

export default PercentInput;

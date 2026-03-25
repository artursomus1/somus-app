import React from 'react';
import { ChevronDown } from 'lucide-react';
import { cn } from '@/utils/cn';

export interface SelectOption {
  value: string;
  label: string;
  disabled?: boolean;
}

export interface SelectProps extends Omit<React.SelectHTMLAttributes<HTMLSelectElement>, 'size'> {
  options: SelectOption[];
  placeholder?: string;
  size?: 'sm' | 'md' | 'lg';
  error?: boolean;
}

const sizeStyles = {
  sm: 'text-xs py-1.5 pl-3 pr-8',
  md: 'text-sm py-2 pl-3 pr-9',
  lg: 'text-base py-2.5 pl-4 pr-10',
};

export const Select = React.forwardRef<HTMLSelectElement, SelectProps>(
  ({ options, placeholder, size = 'md', error, className, ...props }, ref) => {
    return (
      <div className="relative">
        <select
          ref={ref}
          className={cn(
            'w-full appearance-none rounded-lg border bg-somus-bg-input font-medium transition-colors duration-150',
            'focus:outline-none focus:ring-2 focus:ring-somus-green-500/30 focus:border-somus-green-500/50',
            'disabled:opacity-50 disabled:cursor-not-allowed',
            error
              ? 'border-red-500/50 text-red-400'
              : 'border-somus-border text-somus-text-primary',
            sizeStyles[size],
            className
          )}
          {...props}
        >
          {placeholder && (
            <option value="" disabled className="bg-somus-bg-secondary text-somus-text-tertiary">
              {placeholder}
            </option>
          )}
          {options.map((opt) => (
            <option key={opt.value} value={opt.value} disabled={opt.disabled} className="bg-somus-bg-secondary text-somus-text-primary">
              {opt.label}
            </option>
          ))}
        </select>
        <ChevronDown className="pointer-events-none absolute right-2.5 top-1/2 -translate-y-1/2 h-4 w-4 text-somus-text-tertiary" />
      </div>
    );
  }
);

Select.displayName = 'Select';

export default Select;

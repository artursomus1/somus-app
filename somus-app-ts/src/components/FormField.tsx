import React from 'react';
import { HelpCircle } from 'lucide-react';
import { cn } from '@/utils/cn';

export interface FormFieldProps {
  label: string;
  error?: string;
  required?: boolean;
  tooltip?: string;
  className?: string;
  children: React.ReactNode;
}

export function FormField({
  label,
  error,
  required = false,
  tooltip,
  className,
  children,
}: FormFieldProps) {
  return (
    <div className={cn('flex flex-col gap-1.5', className)}>
      <label className="flex items-center gap-1.5 text-sm font-medium text-somus-text-secondary">
        {label}
        {required && <span className="text-red-400">*</span>}
        {tooltip && (
          <span className="group relative">
            <HelpCircle className="h-3.5 w-3.5 text-somus-text-tertiary cursor-help" />
            <span className="absolute bottom-full left-1/2 -translate-x-1/2 mb-1.5 hidden group-hover:block bg-somus-bg-tertiary text-somus-text-primary text-xs rounded-md px-2.5 py-1.5 whitespace-nowrap shadow-lg border border-somus-border z-50">
              {tooltip}
            </span>
          </span>
        )}
      </label>
      {children}
      {error && (
        <p className="text-xs text-red-400 mt-0.5">{error}</p>
      )}
    </div>
  );
}

export default FormField;

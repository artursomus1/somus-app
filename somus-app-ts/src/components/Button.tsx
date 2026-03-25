import React from 'react';
import { Loader2 } from 'lucide-react';
import { cn } from '@/utils/cn';

export interface ButtonProps extends React.ButtonHTMLAttributes<HTMLButtonElement> {
  variant?: 'primary' | 'secondary' | 'success' | 'danger' | 'ghost';
  size?: 'sm' | 'md' | 'lg';
  loading?: boolean;
  icon?: React.ReactNode;
  children: React.ReactNode;
  fullWidth?: boolean;
}

const variantStyles: Record<NonNullable<ButtonProps['variant']>, string> = {
  primary:
    'text-white shadow-sm',
  secondary:
    'bg-transparent text-somus-text-secondary border border-somus-border hover:bg-somus-bg-hover hover:text-somus-text-primary hover:border-somus-border-light active:bg-somus-bg-tertiary',
  success:
    'text-white shadow-sm',
  danger:
    'bg-red-700 text-white hover:bg-red-600 active:bg-red-800 shadow-sm',
  ghost:
    'bg-transparent text-somus-text-secondary hover:bg-somus-bg-hover hover:text-somus-text-primary active:bg-somus-bg-tertiary',
};

const sizeStyles: Record<NonNullable<ButtonProps['size']>, string> = {
  sm: 'text-xs px-3 py-1.5 gap-1.5',
  md: 'text-sm px-4 py-2 gap-2',
  lg: 'text-base px-6 py-2.5 gap-2.5',
};

export const Button = React.forwardRef<HTMLButtonElement, ButtonProps>(
  (
    {
      variant = 'primary',
      size = 'md',
      loading = false,
      icon,
      children,
      fullWidth = false,
      disabled,
      className,
      style,
      ...props
    },
    ref
  ) => {
    // Green gradient for primary and success variants
    const gradientStyle =
      variant === 'primary' || variant === 'success'
        ? {
            background: 'linear-gradient(135deg, #1A7A3E 0%, #0D5C2C 100%)',
            ...style,
          }
        : style;

    return (
      <button
        ref={ref}
        disabled={disabled || loading}
        className={cn(
          'inline-flex items-center justify-center font-medium rounded-lg transition-all duration-150',
          'focus:outline-none focus:ring-2 focus:ring-somus-green-500/30 focus:ring-offset-1 focus:ring-offset-somus-bg-primary',
          'disabled:opacity-50 disabled:cursor-not-allowed select-none',
          'active:scale-[0.97]',
          variantStyles[variant],
          sizeStyles[size],
          fullWidth && 'w-full',
          (variant === 'primary' || variant === 'success') && 'hover:shadow-glow-green',
          className
        )}
        style={gradientStyle}
        {...props}
      >
        {loading ? (
          <Loader2 className="h-4 w-4 animate-spin" />
        ) : icon ? (
          <span className="shrink-0">{icon}</span>
        ) : null}
        {children}
      </button>
    );
  }
);

Button.displayName = 'Button';

export default Button;

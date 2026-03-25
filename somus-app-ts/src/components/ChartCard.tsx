import React from 'react';
import {
  ResponsiveContainer,
  LineChart,
  Line,
  BarChart,
  Bar,
  AreaChart,
  Area,
  PieChart,
  Pie,
  Cell,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  Legend,
} from 'recharts';
import { cn } from '@/utils/cn';

// ─── NASA-inspired chart colors for dark theme ──────────────────────
export const CHART_COLORS = [
  '#1A7A3E', // somus green (primary)
  '#7030A0', // purple (core data)
  '#D4A017', // gold (highlights)
  '#002060', // navy (comparisons)
  '#00B0F0', // sky blue (consolidation)
  '#ED7D31', // orange (alerts)
  '#1B6B5F', // teal (VPL)
  '#C00000', // red (negative)
];

// ─── Types ───────────────────────────────────────────────────────────

export interface ChartSeries {
  dataKey: string;
  name?: string;
  color?: string;
}

export interface ChartCardProps {
  title: string;
  subtitle?: string;
  type: 'line' | 'bar' | 'area' | 'pie';
  data: Record<string, any>[];
  series: ChartSeries[];
  xAxisKey?: string;
  height?: number;
  headerRight?: React.ReactNode;
  className?: string;
  /** Pie chart specific: data key for the value */
  pieValueKey?: string;
  /** Pie chart specific: data key for the name/label */
  pieNameKey?: string;
  /** Custom tooltip formatter for values (e.g. formatBRL) */
  valueFormatter?: (value: number) => string;
}

// ─── Custom Tooltip - Dark theme ────────────────────────────────────

function SomusTooltip({
  active,
  payload,
  label,
  valueFormatter,
}: any) {
  if (!active || !payload?.length) return null;

  return (
    <div className="rounded-lg shadow-lg px-3 py-2 border" style={{ background: '#0F1419', borderColor: '#1A7A3E40' }}>
      {label && (
        <p className="text-xs font-medium mb-1" style={{ color: '#8B95A5' }}>{label}</p>
      )}
      {payload.map((entry: any, i: number) => (
        <div key={i} className="flex items-center gap-2 text-sm">
          <span
            className="h-2.5 w-2.5 rounded-full shrink-0"
            style={{ backgroundColor: entry.color }}
          />
          <span style={{ color: '#8B95A5' }}>{entry.name}:</span>
          <span className="font-semibold" style={{ color: '#E8ECF0' }}>
            {valueFormatter
              ? valueFormatter(entry.value)
              : typeof entry.value === 'number'
              ? entry.value.toLocaleString('pt-BR')
              : entry.value}
          </span>
        </div>
      ))}
    </div>
  );
}

// ─── Component ───────────────────────────────────────────────────────

export function ChartCard({
  title,
  subtitle,
  type,
  data,
  series,
  xAxisKey = 'name',
  height = 300,
  headerRight,
  className,
  pieValueKey,
  pieNameKey,
  valueFormatter,
}: ChartCardProps) {
  const renderChart = () => {
    const commonTooltip = (
      <Tooltip content={<SomusTooltip valueFormatter={valueFormatter} />} />
    );

    const commonGrid = (
      <CartesianGrid strokeDasharray="3 3" stroke="#1E2A3A" vertical={false} />
    );

    const commonXAxis = (
      <XAxis
        dataKey={xAxisKey}
        tick={{ fontSize: 12, fill: '#8B95A5' }}
        tickLine={false}
        axisLine={{ stroke: '#1E2A3A' }}
      />
    );

    const commonYAxis = (
      <YAxis
        tick={{ fontSize: 12, fill: '#8B95A5' }}
        tickLine={false}
        axisLine={false}
        tickFormatter={(v: number) =>
          valueFormatter ? valueFormatter(v) : v.toLocaleString('pt-BR')
        }
      />
    );

    switch (type) {
      case 'line':
        return (
          <LineChart data={data}>
            {commonGrid}
            {commonXAxis}
            {commonYAxis}
            {commonTooltip}
            <Legend
              iconType="circle"
              iconSize={8}
              wrapperStyle={{ fontSize: 12, paddingTop: 8, color: '#8B95A5' }}
            />
            {series.map((s, i) => (
              <Line
                key={s.dataKey}
                type="monotone"
                dataKey={s.dataKey}
                name={s.name ?? s.dataKey}
                stroke={s.color ?? CHART_COLORS[i % CHART_COLORS.length]}
                strokeWidth={2}
                dot={{ r: 3 }}
                activeDot={{ r: 5 }}
              />
            ))}
          </LineChart>
        );

      case 'bar':
        return (
          <BarChart data={data}>
            {commonGrid}
            {commonXAxis}
            {commonYAxis}
            {commonTooltip}
            <Legend
              iconType="circle"
              iconSize={8}
              wrapperStyle={{ fontSize: 12, paddingTop: 8, color: '#8B95A5' }}
            />
            {series.map((s, i) => (
              <Bar
                key={s.dataKey}
                dataKey={s.dataKey}
                name={s.name ?? s.dataKey}
                fill={s.color ?? CHART_COLORS[i % CHART_COLORS.length]}
                radius={[4, 4, 0, 0]}
                maxBarSize={48}
              />
            ))}
          </BarChart>
        );

      case 'area':
        return (
          <AreaChart data={data}>
            {commonGrid}
            {commonXAxis}
            {commonYAxis}
            {commonTooltip}
            <Legend
              iconType="circle"
              iconSize={8}
              wrapperStyle={{ fontSize: 12, paddingTop: 8, color: '#8B95A5' }}
            />
            {series.map((s, i) => {
              const color = s.color ?? CHART_COLORS[i % CHART_COLORS.length];
              return (
                <Area
                  key={s.dataKey}
                  type="monotone"
                  dataKey={s.dataKey}
                  name={s.name ?? s.dataKey}
                  stroke={color}
                  fill={color}
                  fillOpacity={0.15}
                  strokeWidth={2}
                />
              );
            })}
          </AreaChart>
        );

      case 'pie': {
        const vKey = pieValueKey ?? series[0]?.dataKey ?? 'value';
        const nKey = pieNameKey ?? xAxisKey;
        return (
          <PieChart>
            <Pie
              data={data}
              dataKey={vKey}
              nameKey={nKey}
              cx="50%"
              cy="50%"
              outerRadius={height / 3}
              innerRadius={height / 5}
              paddingAngle={2}
              label={({ name, percent }: any) =>
                `${name} ${(percent * 100).toFixed(1)}%`
              }
              labelLine={{ stroke: '#5A6577' }}
            >
              {data.map((_, i) => (
                <Cell
                  key={i}
                  fill={
                    series[i]?.color ?? CHART_COLORS[i % CHART_COLORS.length]
                  }
                />
              ))}
            </Pie>
            {commonTooltip}
            <Legend
              iconType="circle"
              iconSize={8}
              wrapperStyle={{ fontSize: 12, color: '#8B95A5' }}
            />
          </PieChart>
        );
      }

      default:
        return null;
    }
  };

  return (
    <div
      className={cn(
        'bg-somus-bg-secondary/80 backdrop-blur-xl rounded-xl border border-somus-border/50',
        'transition-all duration-200',
        className
      )}
    >
      {/* Header */}
      <div className="flex items-center justify-between px-5 pt-5 pb-2">
        <div>
          <h3 className="text-base font-semibold text-somus-text-primary">
            {title}
          </h3>
          {subtitle && (
            <p className="text-sm text-somus-text-secondary mt-0.5">{subtitle}</p>
          )}
        </div>
        {headerRight && <div>{headerRight}</div>}
      </div>

      {/* Chart */}
      <div className="px-3 pb-4">
        <ResponsiveContainer width="100%" height={height}>
          {renderChart()!}
        </ResponsiveContainer>
      </div>
    </div>
  );
}

export default ChartCard;

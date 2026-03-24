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

// ─── Somus branding colors for charts ────────────────────────────────
export const CHART_COLORS = [
  '#004D33', // somus green
  '#005C3D', // somus green light
  '#0EA5E9', // sky-500
  '#F59E0B', // amber-500
  '#EF4444', // red-500
  '#8B5CF6', // violet-500
  '#EC4899', // pink-500
  '#14B8A6', // teal-500
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

// ─── Custom Tooltip ──────────────────────────────────────────────────

function SomusTooltip({
  active,
  payload,
  label,
  valueFormatter,
}: any) {
  if (!active || !payload?.length) return null;

  return (
    <div className="bg-white rounded-lg shadow-lg border border-somus-gray-200 px-3 py-2">
      {label && (
        <p className="text-xs font-medium text-somus-gray-500 mb-1">{label}</p>
      )}
      {payload.map((entry: any, i: number) => (
        <div key={i} className="flex items-center gap-2 text-sm">
          <span
            className="h-2.5 w-2.5 rounded-full shrink-0"
            style={{ backgroundColor: entry.color }}
          />
          <span className="text-somus-gray-600">{entry.name}:</span>
          <span className="font-semibold text-somus-gray-900">
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
      <CartesianGrid strokeDasharray="3 3" stroke="#E5E7EB" vertical={false} />
    );

    const commonXAxis = (
      <XAxis
        dataKey={xAxisKey}
        tick={{ fontSize: 12, fill: '#6B7280' }}
        tickLine={false}
        axisLine={{ stroke: '#E5E7EB' }}
      />
    );

    const commonYAxis = (
      <YAxis
        tick={{ fontSize: 12, fill: '#6B7280' }}
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
              wrapperStyle={{ fontSize: 12, paddingTop: 8 }}
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
              wrapperStyle={{ fontSize: 12, paddingTop: 8 }}
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
              wrapperStyle={{ fontSize: 12, paddingTop: 8 }}
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
              labelLine={{ stroke: '#9CA3AF' }}
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
              wrapperStyle={{ fontSize: 12 }}
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
        'bg-white rounded-lg shadow-sm border border-somus-gray-200',
        className
      )}
    >
      {/* Header */}
      <div className="flex items-center justify-between px-5 pt-5 pb-2">
        <div>
          <h3 className="text-base font-semibold text-somus-gray-900">
            {title}
          </h3>
          {subtitle && (
            <p className="text-sm text-somus-gray-500 mt-0.5">{subtitle}</p>
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

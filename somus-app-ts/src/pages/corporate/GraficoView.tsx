import React, { useMemo } from 'react';
import { LineChart as LineChartIcon } from 'lucide-react';
import {
  ResponsiveContainer,
  AreaChart,
  Area,
  LineChart,
  Line,
  BarChart,
  Bar,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  Legend,
  ReferenceLine,
  Brush,
} from 'recharts';
import { useAppStore } from '@/stores/appStore';
import { NasaEngine } from '@engine/index';
import type { FluxoResult, FluxoMensal } from '@engine/nasa-engine';

// ── Helpers ─────────────────────────────────────────────────────────────────

function fmtBRL(v: number): string {
  return v.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
}

function DarkTooltip({ active, payload, label }: any) {
  if (!active || !payload?.length) return null;
  return (
    <div className="bg-somus-bg-secondary border border-somus-border rounded-lg px-3 py-2 shadow-lg">
      <p className="text-xs font-medium text-somus-text-secondary mb-1">Mês {label}</p>
      {payload.map((e: any, i: number) => (
        <div key={i} className="flex items-center gap-2 text-xs">
          <span className="h-2 w-2 rounded-full" style={{ backgroundColor: e.color }} />
          <span className="text-somus-text-secondary">{e.name}:</span>
          <span className="font-semibold text-somus-text-primary">{fmtBRL(e.value)}</span>
        </div>
      ))}
    </div>
  );
}

// ── Main Component ──────────────────────────────────────────────────────────

export default function GraficoView() {
  const setPage = useAppStore((s) => s.setPage);

  const engine = useMemo(() => new NasaEngine(), []);
  const result = useMemo<FluxoResult>(() => {
    return engine.calcularFluxoCompleto({
      valor_credito: 500000,
      prazo_meses: 200,
      taxa_adm_pct: 20,
      fundo_reserva_pct: 3,
      seguro_vida_pct: 0.05,
      momento_contemplacao: 36,
      lance_embutido_pct: 10,
      lance_livre_pct: 20,
      reajuste_pre_pct: 7,
      reajuste_pos_pct: 7,
      reajuste_pre_freq: 'Anual',
      reajuste_pos_freq: 'Anual',
    });
  }, [engine]);

  const contemp = 36;
  const fluxo = result.fluxo;

  // Main chart data: Carta de Crédito vs Saldo Devedor
  const mainData = useMemo(() => {
    return fluxo
      .filter((f: FluxoMensal) => f.mes > 0)
      .map((f: FluxoMensal) => ({
        mes: f.mes,
        carta: Math.round(f.carta_credito_reajustada),
        saldo: Math.round(f.saldo_devedor_reajustado),
      }));
  }, [fluxo]);

  // Parcelas over time
  const parcelasData = useMemo(() => {
    return fluxo
      .filter((f: FluxoMensal) => f.mes > 0 && f.mes % 2 === 0)
      .map((f: FluxoMensal) => ({
        mes: f.mes,
        parcela: Math.round(f.parcela_com_seguro),
      }));
  }, [fluxo]);

  // Cumulative payments
  const cumulativeData = useMemo(() => {
    let acum = 0;
    return fluxo
      .filter((f: FluxoMensal) => f.mes > 0)
      .map((f: FluxoMensal) => {
        acum += f.parcela_com_seguro + (f.outros_custos ?? 0);
        return { mes: f.mes, acumulado: Math.round(acum) };
      })
      .filter((_: any, i: number) => i % 3 === 0);
  }, [fluxo]);

  // Composition breakdown (stacked bar)
  const compositionData = useMemo(() => {
    return fluxo
      .filter((f: FluxoMensal) => f.mes > 0 && f.mes % 6 === 0)
      .map((f: FluxoMensal) => ({
        mes: f.mes,
        fc: Math.round(Math.abs(f.amortizacao) * f.fator_reajuste),
        ta: Math.round(f.valor_parcela_ta * f.fator_reajuste),
        fr: Math.round(f.valor_fundo_reserva * f.fator_reajuste),
        seguro: Math.round(f.seguro_vida),
      }));
  }, [fluxo]);

  return (
    <div className="flex flex-col h-full w-full overflow-hidden">
      <header className="sticky top-0 z-20 bg-somus-bg-primary/90 backdrop-blur-md border-b border-somus-border px-6 py-3">
        <div className="flex items-center gap-3">
          <LineChartIcon size={20} className="text-somus-gold" />
          <div>
            <h1 className="text-lg font-semibold text-somus-text-primary">Gráfico</h1>
            <p className="text-xs text-somus-text-tertiary">Visualização interativa - espelha aba "Gráfico" da NASA HD</p>
          </div>
        </div>
      </header>

      <main className="flex-1 overflow-y-auto bg-somus-bg-primary p-5 space-y-5">
        {/* Main chart: Carta de Crédito vs Saldo Devedor */}
        <div className="bg-somus-bg-secondary border border-somus-border rounded-lg p-5">
          <h3 className="text-sm font-semibold text-somus-text-primary mb-4">Carta de Crédito vs Saldo Devedor</h3>
          <ResponsiveContainer width="100%" height={400}>
            <AreaChart data={mainData}>
              <CartesianGrid strokeDasharray="3 3" stroke="#1E2A3A" vertical={false} />
              <XAxis dataKey="mes" tick={{ fontSize: 10, fill: '#5A6577' }} tickLine={false} axisLine={{ stroke: '#1E2A3A' }} />
              <YAxis tick={{ fontSize: 10, fill: '#5A6577' }} tickLine={false} axisLine={false} tickFormatter={(v: number) => `${(v / 1000).toFixed(0)}k`} />
              <Tooltip content={<DarkTooltip />} />
              <Legend iconType="circle" iconSize={8} wrapperStyle={{ fontSize: 11, color: '#8B95A5' }} />
              <ReferenceLine x={contemp} stroke="#D4A017" strokeDasharray="5 5" label={{ value: `Contemplação (T=${contemp})`, fill: '#D4A017', fontSize: 10, position: 'top' }} />
              <Area type="monotone" dataKey="carta" name="Carta de Crédito" stroke="#1A7A3E" fill="#1A7A3E" fillOpacity={0.15} strokeWidth={2} />
              <Line type="monotone" dataKey="saldo" name="Saldo Devedor" stroke="#7030A0" strokeWidth={2} dot={false} />
              <Brush dataKey="mes" height={20} stroke="#2A3544" fill="#0F1419" travellerWidth={8} />
            </AreaChart>
          </ResponsiveContainer>
        </div>

        {/* 3 secondary charts */}
        <div className="grid grid-cols-1 lg:grid-cols-3 gap-5">
          {/* Parcelas over time */}
          <div className="bg-somus-bg-secondary border border-somus-border rounded-lg p-4">
            <h4 className="text-xs font-semibold text-somus-text-primary mb-3">Parcelas ao Longo do Tempo</h4>
            <ResponsiveContainer width="100%" height={220}>
              <BarChart data={parcelasData}>
                <CartesianGrid strokeDasharray="3 3" stroke="#1E2A3A" vertical={false} />
                <XAxis dataKey="mes" tick={{ fontSize: 9, fill: '#5A6577' }} tickLine={false} axisLine={false} />
                <YAxis tick={{ fontSize: 9, fill: '#5A6577' }} tickLine={false} axisLine={false} tickFormatter={(v: number) => `${(v / 1000).toFixed(0)}k`} />
                <Tooltip content={<DarkTooltip />} />
                <ReferenceLine x={contemp} stroke="#D4A017" strokeDasharray="3 3" />
                <Bar dataKey="parcela" name="Parcela" fill="#ED7D31" radius={[2, 2, 0, 0]} maxBarSize={6} />
              </BarChart>
            </ResponsiveContainer>
          </div>

          {/* Cumulative payments */}
          <div className="bg-somus-bg-secondary border border-somus-border rounded-lg p-4">
            <h4 className="text-xs font-semibold text-somus-text-primary mb-3">Pagamentos Acumulados</h4>
            <ResponsiveContainer width="100%" height={220}>
              <AreaChart data={cumulativeData}>
                <CartesianGrid strokeDasharray="3 3" stroke="#1E2A3A" vertical={false} />
                <XAxis dataKey="mes" tick={{ fontSize: 9, fill: '#5A6577' }} tickLine={false} axisLine={false} />
                <YAxis tick={{ fontSize: 9, fill: '#5A6577' }} tickLine={false} axisLine={false} tickFormatter={(v: number) => `${(v / 1e6).toFixed(1)}M`} />
                <Tooltip content={<DarkTooltip />} />
                <ReferenceLine x={contemp} stroke="#D4A017" strokeDasharray="3 3" />
                <Area type="monotone" dataKey="acumulado" name="Acumulado" stroke="#0EA5E9" fill="#0EA5E9" fillOpacity={0.15} strokeWidth={2} />
              </AreaChart>
            </ResponsiveContainer>
          </div>

          {/* Composition breakdown */}
          <div className="bg-somus-bg-secondary border border-somus-border rounded-lg p-4">
            <h4 className="text-xs font-semibold text-somus-text-primary mb-3">Composição (FC + TA + FR + Seguro)</h4>
            <ResponsiveContainer width="100%" height={220}>
              <BarChart data={compositionData}>
                <CartesianGrid strokeDasharray="3 3" stroke="#1E2A3A" vertical={false} />
                <XAxis dataKey="mes" tick={{ fontSize: 9, fill: '#5A6577' }} tickLine={false} axisLine={false} />
                <YAxis tick={{ fontSize: 9, fill: '#5A6577' }} tickLine={false} axisLine={false} tickFormatter={(v: number) => `${(v / 1000).toFixed(0)}k`} />
                <Tooltip content={<DarkTooltip />} />
                <Legend iconType="circle" iconSize={6} wrapperStyle={{ fontSize: 9, color: '#8B95A5' }} />
                <Bar dataKey="fc" name="Fundo Comum" stackId="comp" fill="#1A7A3E" maxBarSize={12} />
                <Bar dataKey="ta" name="Taxa Adm" stackId="comp" fill="#0EA5E9" maxBarSize={12} />
                <Bar dataKey="fr" name="Fdo Reserva" stackId="comp" fill="#F59E0B" maxBarSize={12} />
                <Bar dataKey="seguro" name="Seguro" stackId="comp" fill="#EF4444" radius={[2, 2, 0, 0]} maxBarSize={12} />
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>
      </main>
    </div>
  );
}

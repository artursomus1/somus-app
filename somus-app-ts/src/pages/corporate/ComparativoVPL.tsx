import React, { useState } from 'react';
import { useForm, Controller } from 'react-hook-form';
import { zodResolver } from '@hookform/resolvers/zod';
import { z } from 'zod';
import {
  ArrowLeft,
  Calculator,
  CheckCircle2,
  XCircle,
  Target,
} from 'lucide-react';
import { PageLayout } from '@components/PageLayout';
import { Card } from '@components/Card';
import { Button } from '@components/Button';
import { FormField } from '@components/FormField';
import { CurrencyInput } from '@components/CurrencyInput';
import { PercentInput } from '@components/PercentInput';
import { KPICard } from '@components/KPICard';
import { Select } from '@components/Select';
import { ChartCard, CHART_COLORS } from '@components/ChartCard';
import { useAppStore } from '@/stores/appStore';
import {
  calcularFluxoConsorcio,
  calcularVplHd,
  npv,
  monthlyFromAnnual,
} from '@engine/index';
import type { FluxoResult, VPLResult } from '@engine/nasa-engine';

// ── Schema ──────────────────────────────────────────────────────────────────

const vplSchema = z.object({
  valorCarta: z.number().min(1),
  prazoMeses: z.number().min(36).max(420),
  taxaAdm: z.number().min(0).max(100),
  fundoReserva: z.number().min(0).max(100),
  seguro: z.number().min(0).max(100),
  correcaoAnual: z.number().min(0).max(100),
  prazoContemp: z.number().min(1).max(420),
  parcelaRedPct: z.number(),
  lanceLivrePct: z.number().min(0).max(100),
  lanceEmbutidoPct: z.number().min(0).max(100),
  almAnual: z.number().min(0).max(100),
  hurdleAnual: z.number().min(0).max(100),
});

type VPLFormData = z.infer<typeof vplSchema>;

const PRAZO_OPTIONS = [36, 48, 60, 72, 84, 96, 108, 120, 144, 156, 168, 180, 200, 216, 240, 360, 420]
  .map((p) => ({ value: String(p), label: `${p} meses` }));

const PARCELA_RED_OPTIONS = [
  { value: '100', label: '100% (Integral)' },
  { value: '70', label: '70% (Reduzida)' },
  { value: '50', label: '50% (Reduzida)' },
];

const defaultValues: VPLFormData = {
  valorCarta: 500000,
  prazoMeses: 200,
  taxaAdm: 20,
  fundoReserva: 3,
  seguro: 0.05,
  correcaoAnual: 7,
  prazoContemp: 3,
  parcelaRedPct: 70,
  lanceLivrePct: 20,
  lanceEmbutidoPct: 10,
  almAnual: 12,
  hurdleAnual: 12,
};

function fmtBRL(v: number): string {
  return v.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
}

function fmtPct(v: number, d = 2): string {
  return `${v.toLocaleString('pt-BR', { minimumFractionDigits: d, maximumFractionDigits: d })}%`;
}

// ── Build chart data ────────────────────────────────────────────────────────

function buildChartData(cashflow: number[], almAnual: number) {
  const almM = monthlyFromAnnual(almAnual / 100);
  const points: Array<{ mes: number; pvAcumulado: number }> = [];
  let acum = 0;
  for (let t = 0; t < cashflow.length; t++) {
    const pv = cashflow[t] / Math.pow(1 + almM, t);
    acum += pv;
    if (t % 3 === 0 || t === cashflow.length - 1) {
      points.push({ mes: t, pvAcumulado: Math.round(acum * 100) / 100 });
    }
  }
  return points;
}

// ── Main Component ──────────────────────────────────────────────────────────

export default function ComparativoVPL() {
  const setPage = useAppStore((s) => s.setPage);
  const [result, setResult] = useState<VPLResult | null>(null);
  const [chartData, setChartData] = useState<Array<{ mes: number; pvAcumulado: number }>>([]);
  const [loading, setLoading] = useState(false);

  const { control, handleSubmit } = useForm<VPLFormData>({
    resolver: zodResolver(vplSchema),
    defaultValues,
  });

  const inputClass =
    'w-full px-3 py-2 text-sm border border-somus-gray-300 rounded-lg focus:ring-2 focus:ring-somus-green/40 focus:border-somus-green outline-none bg-white';

  function onCalculate(data: VPLFormData) {
    setLoading(true);
    try {
      const engineParams: Record<string, any> = {
        valor_carta: data.valorCarta,
        prazo_meses: data.prazoMeses,
        taxa_adm: data.taxaAdm,
        fundo_reserva: data.fundoReserva,
        seguro: data.seguro,
        prazo_contemp: data.prazoContemp,
        parcela_red_pct: data.parcelaRedPct,
        lance_livre_pct: data.lanceLivrePct,
        lance_embutido_pct: data.lanceEmbutidoPct,
        correcao_anual: data.correcaoAnual,
        alm_anual: data.almAnual,
        hurdle_anual: data.hurdleAnual,
      };

      const fluxoResult = calcularFluxoConsorcio(engineParams) as unknown as FluxoResult;
      const vplResult = calcularVplHd(engineParams, fluxoResult);

      setResult(vplResult);
      setChartData(buildChartData(fluxoResult.cashflow ?? fluxoResult.cashflow_consorcio, data.almAnual));
    } finally {
      setLoading(false);
    }
  }

  return (
    <PageLayout title="Comparativo de VPL" subtitle="Analise VPL Goal-Based (NASA HD)">
      <div className="max-w-7xl mx-auto">
        <div className="mb-4">
          <button onClick={() => setPage('dashboard')} className="inline-flex items-center gap-1.5 text-sm text-somus-gray-500 hover:text-somus-gray-700 transition-colors">
            <ArrowLeft className="h-4 w-4" /> Voltar ao Dashboard
          </button>
        </div>

        <form onSubmit={handleSubmit(onCalculate)}>
          <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
            {/* ─── Left: Inputs ─────────────────────────────────────── */}
            <div className="lg:col-span-1 space-y-5">
              <Card title="Parametros do Consorcio" padding="md">
                <div className="space-y-4 mt-4">
                  <FormField label="Valor da Carta" required>
                    <Controller name="valorCarta" control={control} render={({ field }) => <CurrencyInput value={field.value} onChange={field.onChange} />} />
                  </FormField>
                  <FormField label="Prazo (meses)">
                    <Controller name="prazoMeses" control={control} render={({ field }) => (
                      <Select options={PRAZO_OPTIONS} value={String(field.value)} onChange={(e) => field.onChange(Number(e.target.value))} />
                    )} />
                  </FormField>
                  <div className="grid grid-cols-2 gap-3">
                    <FormField label="Taxa Adm (%)">
                      <Controller name="taxaAdm" control={control} render={({ field }) => <PercentInput value={field.value} onChange={field.onChange} />} />
                    </FormField>
                    <FormField label="Fdo Reserva (%)">
                      <Controller name="fundoReserva" control={control} render={({ field }) => <PercentInput value={field.value} onChange={field.onChange} />} />
                    </FormField>
                  </div>
                  <div className="grid grid-cols-2 gap-3">
                    <FormField label="Seguro (%)">
                      <Controller name="seguro" control={control} render={({ field }) => <PercentInput value={field.value} onChange={field.onChange} decimals={4} />} />
                    </FormField>
                    <FormField label="Correcao (%)">
                      <Controller name="correcaoAnual" control={control} render={({ field }) => <PercentInput value={field.value} onChange={field.onChange} />} />
                    </FormField>
                  </div>
                  <FormField label="Prazo Contemplacao (meses)">
                    <Controller name="prazoContemp" control={control} render={({ field }) => (
                      <input type="number" min={1} max={420} value={field.value} onChange={(e) => field.onChange(Number(e.target.value))} className={inputClass} />
                    )} />
                  </FormField>
                  <div className="grid grid-cols-2 gap-3">
                    <FormField label="Lance Livre (%)">
                      <Controller name="lanceLivrePct" control={control} render={({ field }) => <PercentInput value={field.value} onChange={field.onChange} />} />
                    </FormField>
                    <FormField label="Lance Embutido (%)">
                      <Controller name="lanceEmbutidoPct" control={control} render={({ field }) => <PercentInput value={field.value} onChange={field.onChange} />} />
                    </FormField>
                  </div>
                  <FormField label="Parcela Reduzida (%)">
                    <Controller name="parcelaRedPct" control={control} render={({ field }) => (
                      <Select options={PARCELA_RED_OPTIONS} value={String(field.value)} onChange={(e) => field.onChange(Number(e.target.value))} />
                    )} />
                  </FormField>
                </div>
              </Card>

              <Card title="Taxas de Desconto" padding="md">
                <div className="space-y-4 mt-4">
                  <FormField label="CDI/ALM Anual (%)" tooltip="Custo de oportunidade para desconto pre-contemplacao">
                    <Controller name="almAnual" control={control} render={({ field }) => <PercentInput value={field.value} onChange={field.onChange} />} />
                  </FormField>
                  <FormField label="Hurdle Rate Anual (%)" tooltip="Taxa hurdle para desconto pos-contemplacao">
                    <Controller name="hurdleAnual" control={control} render={({ field }) => <PercentInput value={field.value} onChange={field.onChange} />} />
                  </FormField>
                </div>
              </Card>

              <Button type="submit" variant="success" fullWidth loading={loading} icon={<Calculator className="h-4 w-4" />}>
                Calcular VPL HD
              </Button>
            </div>

            {/* ─── Right: Results ───────────────────────────────────── */}
            <div className="lg:col-span-2 space-y-5">
              {!result ? (
                <div className="flex items-center justify-center h-64 rounded-lg border-2 border-dashed border-somus-gray-200 bg-white">
                  <div className="text-center">
                    <Target className="h-10 w-10 text-somus-gray-300 mx-auto mb-3" />
                    <p className="text-sm text-somus-gray-400">Preencha os parametros e clique em "Calcular VPL HD"</p>
                  </div>
                </div>
              ) : (
                <>
                  {/* Value Badge */}
                  <div className={`rounded-lg p-6 text-center ${result.cria_valor ? 'bg-emerald-50 border-2 border-emerald-300' : 'bg-red-50 border-2 border-red-300'}`}>
                    <div className="flex items-center justify-center gap-3 mb-2">
                      {result.cria_valor
                        ? <CheckCircle2 className="h-8 w-8 text-emerald-600" />
                        : <XCircle className="h-8 w-8 text-red-600" />
                      }
                      <span className={`text-2xl font-bold ${result.cria_valor ? 'text-emerald-700' : 'text-red-700'}`}>
                        {result.cria_valor ? 'CRIA VALOR' : 'DESTROI VALOR'}
                      </span>
                    </div>
                    <p className={`text-lg font-semibold ${result.cria_valor ? 'text-emerald-600' : 'text-red-600'}`}>
                      Delta VPL: {fmtBRL(result.delta_vpl)}
                    </p>
                  </div>

                  {/* KPI Cards */}
                  <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-5 gap-4">
                    <KPICard title="B0 (PV Credito)" value={fmtBRL(result.b0)} variant="green" />
                    <KPICard title="H0 (PV Pgtos)" value={fmtBRL(result.h0)} variant="blue" />
                    <KPICard title="D0 (B0 - H0)" value={fmtBRL(result.d0)} variant={result.d0 >= 0 ? 'green' : 'red'} />
                    <KPICard title="PV pos-T" value={fmtBRL(result.pv_pos_t)} variant="orange" />
                    <KPICard title="Delta VPL" value={fmtBRL(result.delta_vpl)} variant={result.delta_vpl >= 0 ? 'green' : 'red'} />
                  </div>

                  {/* Break-even & TIR */}
                  <div className="grid grid-cols-2 sm:grid-cols-4 gap-4">
                    <KPICard title="Break-even Lance" value={fmtPct(result.break_even_lance, 2)} icon={<Target className="h-5 w-5" />} variant="blue" />
                    <KPICard title="TIR Mensal" value={fmtPct(result.tir_mensal * 100, 4)} />
                    <KPICard title="TIR Anual" value={fmtPct(result.tir_anual * 100, 2)} />
                    <KPICard title="CET Anual" value={fmtPct(result.cet_anual * 100, 2)} variant={result.cet_anual > 0 ? 'red' : 'green'} />
                  </div>

                  {/* PV Breakdown Chart */}
                  {chartData.length > 0 && (
                    <ChartCard
                      title="PV Acumulado por Mes"
                      type="area"
                      data={chartData}
                      series={[{ dataKey: 'pvAcumulado', name: 'PV Acumulado', color: CHART_COLORS[0] }]}
                      xAxisKey="mes"
                      height={280}
                      valueFormatter={(v) => fmtBRL(v)}
                    />
                  )}
                </>
              )}
            </div>
          </div>
        </form>
      </div>
    </PageLayout>
  );
}

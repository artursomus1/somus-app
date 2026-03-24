import React, { useState } from 'react';
import { useForm, Controller } from 'react-hook-form';
import { zodResolver } from '@hookform/resolvers/zod';
import { z } from 'zod';
import { ArrowLeft, Calculator, Award } from 'lucide-react';
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
import { cn } from '@/utils/cn';
import { compararConsorcioFinanciamentoStandalone, calcularFluxoConsorcio } from '@engine/index';

// ── Schemas ─────────────────────────────────────────────────────────────────

const formSchema = z.object({
  // Consorcio
  valorCarta: z.number().min(1),
  prazoMesesC: z.number().min(36).max(420),
  taxaAdm: z.number().min(0).max(100),
  fundoReserva: z.number().min(0).max(100),
  seguro: z.number().min(0).max(100),
  correcaoAnual: z.number().min(0).max(100),
  prazoContemp: z.number().min(1).max(420),
  parcelaRedPct: z.number(),
  lanceLivrePct: z.number().min(0).max(100),
  lanceEmbutidoPct: z.number().min(0).max(100),
  almAnual: z.number().min(0).max(100),
  // Financiamento
  valorFinanc: z.number().min(1),
  prazoMesesF: z.number().min(1).max(600),
  taxaMensalPct: z.number().min(0).max(100),
  metodo: z.enum(['price', 'sac']),
});

type FormData = z.infer<typeof formSchema>;

const PRAZO_OPTIONS = [36, 48, 60, 72, 84, 96, 108, 120, 144, 156, 168, 180, 200, 216, 240, 360, 420]
  .map((p) => ({ value: String(p), label: `${p} meses` }));

const PARCELA_RED_OPTIONS = [
  { value: '100', label: '100%' },
  { value: '70', label: '70%' },
  { value: '50', label: '50%' },
];

const METODO_OPTIONS = [
  { value: 'price', label: 'Price (parcela fixa)' },
  { value: 'sac', label: 'SAC (amortizacao fixa)' },
];

const defaultValues: FormData = {
  valorCarta: 500000,
  prazoMesesC: 200,
  taxaAdm: 20,
  fundoReserva: 3,
  seguro: 0.05,
  correcaoAnual: 7,
  prazoContemp: 3,
  parcelaRedPct: 70,
  lanceLivrePct: 20,
  lanceEmbutidoPct: 10,
  almAnual: 12,
  valorFinanc: 500000,
  prazoMesesF: 360,
  taxaMensalPct: 0.85,
  metodo: 'price',
};

function fmtBRL(v: number): string {
  return v.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
}

function fmtPct(v: number, d = 2): string {
  return `${v.toLocaleString('pt-BR', { minimumFractionDigits: d, maximumFractionDigits: d })}%`;
}

// ── Result type ─────────────────────────────────────────────────────────────

interface CompResult {
  total_pago_consorcio: number;
  total_pago_financiamento: number;
  economia_nominal: number;
  vpl_consorcio: number;
  vpl_financiamento: number;
  economia_vpl: number;
  tir_consorcio_mensal: number;
  tir_consorcio_anual: number;
  tir_financ_mensal: number;
  tir_financ_anual: number;
  pv_consorcio?: number[];
  pv_financiamento?: number[];
  consorcio?: any;
  financiamento?: any;
}

// ── Build chart data ────────────────────────────────────────────────────────

function buildCumulativeChart(result: CompResult): Array<{ mes: number; consorcio: number; financiamento: number }> {
  const cCf = result.consorcio?.cashflow_consorcio ?? result.consorcio?.cashflow ?? [];
  const fCf = result.financiamento?.cashflow ?? [];
  const maxLen = Math.max(cCf.length, fCf.length);
  const points: Array<{ mes: number; consorcio: number; financiamento: number }> = [];
  let acumC = 0;
  let acumF = 0;

  for (let t = 0; t < maxLen; t++) {
    const cfC = t < cCf.length ? cCf[t] : 0;
    const cfF = t < fCf.length ? fCf[t] : 0;
    acumC += cfC < 0 ? Math.abs(cfC) : 0;
    acumF += cfF < 0 ? Math.abs(cfF) : 0;
    if (t % 6 === 0 || t === maxLen - 1) {
      points.push({ mes: t, consorcio: Math.round(acumC), financiamento: Math.round(acumF) });
    }
  }
  return points;
}

// ── Main Component ──────────────────────────────────────────────────────────

export default function ConsorcioVsFinanc() {
  const setPage = useAppStore((s) => s.setPage);
  const [result, setResult] = useState<CompResult | null>(null);
  const [chartData, setChartData] = useState<Array<{ mes: number; consorcio: number; financiamento: number }>>([]);
  const [loading, setLoading] = useState(false);

  const { control, handleSubmit } = useForm<FormData>({
    resolver: zodResolver(formSchema),
    defaultValues,
  });

  const inputClass =
    'w-full px-3 py-2 text-sm border border-somus-gray-300 rounded-lg focus:ring-2 focus:ring-somus-green/40 focus:border-somus-green outline-none bg-white';

  function onCalculate(data: FormData) {
    setLoading(true);
    try {
      const paramsC: Record<string, any> = {
        valor_carta: data.valorCarta,
        prazo_meses: data.prazoMesesC,
        taxa_adm: data.taxaAdm,
        fundo_reserva: data.fundoReserva,
        seguro: data.seguro,
        prazo_contemp: data.prazoContemp,
        parcela_red_pct: data.parcelaRedPct,
        lance_livre_pct: data.lanceLivrePct,
        lance_embutido_pct: data.lanceEmbutidoPct,
        correcao_anual: data.correcaoAnual,
        alm_anual: data.almAnual,
      };

      const paramsF: Record<string, any> = {
        valor: data.valorFinanc,
        prazo_meses: data.prazoMesesF,
        taxa_mensal_pct: data.taxaMensalPct,
        metodo: data.metodo,
      };

      const res = compararConsorcioFinanciamentoStandalone(paramsC, paramsF) as unknown as CompResult;
      setResult(res);
      setChartData(buildCumulativeChart(res));
    } finally {
      setLoading(false);
    }
  }

  const econLabel =
    result && result.economia_vpl >= 0
      ? 'Consorcio mais barato (VPL)'
      : 'Financiamento mais barato (VPL)';

  return (
    <PageLayout title="Consorcio vs Financiamento" subtitle="Compare lado a lado as duas modalidades">
      <div className="max-w-7xl mx-auto">
        <div className="mb-4">
          <button onClick={() => setPage('dashboard')} className="inline-flex items-center gap-1.5 text-sm text-somus-gray-500 hover:text-somus-gray-700 transition-colors">
            <ArrowLeft className="h-4 w-4" /> Voltar ao Dashboard
          </button>
        </div>

        <form onSubmit={handleSubmit(onCalculate)}>
          <div className="grid grid-cols-1 lg:grid-cols-2 gap-6 mb-6">
            {/* Consorcio */}
            <Card title="Consorcio" padding="md" className="border-l-4 border-l-emerald-500">
              <div className="space-y-4 mt-4">
                <FormField label="Valor da Carta" required>
                  <Controller name="valorCarta" control={control} render={({ field }) => <CurrencyInput value={field.value} onChange={field.onChange} />} />
                </FormField>
                <FormField label="Prazo (meses)">
                  <Controller name="prazoMesesC" control={control} render={({ field }) => (
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
                <FormField label="Parcela Reduzida">
                  <Controller name="parcelaRedPct" control={control} render={({ field }) => (
                    <Select options={PARCELA_RED_OPTIONS} value={String(field.value)} onChange={(e) => field.onChange(Number(e.target.value))} />
                  )} />
                </FormField>
                <FormField label="ALM/CDI Anual (%)">
                  <Controller name="almAnual" control={control} render={({ field }) => <PercentInput value={field.value} onChange={field.onChange} />} />
                </FormField>
              </div>
            </Card>

            {/* Financiamento */}
            <Card title="Financiamento" padding="md" className="border-l-4 border-l-blue-500">
              <div className="space-y-4 mt-4">
                <FormField label="Valor Financiado" required>
                  <Controller name="valorFinanc" control={control} render={({ field }) => <CurrencyInput value={field.value} onChange={field.onChange} />} />
                </FormField>
                <FormField label="Prazo (meses)">
                  <Controller name="prazoMesesF" control={control} render={({ field }) => (
                    <input type="number" min={1} max={600} value={field.value} onChange={(e) => field.onChange(Number(e.target.value))} className={inputClass} />
                  )} />
                </FormField>
                <FormField label="Taxa Mensal (%)">
                  <Controller name="taxaMensalPct" control={control} render={({ field }) => <PercentInput value={field.value} onChange={field.onChange} decimals={4} />} />
                </FormField>
                <FormField label="Metodo">
                  <Controller name="metodo" control={control} render={({ field }) => (
                    <Select options={METODO_OPTIONS} value={field.value} onChange={(e) => field.onChange(e.target.value)} />
                  )} />
                </FormField>
              </div>
            </Card>
          </div>

          <Button type="submit" variant="success" fullWidth loading={loading} icon={<Calculator className="h-4 w-4" />} size="lg">
            Comparar
          </Button>
        </form>

        {/* ─── Results ─────────────────────────────────────────────── */}
        {result && (
          <div className="mt-8 space-y-6">
            {/* Economia Highlight */}
            <div className={`rounded-lg p-6 text-center ${result.economia_vpl >= 0 ? 'bg-emerald-50 border-2 border-emerald-300' : 'bg-blue-50 border-2 border-blue-300'}`}>
              <Award className={`h-8 w-8 mx-auto mb-2 ${result.economia_vpl >= 0 ? 'text-emerald-600' : 'text-blue-600'}`} />
              <p className={`text-lg font-bold ${result.economia_vpl >= 0 ? 'text-emerald-700' : 'text-blue-700'}`}>
                {econLabel}
              </p>
              <p className="text-sm text-somus-gray-600 mt-1">
                Economia VPL: {fmtBRL(Math.abs(result.economia_vpl))}
              </p>
            </div>

            {/* KPIs */}
            <div className="grid grid-cols-2 sm:grid-cols-4 gap-4">
              <KPICard title="Total Consorcio" value={fmtBRL(result.total_pago_consorcio)} variant="green" />
              <KPICard title="Total Financ." value={fmtBRL(result.total_pago_financiamento)} variant="blue" />
              <KPICard title="VPL Consorcio" value={fmtBRL(result.vpl_consorcio)} />
              <KPICard title="VPL Financ." value={fmtBRL(result.vpl_financiamento)} />
            </div>

            {/* Comparison Table */}
            <Card title="Comparativo" padding="none">
              <table className="w-full text-sm">
                <thead>
                  <tr className="bg-somus-gray-50 border-b border-somus-gray-200">
                    <th className="text-left px-5 py-3 font-medium text-somus-gray-600">Metrica</th>
                    <th className="text-right px-5 py-3 font-medium text-emerald-700">Consorcio</th>
                    <th className="text-right px-5 py-3 font-medium text-blue-700">Financiamento</th>
                  </tr>
                </thead>
                <tbody>
                  {[
                    ['Total Pago', fmtBRL(result.total_pago_consorcio), fmtBRL(result.total_pago_financiamento)],
                    ['VPL', fmtBRL(result.vpl_consorcio), fmtBRL(result.vpl_financiamento)],
                    ['TIR Mensal', fmtPct(result.tir_consorcio_mensal * 100, 4), fmtPct(result.tir_financ_mensal * 100, 4)],
                    ['TIR Anual', fmtPct(result.tir_consorcio_anual * 100, 2), fmtPct(result.tir_financ_anual * 100, 2)],
                    ['Economia Nominal', fmtBRL(result.economia_nominal), '-'],
                  ].map(([label, c, f]) => (
                    <tr key={label} className="border-b border-somus-gray-100">
                      <td className="px-5 py-3 text-somus-gray-700">{label}</td>
                      <td className="text-right px-5 py-3 font-medium">{c}</td>
                      <td className="text-right px-5 py-3 font-medium">{f}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </Card>

            {/* Cumulative Chart */}
            {chartData.length > 0 && (
              <ChartCard
                title="Pagamentos Acumulados"
                type="line"
                data={chartData}
                series={[
                  { dataKey: 'consorcio', name: 'Consorcio', color: '#059669' },
                  { dataKey: 'financiamento', name: 'Financiamento', color: '#2563EB' },
                ]}
                xAxisKey="mes"
                height={300}
                valueFormatter={(v) => fmtBRL(v)}
              />
            )}
          </div>
        )}
      </div>
    </PageLayout>
  );
}

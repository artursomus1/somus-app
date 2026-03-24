import React, { useState } from 'react';
import { useForm, Controller } from 'react-hook-form';
import { zodResolver } from '@hookform/resolvers/zod';
import { z } from 'zod';
import {
  ArrowLeft,
  Calculator,
  FileDown,
  FileText,
  Mail,
  RotateCcw,
  Wallet,
  Receipt,
  Percent,
  PiggyBank,
} from 'lucide-react';
import { PageLayout } from '@components/PageLayout';
import { Card } from '@components/Card';
import { Button } from '@components/Button';
import { FormField } from '@components/FormField';
import { CurrencyInput } from '@components/CurrencyInput';
import { PercentInput } from '@components/PercentInput';
import { KPICard } from '@components/KPICard';
import { Select } from '@components/Select';
import { useAppStore } from '@/stores/appStore';
import { cn } from '@/utils/cn';
import { calcularFluxoConsorcio } from '@engine/nasa-engine';

// ── Zod Schema ──────────────────────────────────────────────────────────────

const simuladorSchema = z.object({
  nomeCliente: z.string().min(1, 'Nome obrigatorio'),
  assessor: z.string().min(1, 'Assessor obrigatorio'),
  tipoBem: z.string().min(1, 'Selecione o tipo do bem'),
  administradora: z.string().min(1, 'Selecione a administradora'),
  prazoMeses: z.number().min(36).max(420),
  valorCarta: z.number().min(1, 'Valor obrigatorio'),
  taxaAdm: z.number().min(0).max(100),
  fundoReserva: z.number().min(0).max(100),
  seguro: z.number().min(0).max(100),
  correcaoAnual: z.number().min(0).max(100),
  tipoCorrecao: z.string(),
  indiceCorrecao: z.string(),
  prazoContemp: z.number().min(1).max(420),
  parcelaRedPct: z.number(),
  lanceLivrePct: z.number().min(0).max(100),
  lanceEmbutidoPct: z.number().min(0).max(100),
});

type SimuladorFormData = z.infer<typeof simuladorSchema>;

const PRAZO_OPTIONS = [36, 48, 60, 72, 84, 96, 108, 120, 144, 156, 168, 180, 200, 216, 240, 360, 420]
  .map((p) => ({ value: String(p), label: `${p} meses` }));

const TIPO_BEM_OPTIONS = ['Imovel', 'Automovel', 'Servico', 'Caminhao', 'Maquina Agricola', 'Outro']
  .map((t) => ({ value: t, label: t }));

const ADMINISTRADORAS = ['Embracon', 'Rodobens', 'Magalu', 'Porto Seguro', 'Itau', 'Bradesco', 'Outra']
  .map((a) => ({ value: a, label: a }));

const PARCELA_RED_OPTIONS = [
  { value: '100', label: '100% (Integral)' },
  { value: '70', label: '70% (Reduzida 30%)' },
  { value: '50', label: '50% (Reduzida 50%)' },
];

const CORRECAO_OPTIONS = [
  { value: 'Pre-fixado', label: 'Pre-fixado' },
  { value: 'Pos-fixado', label: 'Pos-fixado' },
];

const INDICE_OPTIONS = [
  { value: 'INCC', label: 'INCC' },
  { value: 'IPCA', label: 'IPCA' },
  { value: 'IGP-M', label: 'IGP-M' },
  { value: 'Outro', label: 'Outro' },
];

const defaultValues: SimuladorFormData = {
  nomeCliente: '',
  assessor: '',
  tipoBem: 'Imovel',
  administradora: '',
  prazoMeses: 200,
  valorCarta: 500000,
  taxaAdm: 20,
  fundoReserva: 3,
  seguro: 0.05,
  correcaoAnual: 7,
  tipoCorrecao: 'Pos-fixado',
  indiceCorrecao: 'INCC',
  prazoContemp: 3,
  parcelaRedPct: 70,
  lanceLivrePct: 20,
  lanceEmbutidoPct: 10,
};

// ── Format helpers ──────────────────────────────────────────────────────────

function fmtBRL(v: number): string {
  return v.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
}

function fmtPct(v: number, decimals = 2): string {
  return `${v.toLocaleString('pt-BR', { minimumFractionDigits: decimals, maximumFractionDigits: decimals })}%`;
}

// ── Result type (from engine compat layer) ──────────────────────────────────

interface SimResult {
  fluxo_mensal: Array<{
    mes: number;
    parcela: number;
    fundo_comum: number;
    taxa_adm: number;
    fundo_reserva: number;
    seguro: number;
    lance: number;
    credito: number;
    fluxo_liquido: number;
    fator_correcao: number;
  }>;
  cashflow_consorcio: number[];
  total_pago: number;
  carta_liquida: number;
  lance_livre_valor: number;
  lance_embutido_valor: number;
  parcela_f1_base: number;
  parcela_f2_base: number;
  meses_restantes: number;
  metricas?: {
    tir_mensal: number;
    tir_anual: number;
    cet_anual: number;
  };
}

// ── Main Component ──────────────────────────────────────────────────────────

export default function Simulador() {
  const setPage = useAppStore((s) => s.setPage);
  const [result, setResult] = useState<SimResult | null>(null);
  const [loading, setLoading] = useState(false);

  const {
    register,
    control,
    handleSubmit,
    reset,
    formState: { errors },
  } = useForm<SimuladorFormData>({
    resolver: zodResolver(simuladorSchema),
    defaultValues,
  });

  function onCalculate(data: SimuladorFormData) {
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
      };
      const res = calcularFluxoConsorcio(engineParams) as unknown as SimResult;
      setResult(res);
    } finally {
      setLoading(false);
    }
  }

  function handleClear() {
    reset(defaultValues);
    setResult(null);
  }

  // Composition breakdown from first payment
  const composicao = result?.fluxo_mensal?.[1]
    ? (() => {
        const f = result.fluxo_mensal[1];
        const total = f.parcela;
        return {
          fc: f.fundo_comum,
          ta: f.taxa_adm,
          fr: f.fundo_reserva,
          sg: f.seguro,
          total,
        };
      })()
    : null;

  const tirMensal = result?.metricas?.tir_mensal ?? 0;
  const cetAnual = result?.metricas?.cet_anual ?? 0;

  const inputClass =
    'w-full px-3 py-2 text-sm border border-somus-gray-300 rounded-lg focus:ring-2 focus:ring-somus-green/40 focus:border-somus-green outline-none bg-white';

  return (
    <PageLayout title="Simulador de Consorcio" subtitle="Simule operacoes de consorcio com calculo de CET e fluxo de caixa">
      <div className="max-w-7xl mx-auto">
        <div className="mb-4">
          <button onClick={() => setPage('dashboard')} className="inline-flex items-center gap-1.5 text-sm text-somus-gray-500 hover:text-somus-gray-700 transition-colors">
            <ArrowLeft className="h-4 w-4" />
            Voltar ao Dashboard
          </button>
        </div>

        <form onSubmit={handleSubmit(onCalculate)}>
          <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
            {/* ─── Left Column: Inputs ─────────────────────────────── */}
            <div className="lg:col-span-1 space-y-5">
              {/* Dados do Cliente */}
              <Card title="Dados do Cliente" padding="md">
                <div className="space-y-4 mt-4">
                  <FormField label="Nome do Cliente" required error={errors.nomeCliente?.message}>
                    <input {...register('nomeCliente')} className={inputClass} placeholder="Nome completo" />
                  </FormField>
                  <FormField label="Assessor" required error={errors.assessor?.message}>
                    <input {...register('assessor')} className={inputClass} placeholder="Nome do assessor" />
                  </FormField>
                </div>
              </Card>

              {/* Parametros do Consorcio */}
              <Card title="Parametros do Consorcio" padding="md">
                <div className="space-y-4 mt-4">
                  <FormField label="Tipo do Bem" required>
                    <Controller name="tipoBem" control={control} render={({ field }) => (
                      <Select options={TIPO_BEM_OPTIONS} value={field.value} onChange={(e) => field.onChange(e.target.value)} />
                    )} />
                  </FormField>

                  <FormField label="Administradora" required error={errors.administradora?.message}>
                    <Controller name="administradora" control={control} render={({ field }) => (
                      <Select options={ADMINISTRADORAS} placeholder="Selecione..." value={field.value} onChange={(e) => field.onChange(e.target.value)} />
                    )} />
                  </FormField>

                  <FormField label="Prazo (meses)" required>
                    <Controller name="prazoMeses" control={control} render={({ field }) => (
                      <Select options={PRAZO_OPTIONS} value={String(field.value)} onChange={(e) => field.onChange(Number(e.target.value))} />
                    )} />
                  </FormField>

                  <FormField label="Valor da Carta" required>
                    <Controller name="valorCarta" control={control} render={({ field }) => (
                      <CurrencyInput value={field.value} onChange={field.onChange} />
                    )} />
                  </FormField>

                  <div className="grid grid-cols-2 gap-3">
                    <FormField label="Taxa de Adm (%)" required>
                      <Controller name="taxaAdm" control={control} render={({ field }) => (
                        <PercentInput value={field.value} onChange={field.onChange} />
                      )} />
                    </FormField>
                    <FormField label="Fundo de Reserva (%)">
                      <Controller name="fundoReserva" control={control} render={({ field }) => (
                        <PercentInput value={field.value} onChange={field.onChange} />
                      )} />
                    </FormField>
                  </div>

                  <div className="grid grid-cols-2 gap-3">
                    <FormField label="Seguro (%)">
                      <Controller name="seguro" control={control} render={({ field }) => (
                        <PercentInput value={field.value} onChange={field.onChange} decimals={4} />
                      )} />
                    </FormField>
                    <FormField label="Correcao Anual (%)">
                      <Controller name="correcaoAnual" control={control} render={({ field }) => (
                        <PercentInput value={field.value} onChange={field.onChange} />
                      )} />
                    </FormField>
                  </div>

                  <div className="grid grid-cols-2 gap-3">
                    <FormField label="Tipo Correcao">
                      <Controller name="tipoCorrecao" control={control} render={({ field }) => (
                        <Select options={CORRECAO_OPTIONS} value={field.value} onChange={(e) => field.onChange(e.target.value)} />
                      )} />
                    </FormField>
                    <FormField label="Indice">
                      <Controller name="indiceCorrecao" control={control} render={({ field }) => (
                        <Select options={INDICE_OPTIONS} value={field.value} onChange={(e) => field.onChange(e.target.value)} />
                      )} />
                    </FormField>
                  </div>
                </div>
              </Card>

              {/* Contemplacao e Lances */}
              <Card title="Contemplacao e Lances" padding="md">
                <div className="space-y-4 mt-4">
                  <FormField label="Prazo de Contemplacao (meses)" required>
                    <Controller name="prazoContemp" control={control} render={({ field }) => (
                      <input type="number" min={1} max={420} value={field.value} onChange={(e) => field.onChange(Number(e.target.value))} className={inputClass} />
                    )} />
                  </FormField>

                  <FormField label="Parcela Reduzida (%)" tooltip="100% = integral, 70% = parcela reduzida">
                    <Controller name="parcelaRedPct" control={control} render={({ field }) => (
                      <Select options={PARCELA_RED_OPTIONS} value={String(field.value)} onChange={(e) => field.onChange(Number(e.target.value))} />
                    )} />
                  </FormField>

                  <div className="grid grid-cols-2 gap-3">
                    <FormField label="Lance Livre (%)">
                      <Controller name="lanceLivrePct" control={control} render={({ field }) => (
                        <PercentInput value={field.value} onChange={field.onChange} />
                      )} />
                    </FormField>
                    <FormField label="Lance Embutido (%)">
                      <Controller name="lanceEmbutidoPct" control={control} render={({ field }) => (
                        <PercentInput value={field.value} onChange={field.onChange} />
                      )} />
                    </FormField>
                  </div>
                </div>
              </Card>

              {/* Action Buttons */}
              <div className="flex flex-wrap gap-3">
                <Button type="submit" variant="success" loading={loading} icon={<Calculator className="h-4 w-4" />} className="flex-1">
                  Calcular Simulacao
                </Button>
                <Button type="button" variant="secondary" icon={<RotateCcw className="h-4 w-4" />} onClick={handleClear}>
                  Limpar
                </Button>
              </div>
              {result && (
                <div className="flex flex-wrap gap-3">
                  <Button type="button" variant="primary" icon={<FileDown className="h-4 w-4" />} className="flex-1">
                    Gerar PDF
                  </Button>
                  <Button type="button" variant="secondary" icon={<FileText className="h-4 w-4" />} className="flex-1">
                    Gerar PPTX
                  </Button>
                  <Button type="button" variant="secondary" icon={<Mail className="h-4 w-4" />} className="flex-1">
                    Enviar Email
                  </Button>
                </div>
              )}
            </div>

            {/* ─── Right Column: Results ────────────────────────────── */}
            <div className="lg:col-span-2 space-y-5">
              {!result ? (
                <div className="flex items-center justify-center h-64 rounded-lg border-2 border-dashed border-somus-gray-200 bg-white">
                  <div className="text-center">
                    <Calculator className="h-10 w-10 text-somus-gray-300 mx-auto mb-3" />
                    <p className="text-sm text-somus-gray-400">
                      Preencha os parametros e clique em "Calcular Simulacao"
                    </p>
                  </div>
                </div>
              ) : (
                <>
                  {/* Phase Summary Cards */}
                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                    <div className="rounded-lg border-2 border-emerald-200 bg-emerald-50 p-5">
                      <div className="flex items-center gap-2 mb-2">
                        <span className="inline-flex items-center justify-center h-7 w-7 rounded-full bg-emerald-200 text-emerald-700 text-xs font-bold">F1</span>
                        <span className="text-sm font-semibold text-emerald-800">Fase 1 (Pre-Contemplacao)</span>
                      </div>
                      <p className="text-2xl font-bold text-emerald-700">{fmtBRL(result.parcela_f1_base)}</p>
                      <p className="text-xs text-emerald-600 mt-1">Parcela base mensal</p>
                    </div>
                    <div className="rounded-lg border-2 border-blue-200 bg-blue-50 p-5">
                      <div className="flex items-center gap-2 mb-2">
                        <span className="inline-flex items-center justify-center h-7 w-7 rounded-full bg-blue-200 text-blue-700 text-xs font-bold">F2</span>
                        <span className="text-sm font-semibold text-blue-800">Fase 2 (Pos-Contemplacao)</span>
                      </div>
                      <p className="text-2xl font-bold text-blue-700">{fmtBRL(result.parcela_f2_base)}</p>
                      <p className="text-xs text-blue-600 mt-1">Parcela base mensal</p>
                    </div>
                  </div>

                  {/* KPI Cards */}
                  <div className="grid grid-cols-2 sm:grid-cols-3 gap-4">
                    <KPICard
                      title="Carta Liquida"
                      value={fmtBRL(result.carta_liquida)}
                      icon={<Wallet className="h-5 w-5" />}
                      variant="green"
                    />
                    <KPICard
                      title="Total Desembolsado"
                      value={fmtBRL(result.total_pago)}
                      icon={<Receipt className="h-5 w-5" />}
                    />
                    <KPICard
                      title="CET a.m."
                      value={fmtPct(tirMensal * 100, 4)}
                      icon={<Percent className="h-5 w-5" />}
                      variant={tirMensal > 0 ? 'red' : 'green'}
                    />
                    <KPICard
                      title="CET a.a."
                      value={fmtPct(cetAnual * 100, 2)}
                      variant={cetAnual > 0 ? 'red' : 'green'}
                    />
                    <KPICard
                      title="Lance Livre"
                      value={fmtBRL(result.lance_livre_valor)}
                      icon={<PiggyBank className="h-5 w-5" />}
                      variant="blue"
                    />
                    <KPICard
                      title="Lance Embutido"
                      value={fmtBRL(result.lance_embutido_valor)}
                      variant="orange"
                    />
                  </div>

                  {/* Composition Breakdown */}
                  {composicao && (
                    <Card title="Composicao da Parcela" padding="md">
                      <div className="mt-4 space-y-3">
                        {[
                          { label: 'Fundo Comum', value: composicao.fc, color: 'bg-emerald-500' },
                          { label: 'Taxa Administracao', value: composicao.ta, color: 'bg-blue-500' },
                          { label: 'Fundo Reserva', value: composicao.fr, color: 'bg-amber-500' },
                          { label: 'Seguro', value: composicao.sg, color: 'bg-red-400' },
                        ].map((item) => {
                          const pct = composicao.total > 0 ? (item.value / composicao.total) * 100 : 0;
                          return (
                            <div key={item.label}>
                              <div className="flex items-center justify-between text-sm mb-1">
                                <span className="text-somus-gray-600">{item.label}</span>
                                <span className="font-medium text-somus-gray-900">
                                  {fmtBRL(item.value)} ({fmtPct(pct, 1)})
                                </span>
                              </div>
                              <div className="h-2 bg-somus-gray-100 rounded-full overflow-hidden">
                                <div className={`h-full ${item.color} rounded-full transition-all`} style={{ width: `${pct}%` }} />
                              </div>
                            </div>
                          );
                        })}
                      </div>
                    </Card>
                  )}

                  {/* Payment Schedule Table */}
                  <Card title="Cronograma de Pagamentos" padding="none">
                    <div className="overflow-x-auto max-h-[500px]">
                      <table className="w-full text-sm">
                        <thead className="sticky top-0 z-10">
                          <tr className="bg-somus-gray-50 border-b border-somus-gray-200">
                            <th className="text-left px-4 py-3 font-medium text-somus-gray-600">Mes</th>
                            <th className="text-right px-4 py-3 font-medium text-somus-gray-600">Fundo Comum</th>
                            <th className="text-right px-4 py-3 font-medium text-somus-gray-600">Tx Adm</th>
                            <th className="text-right px-4 py-3 font-medium text-somus-gray-600">Fdo Reserva</th>
                            <th className="text-right px-4 py-3 font-medium text-somus-gray-600">Seguro</th>
                            <th className="text-right px-4 py-3 font-medium text-somus-gray-600">Parcela</th>
                            <th className="text-right px-4 py-3 font-medium text-somus-gray-600">Lance</th>
                            <th className="text-right px-4 py-3 font-medium text-somus-gray-600">Credito</th>
                            <th className="text-right px-4 py-3 font-medium text-somus-gray-600">Fluxo</th>
                          </tr>
                        </thead>
                        <tbody>
                          {result.fluxo_mensal
                            .filter((f) => f.mes > 0)
                            .map((f) => {
                              const hasCredito = f.credito > 0;
                              return (
                                <tr
                                  key={f.mes}
                                  className={cn(
                                    'border-b border-somus-gray-100 hover:bg-somus-gray-50',
                                    hasCredito && 'bg-emerald-50',
                                  )}
                                >
                                  <td className="px-4 py-2.5 font-medium text-somus-gray-900">
                                    {f.mes}
                                    {hasCredito && (
                                      <span className="ml-2 text-[10px] bg-emerald-100 text-emerald-700 px-1.5 py-0.5 rounded-full">
                                        CONTEMP
                                      </span>
                                    )}
                                  </td>
                                  <td className="text-right px-4 py-2.5">{fmtBRL(f.fundo_comum)}</td>
                                  <td className="text-right px-4 py-2.5">{fmtBRL(f.taxa_adm)}</td>
                                  <td className="text-right px-4 py-2.5">{fmtBRL(f.fundo_reserva)}</td>
                                  <td className="text-right px-4 py-2.5">{fmtBRL(f.seguro)}</td>
                                  <td className="text-right px-4 py-2.5 font-medium">{fmtBRL(f.parcela)}</td>
                                  <td className="text-right px-4 py-2.5">{f.lance > 0 ? fmtBRL(f.lance) : '-'}</td>
                                  <td className="text-right px-4 py-2.5 text-emerald-600">
                                    {f.credito > 0 ? fmtBRL(f.credito) : '-'}
                                  </td>
                                  <td className={cn('text-right px-4 py-2.5 font-medium', f.fluxo_liquido >= 0 ? 'text-emerald-600' : 'text-red-600')}>
                                    {fmtBRL(f.fluxo_liquido)}
                                  </td>
                                </tr>
                              );
                            })}
                        </tbody>
                      </table>
                    </div>
                  </Card>
                </>
              )}
            </div>
          </div>
        </form>
      </div>
    </PageLayout>
  );
}

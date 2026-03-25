import React, { useState, useMemo } from 'react';
import { useForm, Controller } from 'react-hook-form';
import { zodResolver } from '@hookform/resolvers/zod';
import { z } from 'zod';
import {
  Calculator,
  ChevronDown,
  ChevronUp,
  FileDown,
  RotateCcw,
  Eye,
  Table,
  BarChart3,
  List,
} from 'lucide-react';
import { useAppStore } from '@/stores/appStore';
import { NasaEngine } from '@engine/index';
import type { FluxoResult, VPLResult, FluxoMensal } from '@engine/nasa-engine';

// ── Format helpers ──────────────────────────────────────────────────────────

function fmtBRL(v: number): string {
  return v.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
}

function fmtPct(v: number, d = 2): string {
  return `${v.toLocaleString('pt-BR', { minimumFractionDigits: d, maximumFractionDigits: d })}%`;
}

// ── Schema ──────────────────────────────────────────────────────────────────

const schema = z.object({
  valorCredito: z.number().min(1),
  prazoMeses: z.number().min(36).max(420),
  taxaAdmPct: z.number().min(0).max(100),
  fundoReservaPct: z.number().min(0).max(100),
  seguroVidaPct: z.number().min(0).max(100),
  periodoInicio: z.number().min(1),
  // Period distribution
  p1Start: z.number().min(1),
  p1End: z.number().min(1),
  p1FcPct: z.number(),
  p1TaPct: z.number(),
  p1FrPct: z.number(),
  p2Start: z.number().optional(),
  p2End: z.number().optional(),
  p2FcPct: z.number().optional(),
  p2TaPct: z.number().optional(),
  p2FrPct: z.number().optional(),
  // Contemplação
  momentoContemplacao: z.number().min(1),
  lanceEmbutidoPct: z.number().min(0).max(100),
  lanceLivrePct: z.number().min(0).max(100),
  // Reajuste
  reajustePrePct: z.number().min(0),
  reajustePosPct: z.number().min(0),
  reajustePreFreq: z.string(),
  reajustePosFreq: z.string(),
  // VPL
  almAnual: z.number().min(0),
  hurdleAnual: z.number().min(0),
  tma: z.number().min(0),
  // Inspection moments
  insp1: z.number(),
  insp2: z.number(),
  insp3: z.number(),
  insp4: z.number(),
});

type FormData = z.infer<typeof schema>;

const FREQ_OPTIONS = ['Mensal', 'Bimestral', 'Trimestral', 'Semestral', 'Anual']
  .map((f) => ({ value: f, label: f }));

const defaults: FormData = {
  valorCredito: 500000,
  prazoMeses: 200,
  taxaAdmPct: 20,
  fundoReservaPct: 3,
  seguroVidaPct: 0.05,
  periodoInicio: 1,
  p1Start: 1,
  p1End: 200,
  p1FcPct: 1.0,
  p1TaPct: 100,
  p1FrPct: 100,
  momentoContemplacao: 36,
  lanceEmbutidoPct: 10,
  lanceLivrePct: 20,
  reajustePrePct: 7,
  reajustePosPct: 7,
  reajustePreFreq: 'Anual',
  reajustePosFreq: 'Anual',
  almAnual: 12,
  hurdleAnual: 12,
  tma: 1,
  insp1: 1,
  insp2: 13,
  insp3: 25,
  insp4: 37,
};

// ── Collapsible Section ─────────────────────────────────────────────────────

function Section({ title, tag, children, defaultOpen = true }: {
  title: string;
  tag: string;
  children: React.ReactNode;
  defaultOpen?: boolean;
}) {
  const [open, setOpen] = useState(defaultOpen);
  return (
    <div className="bg-somus-bg-secondary border border-somus-border rounded-lg overflow-hidden">
      <button
        type="button"
        onClick={() => setOpen(!open)}
        className="w-full flex items-center justify-between px-4 py-3 hover:bg-somus-bg-hover transition-colors"
      >
        <div className="flex items-center gap-2">
          <span className="text-[10px] font-bold bg-somus-green/20 text-somus-green px-1.5 py-0.5 rounded">{tag}</span>
          <span className="text-sm font-semibold text-somus-text-primary">{title}</span>
        </div>
        {open ? <ChevronUp size={16} className="text-somus-text-tertiary" /> : <ChevronDown size={16} className="text-somus-text-tertiary" />}
      </button>
      {open && <div className="px-4 pb-4 pt-2">{children}</div>}
    </div>
  );
}

// ── Compact Input ───────────────────────────────────────────────────────────

function DInput({ label, children, className = '' }: { label: string; children: React.ReactNode; className?: string }) {
  return (
    <div className={className}>
      <label className="block text-[10px] font-medium text-somus-text-secondary uppercase tracking-wider mb-1">{label}</label>
      {children}
    </div>
  );
}

const inputCls = 'w-full px-2.5 py-1.5 text-sm bg-somus-bg-input border border-somus-border rounded-md text-somus-text-primary focus:ring-1 focus:ring-somus-green/50 focus:border-somus-green outline-none';
const selectCls = inputCls;

// ── Main Component ──────────────────────────────────────────────────────────

export default function Simulador() {
  const setPage = useAppStore((s) => s.setPage);
  const [result, setResult] = useState<FluxoResult | null>(null);
  const [vplResult, setVplResult] = useState<VPLResult | null>(null);
  const [loading, setLoading] = useState(false);
  const [activeTab, setActiveTab] = useState<'results' | 'fluxo' | 'parcelas' | 'vpl'>('results');

  const engine = useMemo(() => new NasaEngine(), []);

  const { register, control, handleSubmit, reset, watch, setValue } = useForm<FormData>({
    resolver: zodResolver(schema),
    defaultValues: defaults,
  });

  const watchedContemp = watch('momentoContemplacao');
  const watchedPrazo = watch('prazoMeses');

  function onCalculate(data: FormData) {
    setLoading(true);
    try {
      const periodos: any[] = [
        { start: data.p1Start, end: data.p1End, fc_pct: data.p1FcPct, ta_pct: data.p1TaPct / 100, fr_pct: data.p1FrPct / 100 },
      ];
      if (data.p2Start && data.p2End) {
        periodos.push({
          start: data.p2Start,
          end: data.p2End,
          fc_pct: data.p2FcPct ?? 1.0,
          ta_pct: (data.p2TaPct ?? 100) / 100,
          fr_pct: (data.p2FrPct ?? 100) / 100,
        });
      }

      const params: Record<string, any> = {
        valor_credito: data.valorCredito,
        prazo_meses: data.prazoMeses,
        taxa_adm_pct: data.taxaAdmPct,
        fundo_reserva_pct: data.fundoReservaPct,
        seguro_vida_pct: data.seguroVidaPct,
        momento_contemplacao: data.momentoContemplacao,
        lance_embutido_pct: data.lanceEmbutidoPct,
        lance_livre_pct: data.lanceLivrePct,
        reajuste_pre_pct: data.reajustePrePct,
        reajuste_pos_pct: data.reajustePosPct,
        reajuste_pre_freq: data.reajustePreFreq,
        reajuste_pos_freq: data.reajustePosFreq,
        alm_anual: data.almAnual,
        hurdle_anual: data.hurdleAnual,
        tma: data.tma / 100,
        periodos,
      };

      const fr = engine.calcularFluxoCompleto(params);
      setResult(fr);
      const vr = engine.calcularVPLHD(params, fr);
      setVplResult(vr);
      setActiveTab('results');
    } finally {
      setLoading(false);
    }
  }

  function handleClear() {
    reset(defaults);
    setResult(null);
    setVplResult(null);
  }

  const fluxo = result?.fluxo ?? [];
  const totais = result?.totais;
  const metricas = result?.metricas;

  // Inspection moments
  const inspMeses = [watch('insp1'), watch('insp2'), watch('insp3'), watch('insp4')];
  const inspData = inspMeses.map((m) => {
    const f = fluxo.find((r: FluxoMensal) => r.mes === m);
    return f ? {
      mes: m,
      parcela: f.parcela_apos_reajuste,
      saldo: f.saldo_devedor_reajustado,
      carta: f.carta_credito_reajustada,
    } : { mes: m, parcela: 0, saldo: 0, carta: 0 };
  });

  const lanceEmbValor = (watch('valorCredito') * watch('lanceEmbutidoPct')) / 100;
  const lanceLivreValor = (watch('valorCredito') * watch('lanceLivrePct')) / 100;
  const creditoLiquido = watch('valorCredito') - lanceEmbValor;

  return (
    <div className="flex flex-col h-full w-full overflow-hidden">
      <header className="sticky top-0 z-20 bg-somus-bg-primary/90 backdrop-blur-md border-b border-somus-border px-6 py-3">
        <div className="flex items-center justify-between">
          <div className="flex items-center gap-3">
            <Calculator size={20} className="text-somus-green" />
            <div>
              <h1 className="text-lg font-semibold text-somus-text-primary">Dados do Consórcio</h1>
              <p className="text-xs text-somus-text-tertiary">Simulador completo - espelha a aba "Dados do Consórcio" da NASA HD</p>
            </div>
          </div>
          <div className="flex items-center gap-2">
            <button type="button" onClick={handleClear} className="inline-flex items-center gap-1 px-3 py-1.5 text-xs font-medium bg-somus-bg-secondary border border-somus-border text-somus-text-secondary rounded-lg hover:bg-somus-bg-hover transition-colors">
              <RotateCcw size={12} /> Limpar
            </button>
          </div>
        </div>
      </header>

      <main className="flex-1 overflow-y-auto bg-somus-bg-primary p-5">
        <form onSubmit={handleSubmit(onCalculate)} className="max-w-[1400px] mx-auto space-y-4">

          {/* SECTION A: Main Inputs */}
          <Section title="Parâmetros Principais" tag="A" defaultOpen={true}>
            <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-6 gap-3">
              <DInput label="Valor do Crédito">
                <Controller name="valorCredito" control={control} render={({ field }) => (
                  <input type="number" step={1000} value={field.value} onChange={(e) => field.onChange(Number(e.target.value))} className={inputCls} />
                )} />
              </DInput>
              <DInput label="Prazo (meses)">
                <Controller name="prazoMeses" control={control} render={({ field }) => (
                  <input type="number" min={36} max={420} value={field.value} onChange={(e) => field.onChange(Number(e.target.value))} className={inputCls} />
                )} />
              </DInput>
              <DInput label="Taxa Adm (%)">
                <Controller name="taxaAdmPct" control={control} render={({ field }) => (
                  <input type="number" step={0.1} value={field.value} onChange={(e) => field.onChange(Number(e.target.value))} className={inputCls} />
                )} />
              </DInput>
              <DInput label="Fundo Reserva (%)">
                <Controller name="fundoReservaPct" control={control} render={({ field }) => (
                  <input type="number" step={0.1} value={field.value} onChange={(e) => field.onChange(Number(e.target.value))} className={inputCls} />
                )} />
              </DInput>
              <DInput label="Seguro Vida (%)">
                <Controller name="seguroVidaPct" control={control} render={({ field }) => (
                  <input type="number" step={0.01} value={field.value} onChange={(e) => field.onChange(Number(e.target.value))} className={inputCls} />
                )} />
              </DInput>
              <DInput label="Período Início">
                <Controller name="periodoInicio" control={control} render={({ field }) => (
                  <input type="number" min={1} value={field.value} onChange={(e) => field.onChange(Number(e.target.value))} className={inputCls} />
                )} />
              </DInput>
            </div>
          </Section>

          {/* SECTION B: Composição das Parcelas */}
          <Section title="Composição das Parcelas Mensais" tag="B" defaultOpen={false}>
            <div className="overflow-x-auto">
              <table className="w-full text-xs">
                <thead>
                  <tr className="border-b border-somus-border">
                    <th className="px-2 py-2 text-left text-somus-text-secondary font-medium">Período</th>
                    <th className="px-2 py-2 text-center text-somus-text-secondary font-medium">Início</th>
                    <th className="px-2 py-2 text-center text-somus-text-secondary font-medium">Fim</th>
                    <th className="px-2 py-2 text-center text-somus-text-secondary font-medium">FC Peso</th>
                    <th className="px-2 py-2 text-center text-somus-text-secondary font-medium">TA (%)</th>
                    <th className="px-2 py-2 text-center text-somus-text-secondary font-medium">FR (%)</th>
                  </tr>
                </thead>
                <tbody>
                  <tr className="border-b border-somus-border/50">
                    <td className="px-2 py-2 text-somus-text-primary font-medium">Período 1</td>
                    <td className="px-2 py-1"><Controller name="p1Start" control={control} render={({ field }) => (
                      <input type="number" min={1} value={field.value} onChange={(e) => field.onChange(Number(e.target.value))} className={`${inputCls} text-center`} />
                    )} /></td>
                    <td className="px-2 py-1"><Controller name="p1End" control={control} render={({ field }) => (
                      <input type="number" min={1} value={field.value} onChange={(e) => field.onChange(Number(e.target.value))} className={`${inputCls} text-center`} />
                    )} /></td>
                    <td className="px-2 py-1"><Controller name="p1FcPct" control={control} render={({ field }) => (
                      <input type="number" step={0.1} value={field.value} onChange={(e) => field.onChange(Number(e.target.value))} className={`${inputCls} text-center`} />
                    )} /></td>
                    <td className="px-2 py-1"><Controller name="p1TaPct" control={control} render={({ field }) => (
                      <input type="number" step={1} value={field.value} onChange={(e) => field.onChange(Number(e.target.value))} className={`${inputCls} text-center`} />
                    )} /></td>
                    <td className="px-2 py-1"><Controller name="p1FrPct" control={control} render={({ field }) => (
                      <input type="number" step={1} value={field.value} onChange={(e) => field.onChange(Number(e.target.value))} className={`${inputCls} text-center`} />
                    )} /></td>
                  </tr>
                  <tr className="border-b border-somus-border/50">
                    <td className="px-2 py-2 text-somus-text-primary font-medium">Período 2</td>
                    <td className="px-2 py-1"><Controller name="p2Start" control={control} render={({ field }) => (
                      <input type="number" min={1} value={field.value ?? ''} onChange={(e) => field.onChange(e.target.value ? Number(e.target.value) : undefined)} className={`${inputCls} text-center`} />
                    )} /></td>
                    <td className="px-2 py-1"><Controller name="p2End" control={control} render={({ field }) => (
                      <input type="number" min={1} value={field.value ?? ''} onChange={(e) => field.onChange(e.target.value ? Number(e.target.value) : undefined)} className={`${inputCls} text-center`} />
                    )} /></td>
                    <td className="px-2 py-1"><Controller name="p2FcPct" control={control} render={({ field }) => (
                      <input type="number" step={0.1} value={field.value ?? ''} onChange={(e) => field.onChange(e.target.value ? Number(e.target.value) : undefined)} className={`${inputCls} text-center`} />
                    )} /></td>
                    <td className="px-2 py-1"><Controller name="p2TaPct" control={control} render={({ field }) => (
                      <input type="number" step={1} value={field.value ?? ''} onChange={(e) => field.onChange(e.target.value ? Number(e.target.value) : undefined)} className={`${inputCls} text-center`} />
                    )} /></td>
                    <td className="px-2 py-1"><Controller name="p2FrPct" control={control} render={({ field }) => (
                      <input type="number" step={1} value={field.value ?? ''} onChange={(e) => field.onChange(e.target.value ? Number(e.target.value) : undefined)} className={`${inputCls} text-center`} />
                    )} /></td>
                  </tr>
                </tbody>
              </table>
            </div>
          </Section>

          {/* SECTION C: Contemplação e Lances */}
          <Section title="Contemplação e Lances" tag="C">
            <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-6 gap-3">
              <DInput label="Mês Contemplação">
                <Controller name="momentoContemplacao" control={control} render={({ field }) => (
                  <input type="number" min={1} max={420} value={field.value} onChange={(e) => field.onChange(Number(e.target.value))} className={inputCls} />
                )} />
              </DInput>
              <DInput label="Lance Embutido (%)">
                <Controller name="lanceEmbutidoPct" control={control} render={({ field }) => (
                  <input type="number" step={0.1} min={0} max={100} value={field.value} onChange={(e) => field.onChange(Number(e.target.value))} className={inputCls} />
                )} />
              </DInput>
              <DInput label="Lance Emb. (R$)">
                <input type="text" readOnly value={fmtBRL(lanceEmbValor)} className={`${inputCls} bg-somus-bg-tertiary`} />
              </DInput>
              <DInput label="Lance Livre (%)">
                <Controller name="lanceLivrePct" control={control} render={({ field }) => (
                  <input type="number" step={0.1} min={0} max={100} value={field.value} onChange={(e) => field.onChange(Number(e.target.value))} className={inputCls} />
                )} />
              </DInput>
              <DInput label="Lance Livre (R$)">
                <input type="text" readOnly value={fmtBRL(lanceLivreValor)} className={`${inputCls} bg-somus-bg-tertiary`} />
              </DInput>
              <DInput label="Crédito Líquido">
                <input type="text" readOnly value={fmtBRL(creditoLiquido)} className={`${inputCls} bg-somus-bg-tertiary text-somus-green font-bold`} />
              </DInput>
            </div>
          </Section>

          {/* SECTION E: Reajuste */}
          <Section title="Reajuste" tag="E" defaultOpen={false}>
            <div className="grid grid-cols-2 lg:grid-cols-4 gap-3">
              <DInput label="Pré-T Percentual (% a.a.)">
                <Controller name="reajustePrePct" control={control} render={({ field }) => (
                  <input type="number" step={0.1} value={field.value} onChange={(e) => field.onChange(Number(e.target.value))} className={inputCls} />
                )} />
              </DInput>
              <DInput label="Pré-T Frequência">
                <Controller name="reajustePreFreq" control={control} render={({ field }) => (
                  <select value={field.value} onChange={(e) => field.onChange(e.target.value)} className={selectCls}>
                    {FREQ_OPTIONS.map((o) => <option key={o.value} value={o.value}>{o.label}</option>)}
                  </select>
                )} />
              </DInput>
              <DInput label="Pós-T Percentual (% a.a.)">
                <Controller name="reajustePosPct" control={control} render={({ field }) => (
                  <input type="number" step={0.1} value={field.value} onChange={(e) => field.onChange(Number(e.target.value))} className={inputCls} />
                )} />
              </DInput>
              <DInput label="Pós-T Frequência">
                <Controller name="reajustePosFreq" control={control} render={({ field }) => (
                  <select value={field.value} onChange={(e) => field.onChange(e.target.value)} className={selectCls}>
                    {FREQ_OPTIONS.map((o) => <option key={o.value} value={o.value}>{o.label}</option>)}
                  </select>
                )} />
              </DInput>
            </div>
          </Section>

          {/* SECTION F: VPL Parameters */}
          <Section title="Parâmetros VPL" tag="F" defaultOpen={false}>
            <div className="grid grid-cols-2 lg:grid-cols-3 gap-3">
              <DInput label="ALM / CDI Anual (%)">
                <Controller name="almAnual" control={control} render={({ field }) => (
                  <input type="number" step={0.1} value={field.value} onChange={(e) => field.onChange(Number(e.target.value))} className={inputCls} />
                )} />
              </DInput>
              <DInput label="Hurdle Rate Anual (%)">
                <Controller name="hurdleAnual" control={control} render={({ field }) => (
                  <input type="number" step={0.1} value={field.value} onChange={(e) => field.onChange(Number(e.target.value))} className={inputCls} />
                )} />
              </DInput>
              <DInput label="TMA (% a.m.)">
                <Controller name="tma" control={control} render={({ field }) => (
                  <input type="number" step={0.01} value={field.value} onChange={(e) => field.onChange(Number(e.target.value))} className={inputCls} />
                )} />
              </DInput>
            </div>
          </Section>

          {/* Calculate Button */}
          <div className="flex items-center gap-3">
            <button
              type="submit"
              disabled={loading}
              className="flex-1 inline-flex items-center justify-center gap-2 px-4 py-2.5 text-sm font-semibold bg-somus-green text-white rounded-lg hover:bg-somus-green-light transition-colors disabled:opacity-50"
            >
              <Calculator size={16} />
              {loading ? 'Calculando...' : 'Calcular Simulação'}
            </button>
            {result && (
              <>
                <button type="button" onClick={() => setPage('fluxo-financeiro')} className="inline-flex items-center gap-1.5 px-3 py-2.5 text-xs font-medium bg-somus-bg-secondary border border-somus-border text-somus-text-secondary rounded-lg hover:bg-somus-bg-hover transition-colors">
                  <Table size={14} /> Ver Fluxo Financeiro
                </button>
                <button type="button" onClick={() => setPage('parcelas')} className="inline-flex items-center gap-1.5 px-3 py-2.5 text-xs font-medium bg-somus-bg-secondary border border-somus-border text-somus-text-secondary rounded-lg hover:bg-somus-bg-hover transition-colors">
                  <List size={14} /> Ver Parcelas
                </button>
                <button type="button" onClick={() => setPage('comparativo-vpl')} className="inline-flex items-center gap-1.5 px-3 py-2.5 text-xs font-medium bg-somus-bg-secondary border border-somus-border text-somus-text-secondary rounded-lg hover:bg-somus-bg-hover transition-colors">
                  <BarChart3 size={14} /> Ver VPL
                </button>
              </>
            )}
          </div>

          {/* ─── RESULTS ─────────────────────────────────────────────── */}
          {result && (
            <div className="space-y-4">
              {/* SECTION D: Análise Operacional */}
              <Section title="Análise Operacional" tag="D">
                <div className="grid grid-cols-2 lg:grid-cols-4 gap-3 mb-4">
                  {inspMeses.map((m, i) => (
                    <DInput key={i} label={`Mês Inspeção ${i + 1}`}>
                      <Controller name={`insp${i + 1}` as any} control={control} render={({ field }) => (
                        <input type="number" min={1} value={field.value} onChange={(e) => field.onChange(Number(e.target.value))} className={inputCls} />
                      )} />
                    </DInput>
                  ))}
                </div>
                <div className="overflow-x-auto">
                  <table className="w-full text-xs">
                    <thead>
                      <tr className="border-b border-somus-border">
                        <th className="px-3 py-2 text-left text-somus-text-secondary font-medium">Mês</th>
                        <th className="px-3 py-2 text-right text-somus-text-secondary font-medium">Parcela Reajustada</th>
                        <th className="px-3 py-2 text-right text-somus-text-secondary font-medium">Saldo Devedor Reaj.</th>
                        <th className="px-3 py-2 text-right text-somus-text-secondary font-medium">Carta Crédito Reaj.</th>
                      </tr>
                    </thead>
                    <tbody>
                      {inspData.map((d) => (
                        <tr key={d.mes} className="border-b border-somus-border/50">
                          <td className="px-3 py-2 text-somus-text-primary font-medium">{d.mes}</td>
                          <td className="px-3 py-2 text-right text-somus-text-primary">{fmtBRL(d.parcela)}</td>
                          <td className="px-3 py-2 text-right" style={{ color: '#C00000' }}>{fmtBRL(d.saldo)}</td>
                          <td className="px-3 py-2 text-right text-somus-green">{fmtBRL(d.carta)}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </Section>

              {/* Results KPIs */}
              <div className="grid grid-cols-2 sm:grid-cols-4 lg:grid-cols-8 gap-3">
                <div className="bg-somus-bg-secondary border border-somus-border rounded-lg p-3">
                  <span className="text-[10px] text-somus-text-secondary uppercase">Total Pago</span>
                  <p className="text-sm font-bold text-somus-text-primary mt-1">{fmtBRL(totais?.total_pago ?? 0)}</p>
                </div>
                <div className="bg-somus-bg-secondary border border-somus-border rounded-lg p-3">
                  <span className="text-[10px] text-somus-text-secondary uppercase">Carta Líquida</span>
                  <p className="text-sm font-bold text-somus-green mt-1">{fmtBRL(totais?.carta_liquida ?? 0)}</p>
                </div>
                <div className="bg-somus-bg-secondary border border-somus-border rounded-lg p-3">
                  <span className="text-[10px] text-somus-text-secondary uppercase">TIR Mensal</span>
                  <p className="text-sm font-bold text-somus-gold mt-1">{fmtPct((metricas?.tir_mensal ?? 0) * 100, 4)}</p>
                </div>
                <div className="bg-somus-bg-secondary border border-somus-border rounded-lg p-3">
                  <span className="text-[10px] text-somus-text-secondary uppercase">CET Anual</span>
                  <p className="text-sm font-bold text-somus-gold mt-1">{fmtPct((metricas?.cet_anual ?? 0) * 100, 2)}</p>
                </div>
                <div className="bg-somus-bg-secondary border border-somus-border rounded-lg p-3">
                  <span className="text-[10px] text-somus-text-secondary uppercase">VPL do Fluxo</span>
                  <p className={`text-sm font-bold mt-1 ${(vplResult?.vpl_total ?? 0) >= 0 ? 'text-emerald-400' : 'text-red-400'}`}>{fmtBRL(vplResult?.vpl_total ?? 0)}</p>
                </div>
                <div className="bg-somus-bg-secondary border border-somus-border rounded-lg p-3">
                  <span className="text-[10px] text-somus-text-secondary uppercase">Delta VPL</span>
                  <p className={`text-sm font-bold mt-1 ${(vplResult?.delta_vpl ?? 0) >= 0 ? 'text-emerald-400' : 'text-red-400'}`}>{fmtBRL(vplResult?.delta_vpl ?? 0)}</p>
                </div>
                <div className="bg-somus-bg-secondary border border-somus-border rounded-lg p-3">
                  <span className="text-[10px] text-somus-text-secondary uppercase">Custo Efetivo (TIR)</span>
                  <p className="text-sm font-bold text-somus-gold mt-1">{fmtPct((metricas?.tir_anual ?? 0) * 100, 2)}</p>
                </div>
                <div className="bg-somus-bg-secondary border border-somus-border rounded-lg p-3">
                  <span className="text-[10px] text-somus-text-secondary uppercase">Custo Total %</span>
                  <p className="text-sm font-bold text-somus-text-primary mt-1">{fmtPct(metricas?.custo_total_pct ?? 0, 2)}</p>
                </div>
              </div>

              {/* Parcela Summary */}
              <div className="grid grid-cols-3 gap-3">
                <div className="bg-somus-bg-secondary border border-somus-border rounded-lg p-4">
                  <span className="text-[10px] text-somus-text-secondary uppercase">Parcela Média</span>
                  <p className="text-lg font-bold text-somus-text-primary mt-1">{fmtBRL(metricas?.parcela_media ?? 0)}</p>
                </div>
                <div className="bg-somus-bg-secondary border border-somus-border rounded-lg p-4">
                  <span className="text-[10px] text-somus-text-secondary uppercase">Parcela Máxima</span>
                  <p className="text-lg font-bold text-somus-text-primary mt-1">{fmtBRL(metricas?.parcela_maxima ?? 0)}</p>
                </div>
                <div className="bg-somus-bg-secondary border border-somus-border rounded-lg p-4">
                  <span className="text-[10px] text-somus-text-secondary uppercase">Parcela Mínima</span>
                  <p className="text-lg font-bold text-somus-text-primary mt-1">{fmtBRL(metricas?.parcela_minima ?? 0)}</p>
                </div>
              </div>
            </div>
          )}
        </form>
      </main>
    </div>
  );
}

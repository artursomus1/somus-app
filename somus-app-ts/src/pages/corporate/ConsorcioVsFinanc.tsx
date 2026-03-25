import React, { useState, useMemo } from 'react';
import { useForm, Controller } from 'react-hook-form';
import { zodResolver } from '@hookform/resolvers/zod';
import { z } from 'zod';
import { Calculator, Award, Layers } from 'lucide-react';
import {
  ResponsiveContainer,
  BarChart,
  Bar,
  LineChart,
  Line,
  PieChart,
  Pie,
  Cell,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  Legend,
} from 'recharts';
import { useAppStore } from '@/stores/appStore';
import { NasaEngine } from '@engine/index';
import type { FluxoResult, FinanciamentoResult } from '@engine/nasa-engine';

// ── Helpers ─────────────────────────────────────────────────────────────────

function fmtBRL(v: number): string {
  return v.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
}

function fmtPct(v: number, d = 2): string {
  return `${v.toLocaleString('pt-BR', { minimumFractionDigits: d, maximumFractionDigits: d })}%`;
}

function DarkTooltip({ active, payload, label }: any) {
  if (!active || !payload?.length) return null;
  return (
    <div className="bg-somus-bg-secondary border border-somus-border rounded-lg px-3 py-2 shadow-lg">
      {label != null && <p className="text-xs text-somus-text-secondary mb-1">{label}</p>}
      {payload.map((e: any, i: number) => (
        <div key={i} className="flex items-center gap-2 text-xs">
          <span className="h-2 w-2 rounded-full" style={{ backgroundColor: e.color }} />
          <span className="text-somus-text-secondary">{e.name}:</span>
          <span className="font-semibold text-somus-text-primary">{typeof e.value === 'number' ? fmtBRL(e.value) : e.value}</span>
        </div>
      ))}
    </div>
  );
}

// ── Schema ──────────────────────────────────────────────────────────────────

const formSchema = z.object({
  valorCredito: z.number().min(1),
  prazoMesesC: z.number().min(36).max(420),
  taxaAdmPct: z.number().min(0),
  fundoReservaPct: z.number().min(0),
  seguroVidaPct: z.number().min(0),
  momentoContemplacao: z.number().min(1),
  lanceEmbutidoPct: z.number().min(0),
  lanceLivrePct: z.number().min(0),
  reajustePrePct: z.number().min(0),
  almAnual: z.number().min(0),
  valorFinanc: z.number().min(1),
  prazoMesesF: z.number().min(1).max(600),
  taxaMensalPct: z.number().min(0),
  metodo: z.enum(['price', 'sac']),
});

type FormData = z.infer<typeof formSchema>;

const defaults: FormData = {
  valorCredito: 500000,
  prazoMesesC: 200,
  taxaAdmPct: 20,
  fundoReservaPct: 3,
  seguroVidaPct: 0.05,
  momentoContemplacao: 36,
  lanceEmbutidoPct: 10,
  lanceLivrePct: 20,
  reajustePrePct: 7,
  almAnual: 12,
  valorFinanc: 500000,
  prazoMesesF: 360,
  taxaMensalPct: 0.85,
  metodo: 'price',
};

const inputCls = 'w-full px-2.5 py-1.5 text-sm bg-somus-bg-input border border-somus-border rounded-md text-somus-text-primary focus:ring-1 focus:ring-somus-green/50 focus:border-somus-green outline-none';

function DInput({ label, children }: { label: string; children: React.ReactNode }) {
  return (
    <div>
      <label className="block text-[10px] font-medium text-somus-text-secondary uppercase tracking-wider mb-1">{label}</label>
      {children}
    </div>
  );
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
  consorcio?: any;
  financiamento?: any;
}

// ── Main Component ──────────────────────────────────────────────────────────

export default function ConsorcioVsFinanc() {
  const setPage = useAppStore((s) => s.setPage);
  const [result, setResult] = useState<CompResult | null>(null);
  const [loading, setLoading] = useState(false);
  const [flowMesInicio, setFlowMesInicio] = useState(0);
  const [flowMesFim, setFlowMesFim] = useState(999);

  const engine = useMemo(() => new NasaEngine(), []);

  const { control, handleSubmit } = useForm<FormData>({
    resolver: zodResolver(formSchema),
    defaultValues: defaults,
  });

  function onCalculate(data: FormData) {
    setLoading(true);
    try {
      const paramsC: Record<string, any> = {
        valor_credito: data.valorCredito,
        prazo_meses: data.prazoMesesC,
        taxa_adm_pct: data.taxaAdmPct,
        fundo_reserva_pct: data.fundoReservaPct,
        seguro_vida_pct: data.seguroVidaPct,
        momento_contemplacao: data.momentoContemplacao,
        lance_embutido_pct: data.lanceEmbutidoPct,
        lance_livre_pct: data.lanceLivrePct,
        reajuste_pre_pct: data.reajustePrePct,
        reajuste_pos_pct: data.reajustePrePct,
        reajuste_pre_freq: 'Anual',
        reajuste_pos_freq: 'Anual',
        alm_anual: data.almAnual,
        hurdle_anual: data.almAnual,
      };

      const paramsF: Record<string, any> = {
        valor: data.valorFinanc,
        prazo_meses: data.prazoMesesF,
        taxa_mensal_pct: data.taxaMensalPct,
        metodo: data.metodo,
        calcular_iof: false,
      };

      const res = engine.compararConsorcioFinanciamento(paramsC, paramsF) as unknown as CompResult;
      setResult(res);
    } finally {
      setLoading(false);
    }
  }

  // Chart data
  const cumulativeData = useMemo(() => {
    if (!result) return [];
    const cCf = result.consorcio?.cashflow ?? [];
    const fCf = result.financiamento?.cashflow ?? [];
    const maxLen = Math.max(cCf.length, fCf.length);
    const points: Array<{ mes: number; consorcio: number; financiamento: number }> = [];
    let acumC = 0, acumF = 0;
    for (let t = 0; t < maxLen; t++) {
      if (t < cCf.length) acumC += cCf[t] < 0 ? Math.abs(cCf[t]) : 0;
      if (t < fCf.length) acumF += fCf[t] < 0 ? Math.abs(fCf[t]) : 0;
      if (t % 6 === 0 || t === maxLen - 1) {
        points.push({ mes: t, consorcio: Math.round(acumC), financiamento: Math.round(acumF) });
      }
    }
    return points;
  }, [result]);

  const filteredCumulativeData = useMemo(() => {
    if (flowMesInicio <= 0 && flowMesFim >= 999) return cumulativeData;
    return cumulativeData.filter((p) => p.mes >= flowMesInicio && p.mes <= flowMesFim);
  }, [cumulativeData, flowMesInicio, flowMesFim]);

  const vplPieConsorcio = result ? [
    { name: 'Fundo Comum', value: Math.abs(result.vpl_consorcio * 0.6) },
    { name: 'Custos', value: Math.abs(result.vpl_consorcio * 0.4) },
  ] : [];

  const vplPieFinanc = result ? [
    { name: 'Amortização', value: Math.abs(result.vpl_financiamento * 0.5) },
    { name: 'Juros', value: Math.abs(result.vpl_financiamento * 0.5) },
  ] : [];

  const taxaBarData = result ? [
    { name: 'Consórcio', taxa: result.tir_consorcio_mensal * 100 },
    { name: 'Financiamento', taxa: result.tir_financ_mensal * 100 },
  ] : [];

  const econLabel = result && result.economia_vpl >= 0
    ? 'Consórcio mais barato (VPL)'
    : 'Financiamento mais barato (VPL)';

  // Comparison table metrics
  const compRows = result ? [
    { metrica: 'Total Pago (Nominal)', cons: fmtBRL(result.total_pago_consorcio), fin: fmtBRL(result.total_pago_financiamento), diff: result.total_pago_financiamento - result.total_pago_consorcio },
    { metrica: 'VPL', cons: fmtBRL(result.vpl_consorcio), fin: fmtBRL(result.vpl_financiamento), diff: Math.abs(result.vpl_financiamento) - Math.abs(result.vpl_consorcio) },
    { metrica: 'TIR Mensal', cons: fmtPct(result.tir_consorcio_mensal * 100, 4), fin: fmtPct(result.tir_financ_mensal * 100, 4), diff: result.tir_financ_mensal - result.tir_consorcio_mensal },
    { metrica: 'TIR Anual', cons: fmtPct(result.tir_consorcio_anual * 100, 2), fin: fmtPct(result.tir_financ_anual * 100, 2), diff: result.tir_financ_anual - result.tir_consorcio_anual },
    { metrica: 'Economia Nominal', cons: fmtBRL(result.economia_nominal), fin: '—', diff: result.economia_nominal },
    { metrica: 'Economia VPL', cons: fmtBRL(result.economia_vpl), fin: '—', diff: result.economia_vpl },
  ] : [];

  return (
    <div className="flex flex-col h-full w-full overflow-hidden">
      <header className="sticky top-0 z-20 bg-somus-bg-primary/90 backdrop-blur-md border-b border-somus-border px-6 py-3">
        <div className="flex items-center gap-3">
          <Layers size={20} className="text-somus-purple" />
          <div>
            <h1 className="text-lg font-semibold text-somus-text-primary">Consórcio vs Financiamento</h1>
            <p className="text-xs text-somus-text-tertiary">Comparação lado a lado - espelha aba "Consórcio X Financ."</p>
          </div>
        </div>
      </header>

      <main className="flex-1 overflow-y-auto bg-somus-bg-primary p-5 space-y-5">
        <form onSubmit={handleSubmit(onCalculate)}>
          {/* Two input columns */}
          <div className="grid grid-cols-1 lg:grid-cols-2 gap-5 mb-4">
            {/* Consórcio */}
            <div className="bg-somus-bg-secondary border border-somus-border rounded-lg p-4" style={{ borderLeftColor: '#1A7A3E', borderLeftWidth: 3 }}>
              <h3 className="text-sm font-semibold text-somus-green mb-3">Consórcio</h3>
              <div className="grid grid-cols-2 gap-3">
                <DInput label="Valor Crédito"><Controller name="valorCredito" control={control} render={({ field }) => (
                  <input type="number" step={1000} value={field.value} onChange={(e) => field.onChange(Number(e.target.value))} className={inputCls} />
                )} /></DInput>
                <DInput label="Prazo (meses)"><Controller name="prazoMesesC" control={control} render={({ field }) => (
                  <input type="number" min={36} value={field.value} onChange={(e) => field.onChange(Number(e.target.value))} className={inputCls} />
                )} /></DInput>
                <DInput label="Taxa Adm (%)"><Controller name="taxaAdmPct" control={control} render={({ field }) => (
                  <input type="number" step={0.1} value={field.value} onChange={(e) => field.onChange(Number(e.target.value))} className={inputCls} />
                )} /></DInput>
                <DInput label="Fdo Reserva (%)"><Controller name="fundoReservaPct" control={control} render={({ field }) => (
                  <input type="number" step={0.1} value={field.value} onChange={(e) => field.onChange(Number(e.target.value))} className={inputCls} />
                )} /></DInput>
                <DInput label="Contemplação (mês)"><Controller name="momentoContemplacao" control={control} render={({ field }) => (
                  <input type="number" min={1} value={field.value} onChange={(e) => field.onChange(Number(e.target.value))} className={inputCls} />
                )} /></DInput>
                <DInput label="Correção (% a.a.)"><Controller name="reajustePrePct" control={control} render={({ field }) => (
                  <input type="number" step={0.1} value={field.value} onChange={(e) => field.onChange(Number(e.target.value))} className={inputCls} />
                )} /></DInput>
                <DInput label="Lance Livre (%)"><Controller name="lanceLivrePct" control={control} render={({ field }) => (
                  <input type="number" step={0.1} value={field.value} onChange={(e) => field.onChange(Number(e.target.value))} className={inputCls} />
                )} /></DInput>
                <DInput label="Lance Embutido (%)"><Controller name="lanceEmbutidoPct" control={control} render={({ field }) => (
                  <input type="number" step={0.1} value={field.value} onChange={(e) => field.onChange(Number(e.target.value))} className={inputCls} />
                )} /></DInput>
              </div>
            </div>

            {/* Financiamento */}
            <div className="bg-somus-bg-secondary border border-somus-border rounded-lg p-4" style={{ borderLeftColor: '#0EA5E9', borderLeftWidth: 3 }}>
              <h3 className="text-sm font-semibold text-somus-skyblue mb-3">Financiamento</h3>
              <div className="grid grid-cols-2 gap-3">
                <DInput label="Valor Financiado"><Controller name="valorFinanc" control={control} render={({ field }) => (
                  <input type="number" step={1000} value={field.value} onChange={(e) => field.onChange(Number(e.target.value))} className={inputCls} />
                )} /></DInput>
                <DInput label="Prazo (meses)"><Controller name="prazoMesesF" control={control} render={({ field }) => (
                  <input type="number" min={1} value={field.value} onChange={(e) => field.onChange(Number(e.target.value))} className={inputCls} />
                )} /></DInput>
                <DInput label="Taxa Mensal (%)"><Controller name="taxaMensalPct" control={control} render={({ field }) => (
                  <input type="number" step={0.01} value={field.value} onChange={(e) => field.onChange(Number(e.target.value))} className={inputCls} />
                )} /></DInput>
                <DInput label="Método"><Controller name="metodo" control={control} render={({ field }) => (
                  <select value={field.value} onChange={(e) => field.onChange(e.target.value)} className={inputCls}>
                    <option value="price">Price (parcela fixa)</option>
                    <option value="sac">SAC (amortização fixa)</option>
                  </select>
                )} /></DInput>
              </div>
            </div>
          </div>

          <button type="submit" disabled={loading} className="w-full inline-flex items-center justify-center gap-2 px-4 py-2.5 text-sm font-semibold bg-somus-green text-white rounded-lg hover:bg-somus-green-light transition-colors disabled:opacity-50">
            <Calculator size={16} /> {loading ? 'Comparando...' : 'Comparar'}
          </button>
        </form>

        {result && (
          <div className="space-y-5">
            {/* Economy highlight */}
            <div className={`rounded-lg p-4 text-center ${result.economia_vpl >= 0 ? 'bg-emerald-500/10 border border-emerald-500/30' : 'bg-sky-500/10 border border-sky-500/30'}`}>
              <Award size={24} className={`mx-auto mb-1 ${result.economia_vpl >= 0 ? 'text-emerald-400' : 'text-sky-400'}`} />
              <p className={`text-lg font-bold ${result.economia_vpl >= 0 ? 'text-emerald-400' : 'text-sky-400'}`}>{econLabel}</p>
              <p className="text-sm text-somus-text-secondary mt-1">Economia VPL: {fmtBRL(Math.abs(result.economia_vpl))}</p>
            </div>

            {/* 3 Chart Cards */}
            <div className="grid grid-cols-1 lg:grid-cols-3 gap-5">
              {/* VPL Consórcio Doughnut */}
              <div className="bg-somus-bg-secondary border border-somus-border rounded-lg p-4">
                <h4 className="text-xs font-semibold text-somus-text-primary mb-3">VPL Consórcio</h4>
                <ResponsiveContainer width="100%" height={180}>
                  <PieChart>
                    <Pie data={vplPieConsorcio} dataKey="value" nameKey="name" cx="50%" cy="50%" outerRadius={70} innerRadius={40} paddingAngle={2}>
                      <Cell fill="#1A7A3E" />
                      <Cell fill="#369A5D" />
                    </Pie>
                    <Tooltip content={<DarkTooltip />} />
                  </PieChart>
                </ResponsiveContainer>
                <p className="text-center text-xs text-somus-text-secondary mt-1">{fmtBRL(result.vpl_consorcio)}</p>
              </div>

              {/* VPL Financiamento Doughnut */}
              <div className="bg-somus-bg-secondary border border-somus-border rounded-lg p-4">
                <h4 className="text-xs font-semibold text-somus-text-primary mb-3">VPL Financiamento</h4>
                <ResponsiveContainer width="100%" height={180}>
                  <PieChart>
                    <Pie data={vplPieFinanc} dataKey="value" nameKey="name" cx="50%" cy="50%" outerRadius={70} innerRadius={40} paddingAngle={2}>
                      <Cell fill="#0EA5E9" />
                      <Cell fill="#38BDF8" />
                    </Pie>
                    <Tooltip content={<DarkTooltip />} />
                  </PieChart>
                </ResponsiveContainer>
                <p className="text-center text-xs text-somus-text-secondary mt-1">{fmtBRL(result.vpl_financiamento)}</p>
              </div>

              {/* Taxa Custo Efetivo Bar */}
              <div className="bg-somus-bg-secondary border border-somus-border rounded-lg p-4">
                <h4 className="text-xs font-semibold text-somus-text-primary mb-3">Taxa Custo Efetivo Mensal</h4>
                <ResponsiveContainer width="100%" height={180}>
                  <BarChart data={taxaBarData}>
                    <CartesianGrid strokeDasharray="3 3" stroke="#1E2A3A" vertical={false} />
                    <XAxis dataKey="name" tick={{ fontSize: 10, fill: '#8B95A5' }} tickLine={false} axisLine={false} />
                    <YAxis tick={{ fontSize: 10, fill: '#5A6577' }} tickLine={false} axisLine={false} tickFormatter={(v: number) => `${v.toFixed(2)}%`} />
                    <Tooltip content={<DarkTooltip />} />
                    <Bar dataKey="taxa" name="Taxa Mensal (%)" fill="#D4A017" radius={[4, 4, 0, 0]} maxBarSize={40} />
                  </BarChart>
                </ResponsiveContainer>
              </div>
            </div>

            {/* Comparison Table */}
            <div className="bg-somus-bg-secondary border border-somus-border rounded-lg overflow-hidden">
              <table className="w-full text-xs">
                <thead>
                  <tr className="border-b border-somus-border" style={{ backgroundColor: 'rgba(0, 32, 96, 0.3)' }}>
                    <th className="px-4 py-2.5 text-left text-somus-text-secondary font-medium">Métrica</th>
                    <th className="px-4 py-2.5 text-right font-medium" style={{ color: '#1A7A3E' }}>Consórcio</th>
                    <th className="px-4 py-2.5 text-right font-medium" style={{ color: '#0EA5E9' }}>Financiamento</th>
                    <th className="px-4 py-2.5 text-right text-somus-text-secondary font-medium">Diferença</th>
                  </tr>
                </thead>
                <tbody>
                  {compRows.map((r) => (
                    <tr key={r.metrica} className="border-b border-somus-border/30">
                      <td className="px-4 py-2 text-somus-text-secondary">{r.metrica}</td>
                      <td className="px-4 py-2 text-right font-medium text-somus-text-primary">{r.cons}</td>
                      <td className="px-4 py-2 text-right font-medium text-somus-text-primary">{r.fin}</td>
                      <td className={`px-4 py-2 text-right font-medium ${r.diff >= 0 ? 'text-emerald-400' : 'text-red-400'}`}>
                        {typeof r.diff === 'number' && r.metrica.includes('TIR')
                          ? fmtPct(r.diff * 100, 4)
                          : fmtBRL(r.diff)}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>

            {/* Cumulative payments chart */}
            {cumulativeData.length > 0 && (
              <div className="bg-somus-bg-secondary border border-somus-border rounded-lg p-5">
                <div className="flex items-center justify-between mb-4">
                  <h3 className="text-sm font-semibold text-somus-text-primary">Pagamentos Acumulados</h3>
                  <div className="flex items-center gap-3">
                    <span className="text-xs text-somus-text-secondary">Filtrar:</span>
                    <input
                      type="number"
                      placeholder="Mes inicio"
                      value={flowMesInicio || ''}
                      onChange={(e) => setFlowMesInicio(Number(e.target.value) || 0)}
                      className="w-24 px-2 py-1 text-xs bg-somus-bg-input text-somus-text-primary border border-somus-border rounded"
                    />
                    <span className="text-somus-text-tertiary">a</span>
                    <input
                      type="number"
                      placeholder="Mes fim"
                      value={flowMesFim === 999 ? '' : flowMesFim}
                      onChange={(e) => setFlowMesFim(Number(e.target.value) || 999)}
                      className="w-24 px-2 py-1 text-xs bg-somus-bg-input text-somus-text-primary border border-somus-border rounded"
                    />
                    <button onClick={() => { setFlowMesInicio(0); setFlowMesFim(999); }} className="text-xs text-somus-text-tertiary hover:text-somus-text-primary">Limpar</button>
                  </div>
                </div>
                <ResponsiveContainer width="100%" height={280}>
                  <LineChart data={filteredCumulativeData}>
                    <CartesianGrid strokeDasharray="3 3" stroke="#1E2A3A" vertical={false} />
                    <XAxis dataKey="mes" tick={{ fontSize: 10, fill: '#5A6577' }} tickLine={false} axisLine={{ stroke: '#1E2A3A' }} />
                    <YAxis tick={{ fontSize: 10, fill: '#5A6577' }} tickLine={false} axisLine={false} tickFormatter={(v: number) => `${(v / 1000).toFixed(0)}k`} />
                    <Tooltip content={<DarkTooltip />} />
                    <Legend iconType="circle" iconSize={8} wrapperStyle={{ fontSize: 11, color: '#8B95A5' }} />
                    <Line type="monotone" dataKey="consorcio" name="Consórcio" stroke="#1A7A3E" strokeWidth={2} dot={false} />
                    <Line type="monotone" dataKey="financiamento" name="Financiamento" stroke="#0EA5E9" strokeWidth={2} dot={false} />
                  </LineChart>
                </ResponsiveContainer>
              </div>
            )}
          </div>
        )}
      </main>
    </div>
  );
}

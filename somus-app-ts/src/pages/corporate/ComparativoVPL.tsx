import React, { useState, useMemo } from 'react';
import { useForm, Controller } from 'react-hook-form';
import { zodResolver } from '@hookform/resolvers/zod';
import { z } from 'zod';
import {
  Calculator,
  CheckCircle2,
  XCircle,
  Target,
  TrendingUp,
  AlertTriangle,
} from 'lucide-react';
import {
  ResponsiveContainer,
  AreaChart,
  Area,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  Legend,
  ReferenceLine,
} from 'recharts';
import { useAppStore } from '@/stores/appStore';
import { NasaEngine, monthlyFromAnnual } from '@engine/index';
import type { FluxoResult, VPLResult } from '@engine/nasa-engine';

// ── Format helpers ──────────────────────────────────────────────────────────

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
      <p className="text-xs font-medium text-somus-text-secondary mb-1">Mês {label}</p>
      {payload.map((entry: any, i: number) => (
        <div key={i} className="flex items-center gap-2 text-xs">
          <span className="h-2 w-2 rounded-full shrink-0" style={{ backgroundColor: entry.color }} />
          <span className="text-somus-text-secondary">{entry.name}:</span>
          <span className="font-semibold text-somus-text-primary">{fmtBRL(entry.value)}</span>
        </div>
      ))}
    </div>
  );
}

// ── Schema ──────────────────────────────────────────────────────────────────

const vplSchema = z.object({
  valorCredito: z.number().min(1),
  prazoMeses: z.number().min(36).max(420),
  taxaAdmPct: z.number().min(0).max(100),
  fundoReservaPct: z.number().min(0).max(100),
  seguroVidaPct: z.number().min(0).max(100),
  momentoContemplacao: z.number().min(1),
  lanceEmbutidoPct: z.number().min(0).max(100),
  lanceLivrePct: z.number().min(0).max(100),
  reajustePrePct: z.number().min(0),
  reajustePosPct: z.number().min(0),
  almAnual: z.number().min(0),
  hurdleAnual: z.number().min(0),
  periodoInicio: z.number().min(1),
});

type VPLFormData = z.infer<typeof vplSchema>;

const defaults: VPLFormData = {
  valorCredito: 500000,
  prazoMeses: 200,
  taxaAdmPct: 20,
  fundoReservaPct: 3,
  seguroVidaPct: 0.05,
  momentoContemplacao: 36,
  lanceEmbutidoPct: 10,
  lanceLivrePct: 20,
  reajustePrePct: 7,
  reajustePosPct: 7,
  almAnual: 12,
  hurdleAnual: 12,
  periodoInicio: 1,
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

// ── Main Component ──────────────────────────────────────────────────────────

export default function ComparativoVPL() {
  const setPage = useAppStore((s) => s.setPage);
  const [result, setResult] = useState<VPLResult | null>(null);
  const [fluxoResult, setFluxoResult] = useState<FluxoResult | null>(null);
  const [chartData, setChartData] = useState<Array<{ mes: number; vpPreT: number; vpPosT: number }>>([]);
  const [loading, setLoading] = useState(false);
  const [contemp, setContemp] = useState(36);

  const engine = useMemo(() => new NasaEngine(), []);

  const { control, handleSubmit } = useForm<VPLFormData>({
    resolver: zodResolver(vplSchema),
    defaultValues: defaults,
  });

  function onCalculate(data: VPLFormData) {
    setLoading(true);
    try {
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
        reajuste_pre_freq: 'Anual',
        reajuste_pos_freq: 'Anual',
        alm_anual: data.almAnual,
        hurdle_anual: data.hurdleAnual,
      };

      setContemp(data.momentoContemplacao);
      const fr = engine.calcularFluxoCompleto(params);
      setFluxoResult(fr);
      const vr = engine.calcularVPLHD(params, fr);
      setResult(vr);

      // Build chart
      const almM = monthlyFromAnnual(data.almAnual / 100);
      const points: Array<{ mes: number; vpPreT: number; vpPosT: number }> = [];
      let acumPre = 0;
      let acumPos = 0;
      for (const d of vr.pv_pre_t_detail) {
        acumPre += d.pv;
        points.push({ mes: d.mes, vpPreT: Math.round(acumPre), vpPosT: 0 });
      }
      for (const d of vr.pv_pos_t_detail) {
        acumPos += d.pv;
        if (d.mes % 3 === 0) {
          points.push({ mes: d.mes, vpPreT: Math.round(acumPre), vpPosT: Math.round(acumPos) });
        }
      }
      setChartData(points);
    } finally {
      setLoading(false);
    }
  }

  const [vpMesInicio, setVpMesInicio] = useState(0);
  const [vpMesFim, setVpMesFim] = useState(999);

  // Per-month VP table
  const allVPRows = useMemo(() => {
    if (!result) return [];
    const rows: Array<{ mes: number; pagamento: number; vpPreT: number; vpPosT: number; vpTotal: number }> = [];
    for (const d of result.pv_pre_t_detail) {
      rows.push({ mes: d.mes, pagamento: d.valor, vpPreT: d.pv, vpPosT: 0, vpTotal: d.pv });
    }
    for (const d of result.pv_pos_t_detail) {
      rows.push({ mes: d.mes, pagamento: d.valor, vpPreT: 0, vpPosT: d.pv, vpTotal: d.pv });
    }
    return rows;
  }, [result]);

  const filteredVPRows = useMemo(() => {
    if (vpMesInicio <= 0 && vpMesFim >= 999) return allVPRows;
    return allVPRows.filter((r) => r.mes >= vpMesInicio && r.mes <= vpMesFim);
  }, [allVPRows, vpMesInicio, vpMesFim]);

  const parcelaMediaPos = useMemo(() => {
    if (!result || !result.pv_pos_t_detail.length) return 0;
    return result.pv_pos_t_detail.reduce((s, d) => s + d.valor, 0) / result.pv_pos_t_detail.length;
  }, [result]);

  return (
    <div className="flex flex-col h-full w-full overflow-hidden">
      <header className="sticky top-0 z-20 bg-somus-bg-primary/90 backdrop-blur-md border-b border-somus-border px-6 py-3">
        <div className="flex items-center gap-3">
          <TrendingUp size={20} className="text-somus-skyblue" />
          <div>
            <h1 className="text-lg font-semibold text-somus-text-primary">Comparativo de VPL</h1>
            <p className="text-xs text-somus-text-tertiary">Análise VPL Goal-Based (NASA HD) - espelha aba "COMPARATIVO DE VPL"</p>
          </div>
        </div>
      </header>

      <main className="flex-1 overflow-y-auto bg-somus-bg-primary p-5 space-y-5">
        <form onSubmit={handleSubmit(onCalculate)}>
          {/* TOP: Input parameters */}
          <div className="bg-somus-bg-secondary border border-somus-border rounded-lg p-4 mb-5">
            <div className="grid grid-cols-2 sm:grid-cols-4 lg:grid-cols-7 gap-3">
              <DInput label="Valor Crédito"><Controller name="valorCredito" control={control} render={({ field }) => (
                <input type="number" step={1000} value={field.value} onChange={(e) => field.onChange(Number(e.target.value))} className={inputCls} />
              )} /></DInput>
              <DInput label="T (Contemplação)"><Controller name="momentoContemplacao" control={control} render={({ field }) => (
                <input type="number" min={1} value={field.value} onChange={(e) => field.onChange(Number(e.target.value))} className={inputCls} />
              )} /></DInput>
              <DInput label="Lance Emb. (%)"><Controller name="lanceEmbutidoPct" control={control} render={({ field }) => (
                <input type="number" step={0.1} value={field.value} onChange={(e) => field.onChange(Number(e.target.value))} className={inputCls} />
              )} /></DInput>
              <DInput label="Lance Livre (%)"><Controller name="lanceLivrePct" control={control} render={({ field }) => (
                <input type="number" step={0.1} value={field.value} onChange={(e) => field.onChange(Number(e.target.value))} className={inputCls} />
              )} /></DInput>
              <DInput label="ALM / CDI (% a.a.)"><Controller name="almAnual" control={control} render={({ field }) => (
                <input type="number" step={0.1} value={field.value} onChange={(e) => field.onChange(Number(e.target.value))} className={inputCls} />
              )} /></DInput>
              <DInput label="Hurdle (% a.a.)"><Controller name="hurdleAnual" control={control} render={({ field }) => (
                <input type="number" step={0.1} value={field.value} onChange={(e) => field.onChange(Number(e.target.value))} className={inputCls} />
              )} /></DInput>
              <DInput label="Pgto Início"><Controller name="periodoInicio" control={control} render={({ field }) => (
                <input type="number" min={1} value={field.value} onChange={(e) => field.onChange(Number(e.target.value))} className={inputCls} />
              )} /></DInput>
            </div>
            <div className="grid grid-cols-2 sm:grid-cols-5 gap-3 mt-3">
              <DInput label="Prazo (meses)"><Controller name="prazoMeses" control={control} render={({ field }) => (
                <input type="number" min={36} value={field.value} onChange={(e) => field.onChange(Number(e.target.value))} className={inputCls} />
              )} /></DInput>
              <DInput label="Taxa Adm (%)"><Controller name="taxaAdmPct" control={control} render={({ field }) => (
                <input type="number" step={0.1} value={field.value} onChange={(e) => field.onChange(Number(e.target.value))} className={inputCls} />
              )} /></DInput>
              <DInput label="Fdo Reserva (%)"><Controller name="fundoReservaPct" control={control} render={({ field }) => (
                <input type="number" step={0.1} value={field.value} onChange={(e) => field.onChange(Number(e.target.value))} className={inputCls} />
              )} /></DInput>
              <DInput label="Seguro (%)"><Controller name="seguroVidaPct" control={control} render={({ field }) => (
                <input type="number" step={0.01} value={field.value} onChange={(e) => field.onChange(Number(e.target.value))} className={inputCls} />
              )} /></DInput>
              <button type="submit" disabled={loading} className="self-end inline-flex items-center justify-center gap-2 px-4 py-1.5 text-sm font-semibold bg-somus-green text-white rounded-lg hover:bg-somus-green-light transition-colors disabled:opacity-50">
                <Calculator size={14} /> {loading ? 'Calculando...' : 'Calcular VPL HD'}
              </button>
            </div>
          </div>
        </form>

        {result && (
          <div className="space-y-5">
            {/* Status Badge + Panels Row */}
            <div className="grid grid-cols-1 lg:grid-cols-3 gap-5">
              {/* Left: Resumo Econômico */}
              <div className="lg:col-span-2 bg-somus-bg-secondary border border-somus-border rounded-lg p-5">
                <h3 className="text-sm font-semibold text-somus-text-primary mb-4">Resumo Econômico (Rota A)</h3>
                <div className="grid grid-cols-2 gap-x-6 gap-y-2 text-xs">
                  {[
                    { label: 'Cheque líquido em T', value: fmtBRL(result.b0 * Math.pow(1 + monthlyFromAnnual(0.12), contemp)), color: '' },
                    { label: 'Carta em T', value: fmtBRL(fluxoResult?.fluxo?.[contemp]?.carta_credito_reajustada ?? 0), color: '' },
                    { label: 'Lance próprio em T', value: fmtBRL(fluxoResult?.totais?.lance_livre_valor ?? 0), color: '' },
                    { label: 'ALM mensal', value: fmtPct(monthlyFromAnnual(0.12) * 100, 4), color: '' },
                    { label: 'Hurdle mensal', value: fmtPct(monthlyFromAnnual(0.12) * 100, 4), color: '' },
                    { label: 'B0 (PV Crédito)', value: fmtBRL(result.b0), color: 'text-somus-green' },
                    { label: 'PV parcelas pré-T', value: fmtBRL(result.h0), color: 'text-somus-skyblue' },
                    { label: 'Parcela média pós-T', value: fmtBRL(parcelaMediaPos), color: '' },
                    { label: 'PV parcelas pós-T', value: fmtBRL(result.pv_pos_t), color: 'text-somus-orange' },
                    { label: 'D0 (B0 - H0)', value: fmtBRL(result.d0), color: result.d0 >= 0 ? 'text-emerald-400' : 'text-red-400' },
                    { label: 'Delta-VPL', value: fmtBRL(result.delta_vpl), color: result.delta_vpl >= 0 ? 'text-emerald-400' : 'text-red-400' },
                    { label: 'Break-even Lance', value: fmtPct(result.break_even_lance, 2), color: 'text-somus-skyblue' },
                  ].map((r) => (
                    <div key={r.label} className="flex items-center justify-between py-1.5 border-b border-somus-border/30">
                      <span className="text-somus-text-secondary">{r.label}</span>
                      <span className={`font-semibold ${r.color || 'text-somus-text-primary'}`}>{r.value}</span>
                    </div>
                  ))}
                </div>

                {/* Status Badge */}
                <div className={`mt-4 rounded-lg p-3 text-center ${result.cria_valor ? 'bg-emerald-500/10 border border-emerald-500/30' : 'bg-red-500/10 border border-red-500/30'}`}>
                  <div className="flex items-center justify-center gap-2">
                    {result.cria_valor ? <CheckCircle2 size={20} className="text-emerald-400" /> : <XCircle size={20} className="text-red-400" />}
                    <span className={`text-lg font-bold ${result.cria_valor ? 'text-emerald-400' : 'text-red-400'}`}>
                      {result.cria_valor ? 'CRIA VALOR' : 'NÃO CRIA VALOR'}
                    </span>
                  </div>
                </div>
              </div>

              {/* Right: Guidance + Note */}
              <div className="space-y-4">
                <div className="bg-somus-bg-secondary border border-somus-border rounded-lg p-5">
                  <h3 className="text-sm font-semibold text-somus-text-primary mb-3">Break-even</h3>
                  <p className="text-xs text-somus-text-secondary leading-relaxed">
                    O lance livre de <span className="text-somus-skyblue font-semibold">{fmtPct(result.break_even_lance, 2)}</span> é o ponto
                    em que o Delta-VPL zera. Acima deste lance, a operação passa a destruir valor.
                  </p>
                  <div className="mt-3 grid grid-cols-2 gap-3 text-xs">
                    <div className="bg-somus-bg-tertiary rounded-md p-2">
                      <span className="text-somus-text-tertiary">TIR Mensal</span>
                      <p className="font-bold text-somus-gold">{fmtPct(result.tir_mensal * 100, 4)}</p>
                    </div>
                    <div className="bg-somus-bg-tertiary rounded-md p-2">
                      <span className="text-somus-text-tertiary">TIR Anual</span>
                      <p className="font-bold text-somus-gold">{fmtPct(result.tir_anual * 100, 2)}</p>
                    </div>
                  </div>
                </div>

                <div className="bg-somus-gold/5 border border-somus-gold/20 rounded-lg p-4">
                  <div className="flex items-start gap-2">
                    <AlertTriangle size={16} className="text-somus-gold shrink-0 mt-0.5" />
                    <div>
                      <h4 className="text-xs font-semibold text-somus-gold mb-1">IMPORTANTE</h4>
                      <p className="text-[10px] text-somus-text-secondary leading-relaxed">
                        A análise de VPL utiliza taxas duais: ALM/CDI para desconto pré-contemplação (custo de oportunidade) e
                        Hurdle Rate para pós-contemplação (retorno mínimo exigido). O resultado é sensível a estas premissas.
                      </p>
                    </div>
                  </div>
                </div>
              </div>
            </div>

            {/* Chart: Cumulative VP */}
            {chartData.length > 0 && (
              <div className="bg-somus-bg-secondary border border-somus-border rounded-lg p-5">
                <h3 className="text-sm font-semibold text-somus-text-primary mb-4">VP Acumulado por Mês</h3>
                <ResponsiveContainer width="100%" height={300}>
                  <AreaChart data={chartData}>
                    <CartesianGrid strokeDasharray="3 3" stroke="#1E2A3A" vertical={false} />
                    <XAxis dataKey="mes" tick={{ fontSize: 10, fill: '#5A6577' }} tickLine={false} axisLine={{ stroke: '#1E2A3A' }} />
                    <YAxis tick={{ fontSize: 10, fill: '#5A6577' }} tickLine={false} axisLine={false} tickFormatter={(v: number) => `${(v / 1000).toFixed(0)}k`} />
                    <Tooltip content={<DarkTooltip />} />
                    <Legend iconType="circle" iconSize={8} wrapperStyle={{ fontSize: 11, color: '#8B95A5' }} />
                    <ReferenceLine x={contemp} stroke="#D4A017" strokeDasharray="5 5" label={{ value: `T=${contemp}`, fill: '#D4A017', fontSize: 10 }} />
                    <Area type="monotone" dataKey="vpPreT" name="VP Pré-T @ ALM" stroke="#1A7A3E" fill="#1A7A3E" fillOpacity={0.2} strokeWidth={2} />
                    <Area type="monotone" dataKey="vpPosT" name="VP Pós-T @ Hurdle" stroke="#EF4444" fill="#EF4444" fillOpacity={0.15} strokeWidth={2} />
                  </AreaChart>
                </ResponsiveContainer>
              </div>
            )}

            {/* Per-month VP Table */}
            <div className="bg-somus-bg-secondary border border-somus-border rounded-lg overflow-hidden">
              <div className="px-4 py-3 border-b border-somus-border flex items-center justify-between">
                <h3 className="text-sm font-semibold text-somus-text-primary">Tabela de VP por Mes</h3>
                <span className="text-[10px] text-somus-text-tertiary">{filteredVPRows.length} registros</span>
              </div>
              <div className="flex items-center gap-3 px-4 py-2 border-b border-somus-border/50">
                <span className="text-xs text-somus-text-secondary">Filtrar:</span>
                <input
                  type="number"
                  placeholder="Mes inicio"
                  value={vpMesInicio || ''}
                  onChange={(e) => setVpMesInicio(Number(e.target.value) || 0)}
                  className="w-24 px-2 py-1 text-xs bg-somus-bg-input text-somus-text-primary border border-somus-border rounded"
                />
                <span className="text-somus-text-tertiary">a</span>
                <input
                  type="number"
                  placeholder="Mes fim"
                  value={vpMesFim === 999 ? '' : vpMesFim}
                  onChange={(e) => setVpMesFim(Number(e.target.value) || 999)}
                  className="w-24 px-2 py-1 text-xs bg-somus-bg-input text-somus-text-primary border border-somus-border rounded"
                />
                <button onClick={() => { setVpMesInicio(0); setVpMesFim(999); }} className="text-xs text-somus-text-tertiary hover:text-somus-text-primary">Limpar</button>
              </div>
              <div className="overflow-x-auto max-h-[400px]">
                <table className="w-full text-xs">
                  <thead className="sticky top-0 bg-somus-bg-tertiary">
                    <tr>
                      <th className="px-3 py-2 text-left text-somus-text-secondary font-medium">Mês</th>
                      <th className="px-3 py-2 text-right text-somus-text-secondary font-medium">Pagamento</th>
                      <th className="px-3 py-2 text-right font-medium" style={{ color: '#1A7A3E' }}>VP Pré-T @ ALM</th>
                      <th className="px-3 py-2 text-right font-medium" style={{ color: '#EF4444' }}>VP Pós-T @ Hurdle</th>
                      <th className="px-3 py-2 text-right text-somus-text-secondary font-medium">VP Total</th>
                    </tr>
                  </thead>
                  <tbody>
                    {filteredVPRows.map((r) => (
                      <tr key={r.mes} className={`border-b border-somus-border/30 ${r.mes === contemp ? 'bg-somus-gold/5' : ''}`}>
                        <td className="px-3 py-1.5 text-somus-text-primary font-medium">
                          {r.mes}
                          {r.mes === contemp && <span className="ml-1 text-[9px] text-somus-gold">T</span>}
                        </td>
                        <td className="px-3 py-1.5 text-right text-somus-text-primary">{fmtBRL(r.pagamento)}</td>
                        <td className="px-3 py-1.5 text-right" style={{ color: r.vpPreT > 0 ? '#1A7A3E' : '#5A6577' }}>{r.vpPreT > 0 ? fmtBRL(r.vpPreT) : '—'}</td>
                        <td className="px-3 py-1.5 text-right" style={{ color: r.vpPosT > 0 ? '#EF4444' : '#5A6577' }}>{r.vpPosT > 0 ? fmtBRL(r.vpPosT) : '—'}</td>
                        <td className="px-3 py-1.5 text-right text-somus-text-primary font-medium">{fmtBRL(r.vpTotal)}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        )}
      </main>
    </div>
  );
}

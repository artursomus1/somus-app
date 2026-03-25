import React, { useState, useMemo } from 'react';
import { useForm, Controller } from 'react-hook-form';
import { zodResolver } from '@hookform/resolvers/zod';
import { z } from 'zod';
import {
  Calculator,
  BadgeDollarSign,
  CheckCircle2,
  XCircle,
  ChevronDown,
  ChevronUp,
} from 'lucide-react';
import {
  ResponsiveContainer,
  LineChart,
  Line,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  Legend,
  ReferenceLine,
} from 'recharts';
import { useAppStore } from '@/stores/appStore';
import { NasaEngine, calcularVendaOperacao } from '@engine/index';
// Types inlined to avoid import issues
type FluxoResult = ReturnType<InstanceType<typeof NasaEngine>['calcularFluxoCompleto']>;
type VendaResult = Record<string, any>;

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

const schema = z.object({
  // Consórcio base
  valorCredito: z.number().min(1),
  prazoMeses: z.number().min(36).max(420),
  taxaAdmPct: z.number().min(0).max(100),
  fundoReservaPct: z.number().min(0).max(100),
  seguroVidaPct: z.number().min(0).max(100),
  momentoContemplacao: z.number().min(1),
  lanceEmbutidoPct: z.number().min(0).max(100),
  lanceLivrePct: z.number().min(0).max(100),
  parcelaReduzidaPct: z.number().min(0).max(100),
  correcaoAnual: z.number().min(0),
  // Sale params
  momentoVenda: z.number().min(1),
  valorVenda: z.number().min(0),
  tma: z.number().min(0),
});

type FormData = z.infer<typeof schema>;

const defaults: FormData = {
  valorCredito: 500000,
  prazoMeses: 200,
  taxaAdmPct: 20,
  fundoReservaPct: 3,
  seguroVidaPct: 0.05,
  momentoContemplacao: 36,
  lanceEmbutidoPct: 10,
  lanceLivrePct: 20,
  parcelaReduzidaPct: 0,
  correcaoAnual: 7,
  momentoVenda: 48,
  valorVenda: 350000,
  tma: 1,
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

// ── Section ──────────────────────────────────────────────────────────────────

function Section({ title, tag, children, defaultOpen = true }: {
  title: string; tag: string; children: React.ReactNode; defaultOpen?: boolean;
}) {
  const [open, setOpen] = useState(defaultOpen);
  return (
    <div className="bg-somus-bg-secondary border border-somus-border rounded-lg overflow-hidden">
      <button type="button" onClick={() => setOpen(!open)}
        className="w-full flex items-center justify-between px-4 py-3 hover:bg-somus-bg-hover transition-colors">
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

// ── Main Component ──────────────────────────────────────────────────────────

export default function VendaOperacao() {
  const [result, setResult] = useState<VendaResult | null>(null);
  const [fluxoResult, setFluxoResult] = useState<FluxoResult | null>(null);
  const [loading, setLoading] = useState(false);

  const engine = useMemo(() => new NasaEngine(), []);

  const { control, handleSubmit } = useForm<FormData>({
    resolver: zodResolver(schema),
    defaultValues: defaults,
  });

  function onCalculate(data: FormData) {
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
        reajuste_pre_pct: data.correcaoAnual,
        reajuste_pos_pct: data.correcaoAnual,
        reajuste_pre_freq: 'Anual',
        reajuste_pos_freq: 'Anual',
      };

      const fr = engine.calcularFluxoCompleto(params);
      setFluxoResult(fr);

      const vendaParams = {
        momento_venda: data.momentoVenda,
        valor_venda: data.valorVenda,
        tma: data.tma / 100,
      };

      const vr = calcularVendaOperacao(fr, vendaParams);
      setResult(vr);
    } finally {
      setLoading(false);
    }
  }

  // Build chart data for cumulative flows
  const chartData = useMemo(() => {
    if (!result) return [];
    const points: Array<{ mes: number; vendedor: number; comprador: number }> = [];
    let acumV = 0;
    for (let i = 0; i < result.cashflow_vendedor.length; i++) {
      acumV += result.cashflow_vendedor[i];
      points.push({ mes: i, vendedor: Math.round(acumV), comprador: 0 });
    }
    let acumC = 0;
    for (let i = 0; i < result.cashflow_comprador.length; i++) {
      acumC += result.cashflow_comprador[i];
      const mesReal = (result.cashflow_vendedor.length - 1) + i;
      points.push({ mes: mesReal, vendedor: Math.round(acumV), comprador: Math.round(acumC) });
    }
    return points;
  }, [result]);

  // Build seller table rows
  const sellerRows = useMemo(() => {
    if (!result) return [];
    return result.cashflow_vendedor.map((v: number, i: number) => ({
      mes: i,
      pagamento: v,
      tipo: i === 0 ? 'Crédito' : i === result.cashflow_vendedor.length - 1 && v > 0 ? 'Venda' : v < 0 ? 'Parcela' : 'Lance/Crédito',
    })).filter((r: any) => Math.abs(r.pagamento) > 0.01);
  }, [result]);

  // Build buyer table rows
  const buyerRows = useMemo(() => {
    if (!result) return [];
    return result.cashflow_comprador.map((v: number, i: number) => ({
      mes: i,
      pagamento: v,
      tipo: i === 0 ? 'Compra' : v < 0 ? 'Parcela' : 'Crédito',
    })).filter((r: any) => Math.abs(r.pagamento) > 0.01);
  }, [result]);

  const isLucrativa = result ? result.ganho_nominal > 0 : false;

  return (
    <div className="flex flex-col h-full w-full overflow-hidden">
      <header className="sticky top-0 z-20 bg-somus-bg-primary/90 backdrop-blur-md border-b border-somus-border px-6 py-3">
        <div className="flex items-center gap-3">
          <BadgeDollarSign size={20} className="text-somus-green" />
          <div>
            <h1 className="text-lg font-semibold text-somus-text-primary">Venda de Operação</h1>
            <p className="text-xs text-somus-text-tertiary">Análise de venda - espelha aba "Venda da Operação" da NASA HD</p>
          </div>
        </div>
      </header>

      <main className="flex-1 overflow-y-auto bg-somus-bg-primary p-5 space-y-5">
        <form onSubmit={handleSubmit(onCalculate)} className="space-y-4">
          {/* Consórcio base params */}
          <Section title="Parâmetros do Consórcio" tag="A">
            <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-5 gap-3">
              <DInput label="Valor Crédito"><Controller name="valorCredito" control={control} render={({ field }) => (
                <input type="number" step={1000} value={field.value} onChange={(e) => field.onChange(Number(e.target.value))} className={inputCls} />
              )} /></DInput>
              <DInput label="Prazo (meses)"><Controller name="prazoMeses" control={control} render={({ field }) => (
                <input type="number" min={36} value={field.value} onChange={(e) => field.onChange(Number(e.target.value))} className={inputCls} />
              )} /></DInput>
              <DInput label="Taxa Adm (%)"><Controller name="taxaAdmPct" control={control} render={({ field }) => (
                <input type="number" step={0.1} value={field.value} onChange={(e) => field.onChange(Number(e.target.value))} className={inputCls} />
              )} /></DInput>
              <DInput label="Fundo Reserva (%)"><Controller name="fundoReservaPct" control={control} render={({ field }) => (
                <input type="number" step={0.1} value={field.value} onChange={(e) => field.onChange(Number(e.target.value))} className={inputCls} />
              )} /></DInput>
              <DInput label="Seguro (%)"><Controller name="seguroVidaPct" control={control} render={({ field }) => (
                <input type="number" step={0.01} value={field.value} onChange={(e) => field.onChange(Number(e.target.value))} className={inputCls} />
              )} /></DInput>
              <DInput label="Contemplação (mês)"><Controller name="momentoContemplacao" control={control} render={({ field }) => (
                <input type="number" min={1} value={field.value} onChange={(e) => field.onChange(Number(e.target.value))} className={inputCls} />
              )} /></DInput>
              <DInput label="Lance Embutido (%)"><Controller name="lanceEmbutidoPct" control={control} render={({ field }) => (
                <input type="number" step={0.1} value={field.value} onChange={(e) => field.onChange(Number(e.target.value))} className={inputCls} />
              )} /></DInput>
              <DInput label="Lance Livre (%)"><Controller name="lanceLivrePct" control={control} render={({ field }) => (
                <input type="number" step={0.1} value={field.value} onChange={(e) => field.onChange(Number(e.target.value))} className={inputCls} />
              )} /></DInput>
              <DInput label="Parcela Reduzida (%)"><Controller name="parcelaReduzidaPct" control={control} render={({ field }) => (
                <input type="number" step={0.1} value={field.value} onChange={(e) => field.onChange(Number(e.target.value))} className={inputCls} />
              )} /></DInput>
              <DInput label="Correção Anual (%)"><Controller name="correcaoAnual" control={control} render={({ field }) => (
                <input type="number" step={0.1} value={field.value} onChange={(e) => field.onChange(Number(e.target.value))} className={inputCls} />
              )} /></DInput>
            </div>
          </Section>

          {/* Sale params */}
          <Section title="Parâmetros da Venda" tag="B">
            <div className="grid grid-cols-2 sm:grid-cols-3 gap-3">
              <DInput label="Momento da Venda (mês)"><Controller name="momentoVenda" control={control} render={({ field }) => (
                <input type="number" min={1} value={field.value} onChange={(e) => field.onChange(Number(e.target.value))} className={inputCls} />
              )} /></DInput>
              <DInput label="Valor da Venda (R$)"><Controller name="valorVenda" control={control} render={({ field }) => (
                <input type="number" step={1000} value={field.value} onChange={(e) => field.onChange(Number(e.target.value))} className={inputCls} />
              )} /></DInput>
              <DInput label="TMA (% a.m.)"><Controller name="tma" control={control} render={({ field }) => (
                <input type="number" step={0.01} value={field.value} onChange={(e) => field.onChange(Number(e.target.value))} className={inputCls} />
              )} /></DInput>
            </div>
          </Section>

          {/* Calculate button */}
          <button type="submit" disabled={loading}
            className="w-full inline-flex items-center justify-center gap-2 px-4 py-2.5 text-sm font-semibold bg-somus-green text-white rounded-lg hover:bg-somus-green-light transition-colors disabled:opacity-50">
            <Calculator size={16} /> {loading ? 'Calculando...' : 'Calcular Venda'}
          </button>
        </form>

        {result && (
          <div className="space-y-5">
            {/* Profit/Loss Banner */}
            <div className={`rounded-lg p-4 text-center ${isLucrativa ? 'bg-emerald-500/10 border border-emerald-500/30' : 'bg-red-500/10 border border-red-500/30'}`}>
              <div className="flex items-center justify-center gap-2">
                {isLucrativa ? <CheckCircle2 size={24} className="text-emerald-400" /> : <XCircle size={24} className="text-red-400" />}
                <span className={`text-xl font-bold ${isLucrativa ? 'text-emerald-400' : 'text-red-400'}`}>
                  {isLucrativa ? 'OPERAÇÃO LUCRATIVA' : 'OPERAÇÃO COM PREJUÍZO'}
                </span>
              </div>
              <p className="text-xs text-somus-text-secondary mt-1">
                Ganho nominal: {fmtBRL(result.ganho_nominal)} ({fmtPct(result.ganho_pct)})
              </p>
            </div>

            {/* KPI Cards */}
            <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-6 gap-3">
              <div className="bg-somus-bg-secondary border border-somus-border rounded-lg p-3">
                <span className="text-[10px] text-somus-text-secondary uppercase">VPL Venda</span>
                <p className={`text-sm font-bold mt-1 ${result.vpl_vendedor >= 0 ? 'text-emerald-400' : 'text-red-400'}`}>{fmtBRL(result.vpl_vendedor)}</p>
              </div>
              <div className="bg-somus-bg-secondary border border-somus-border rounded-lg p-3">
                <span className="text-[10px] text-somus-text-secondary uppercase">Ganho %</span>
                <p className={`text-sm font-bold mt-1 ${result.ganho_pct >= 0 ? 'text-emerald-400' : 'text-red-400'}`}>{fmtPct(result.ganho_pct)}</p>
              </div>
              <div className="bg-somus-bg-secondary border border-somus-border rounded-lg p-3">
                <span className="text-[10px] text-somus-text-secondary uppercase">Prazo Médio</span>
                <p className="text-sm font-bold text-somus-text-primary mt-1">{result.prazo_medio.toFixed(0)} meses</p>
              </div>
              <div className="bg-somus-bg-secondary border border-somus-border rounded-lg p-3">
                <span className="text-[10px] text-somus-text-secondary uppercase">Ganho Mensal</span>
                <p className="text-sm font-bold text-somus-gold mt-1">{fmtBRL(result.ganho_mensal)}</p>
              </div>
              <div className="bg-somus-bg-secondary border border-somus-border rounded-lg p-3">
                <span className="text-[10px] text-somus-text-secondary uppercase">Margem Mensal %</span>
                <p className="text-sm font-bold text-somus-gold mt-1">{fmtPct(result.margem_mensal_pct, 4)}</p>
              </div>
              <div className="bg-somus-bg-secondary border border-somus-border rounded-lg p-3">
                <span className="text-[10px] text-somus-text-secondary uppercase">Custo Comprador (TIR)</span>
                <p className="text-sm font-bold text-somus-purple mt-1">{fmtPct(result.tir_comprador_anual * 100, 2)} a.a.</p>
              </div>
            </div>

            {/* Seller/Buyer flow tables side by side */}
            <div className="grid grid-cols-1 lg:grid-cols-2 gap-5">
              {/* Left: Seller Flow */}
              <div className="bg-somus-bg-secondary border border-somus-border rounded-lg overflow-hidden">
                <div className="px-4 py-3 border-b border-somus-border">
                  <h3 className="text-sm font-semibold text-somus-text-primary">Fluxo do Vendedor</h3>
                  <span className="text-[10px] text-somus-text-tertiary">{sellerRows.length} registros</span>
                </div>
                <div className="overflow-x-auto max-h-[350px]">
                  <table className="w-full text-xs">
                    <thead className="sticky top-0 bg-somus-bg-tertiary">
                      <tr>
                        <th className="px-3 py-2 text-left text-somus-text-secondary font-medium">Mês</th>
                        <th className="px-3 py-2 text-right text-somus-text-secondary font-medium">Pagamento</th>
                        <th className="px-3 py-2 text-left text-somus-text-secondary font-medium">Tipo</th>
                      </tr>
                    </thead>
                    <tbody>
                      {sellerRows.map((r: any) => (
                        <tr key={`v-${r.mes}`} className="border-b border-somus-border/30">
                          <td className="px-3 py-1.5 text-somus-text-primary font-medium">{r.mes}</td>
                          <td className={`px-3 py-1.5 text-right font-medium ${r.pagamento >= 0 ? 'text-emerald-400' : 'text-red-400'}`}>{fmtBRL(r.pagamento)}</td>
                          <td className="px-3 py-1.5 text-somus-text-secondary">{r.tipo}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>

              {/* Right: Buyer Flow */}
              <div className="bg-somus-bg-secondary border border-somus-border rounded-lg overflow-hidden">
                <div className="px-4 py-3 border-b border-somus-border">
                  <h3 className="text-sm font-semibold text-somus-text-primary">Fluxo do Comprador</h3>
                  <span className="text-[10px] text-somus-text-tertiary">{buyerRows.length} registros</span>
                </div>
                <div className="overflow-x-auto max-h-[350px]">
                  <table className="w-full text-xs">
                    <thead className="sticky top-0 bg-somus-bg-tertiary">
                      <tr>
                        <th className="px-3 py-2 text-left text-somus-text-secondary font-medium">Mês</th>
                        <th className="px-3 py-2 text-right text-somus-text-secondary font-medium">Pagamento</th>
                        <th className="px-3 py-2 text-left text-somus-text-secondary font-medium">Tipo</th>
                      </tr>
                    </thead>
                    <tbody>
                      {buyerRows.map((r: any, idx: number) => (
                        <tr key={`c-${idx}`} className="border-b border-somus-border/30">
                          <td className="px-3 py-1.5 text-somus-text-primary font-medium">{r.mes}</td>
                          <td className={`px-3 py-1.5 text-right font-medium ${r.pagamento >= 0 ? 'text-emerald-400' : 'text-red-400'}`}>{fmtBRL(r.pagamento)}</td>
                          <td className="px-3 py-1.5 text-somus-text-secondary">{r.tipo}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>

            {/* Chart: Cumulative flows */}
            {chartData.length > 0 && (
              <div className="bg-somus-bg-secondary border border-somus-border rounded-lg p-5">
                <h3 className="text-sm font-semibold text-somus-text-primary mb-4">Fluxo Acumulado - Vendedor vs Comprador</h3>
                <ResponsiveContainer width="100%" height={300}>
                  <LineChart data={chartData}>
                    <CartesianGrid strokeDasharray="3 3" stroke="#1E2A3A" vertical={false} />
                    <XAxis dataKey="mes" tick={{ fontSize: 10, fill: '#5A6577' }} tickLine={false} axisLine={{ stroke: '#1E2A3A' }} />
                    <YAxis tick={{ fontSize: 10, fill: '#5A6577' }} tickLine={false} axisLine={false} tickFormatter={(v: number) => `${(v / 1000).toFixed(0)}k`} />
                    <Tooltip content={<DarkTooltip />} />
                    <Legend iconType="circle" iconSize={8} wrapperStyle={{ fontSize: 11, color: '#8B95A5' }} />
                    <Line type="monotone" dataKey="vendedor" name="Vendedor (Acum.)" stroke="#1A7A3E" strokeWidth={2} dot={false} />
                    <Line type="monotone" dataKey="comprador" name="Comprador (Acum.)" stroke="#7030A0" strokeWidth={2} dot={false} />
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

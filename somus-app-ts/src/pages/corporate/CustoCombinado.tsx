import React, { useState, useMemo } from 'react';
import {
  Calculator,
  Combine,
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
} from 'recharts';
import { calcularCustoCombinado, NasaEngine, calcularCreditoLance } from '@engine/index';
import type { FinanciamentoResult } from '@engine/financiamento';
import type { FluxoResult } from '@engine/nasa-engine';
import type { CustoCombinado as CustoCombinadoType } from '@engine/comparativo';
import { loadData } from '@/services/storage';

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

export default function CustoCombinado() {
  const [result, setResult] = useState<CustoCombinadoType | null>(null);
  const [fluxoConsorcio, setFluxoConsorcio] = useState<FluxoResult | null>(null);
  const [fluxoLance, setFluxoLance] = useState<FinanciamentoResult | null>(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  // Inline lance params for when localStorage doesn't have data
  const [valorLance, setValorLance] = useState(100000);
  const [prazoLance, setPrazoLance] = useState(60);
  const [taxaLance, setTaxaLance] = useState(1.5);
  const [metodoLance, setMetodoLance] = useState<'price' | 'sac'>('price');

  // Inline consorcio params
  const [valorCredito, setValorCredito] = useState(500000);
  const [prazoMeses, setPrazoMeses] = useState(200);
  const [taxaAdmPct, setTaxaAdmPct] = useState(20);
  const [fundoReservaPct, setFundoReservaPct] = useState(3);
  const [seguroVidaPct, setSeguroVidaPct] = useState(0.05);
  const [momentoContemplacao, setMomentoContemplacao] = useState(36);
  const [lanceEmbutidoPct, setLanceEmbutidoPct] = useState(10);
  const [lanceLivrePct, setLanceLivrePct] = useState(20);
  const [correcaoAnual, setCorrecaoAnual] = useState(7);

  const engine = useMemo(() => new NasaEngine(), []);

  function onCalculate() {
    setLoading(true);
    setError(null);
    try {
      // Try to load from localStorage first, otherwise use form values
      let fConsorcio: FluxoResult | null = loadData<FluxoResult | null>('somus-simulador-fluxo', null);
      let fLance: FinanciamentoResult | null = loadData<FinanciamentoResult | null>('somus-credito-lance-result', null);

      // If no stored data, compute from form values
      if (!fConsorcio || !fConsorcio.cashflow) {
        const params: Record<string, any> = {
          valor_credito: valorCredito,
          prazo_meses: prazoMeses,
          taxa_adm_pct: taxaAdmPct,
          fundo_reserva_pct: fundoReservaPct,
          seguro_vida_pct: seguroVidaPct,
          momento_contemplacao: momentoContemplacao,
          lance_embutido_pct: lanceEmbutidoPct,
          lance_livre_pct: lanceLivrePct,
          reajuste_pre_pct: correcaoAnual,
          reajuste_pos_pct: correcaoAnual,
          reajuste_pre_freq: 'Anual',
          reajuste_pos_freq: 'Anual',
        };
        fConsorcio = engine.calcularFluxoCompleto(params);
      }

      if (!fLance || !fLance.cashflow) {
        fLance = calcularCreditoLance({
          valor: valorLance,
          prazo: prazoLance,
          taxa: taxaLance,
          metodo: metodoLance,
        });
      }

      setFluxoConsorcio(fConsorcio);
      setFluxoLance(fLance);

      const combinado = calcularCustoCombinado(fConsorcio, fLance);
      setResult(combinado);
    } catch (err: any) {
      setError(err?.message ?? 'Erro ao calcular custo combinado');
    } finally {
      setLoading(false);
    }
  }

  // Build chart data
  const chartData = useMemo(() => {
    if (!result || !fluxoConsorcio || !fluxoLance) return [];
    const cfCons = fluxoConsorcio.cashflow ?? [];
    const cfLance = fluxoLance.cashflow ?? [];
    const cfComb = result.cashflowCombinado ?? [];

    const points: Array<{ mes: number; consorcio: number; lance: number; total: number }> = [];
    const maxLen = Math.max(cfCons.length, cfLance.length, cfComb.length);

    for (let i = 0; i < maxLen; i++) {
      const vC = i < cfCons.length ? Math.abs(cfCons[i]) : 0;
      const vL = i < cfLance.length ? Math.abs(cfLance[i]) : 0;
      const vT = i < cfComb.length ? Math.abs(cfComb[i]) : 0;
      // Skip month 0 (initial credit flow) and only include every nth month for readability
      if (i > 0 && (i % 3 === 0 || i <= 6)) {
        points.push({
          mes: i,
          consorcio: Math.round(vC),
          lance: Math.round(vL),
          total: Math.round(vT),
        });
      }
    }
    return points;
  }, [result, fluxoConsorcio, fluxoLance]);

  // Build table data
  const tableData = useMemo(() => {
    if (!result || !fluxoConsorcio || !fluxoLance) return [];
    const cfCons = fluxoConsorcio.cashflow ?? [];
    const cfLance = fluxoLance.cashflow ?? [];
    const cfComb = result.cashflowCombinado ?? [];
    const maxLen = Math.max(cfCons.length, cfLance.length, cfComb.length);

    const rows: Array<{ mes: number; consorcio: number; lance: number; total: number }> = [];
    for (let i = 0; i < maxLen; i++) {
      const vC = i < cfCons.length ? cfCons[i] : 0;
      const vL = i < cfLance.length ? cfLance[i] : 0;
      const vT = i < cfComb.length ? cfComb[i] : 0;
      if (Math.abs(vT) > 0.01) {
        rows.push({ mes: i, consorcio: vC, lance: vL, total: vT });
      }
    }
    return rows;
  }, [result, fluxoConsorcio, fluxoLance]);

  const totalConsorcio = fluxoConsorcio
    ? (fluxoConsorcio as any).total_pago ?? (fluxoConsorcio as any).totais?.total_pago ?? 0
    : 0;
  const totalLance = fluxoLance ? fluxoLance.total_pago ?? 0 : 0;

  return (
    <div className="flex flex-col h-full w-full overflow-hidden">
      <header className="sticky top-0 z-20 bg-somus-bg-primary/90 backdrop-blur-md border-b border-somus-border px-6 py-3">
        <div className="flex items-center gap-3">
          <Combine size={20} className="text-somus-gold" />
          <div>
            <h1 className="text-lg font-semibold text-somus-text-primary">Custo Combinado</h1>
            <p className="text-xs text-somus-text-tertiary">Consórcio + Op. Crédito Lance - espelha aba "Custo Consórcio + O.C. Lance" da NASA HD</p>
          </div>
        </div>
      </header>

      <main className="flex-1 overflow-y-auto bg-somus-bg-primary p-5 space-y-5">
        {/* Input forms */}
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-5">
          {/* Consórcio params */}
          <div className="bg-somus-bg-secondary border border-somus-border rounded-lg p-4">
            <h3 className="text-sm font-semibold text-somus-text-primary mb-3">Parâmetros do Consórcio</h3>
            <p className="text-[10px] text-somus-text-tertiary mb-3">Usa dados do Simulador se disponíveis, senão os valores abaixo.</p>
            <div className="grid grid-cols-3 gap-3">
              <DInput label="Valor Crédito">
                <input type="number" step={1000} value={valorCredito} onChange={(e) => setValorCredito(Number(e.target.value))} className={inputCls} />
              </DInput>
              <DInput label="Prazo (meses)">
                <input type="number" min={36} value={prazoMeses} onChange={(e) => setPrazoMeses(Number(e.target.value))} className={inputCls} />
              </DInput>
              <DInput label="Taxa Adm (%)">
                <input type="number" step={0.1} value={taxaAdmPct} onChange={(e) => setTaxaAdmPct(Number(e.target.value))} className={inputCls} />
              </DInput>
              <DInput label="Fdo Reserva (%)">
                <input type="number" step={0.1} value={fundoReservaPct} onChange={(e) => setFundoReservaPct(Number(e.target.value))} className={inputCls} />
              </DInput>
              <DInput label="Seguro (%)">
                <input type="number" step={0.01} value={seguroVidaPct} onChange={(e) => setSeguroVidaPct(Number(e.target.value))} className={inputCls} />
              </DInput>
              <DInput label="Contemplação (mês)">
                <input type="number" min={1} value={momentoContemplacao} onChange={(e) => setMomentoContemplacao(Number(e.target.value))} className={inputCls} />
              </DInput>
              <DInput label="Lance Embutido (%)">
                <input type="number" step={0.1} value={lanceEmbutidoPct} onChange={(e) => setLanceEmbutidoPct(Number(e.target.value))} className={inputCls} />
              </DInput>
              <DInput label="Lance Livre (%)">
                <input type="number" step={0.1} value={lanceLivrePct} onChange={(e) => setLanceLivrePct(Number(e.target.value))} className={inputCls} />
              </DInput>
              <DInput label="Correção Anual (%)">
                <input type="number" step={0.1} value={correcaoAnual} onChange={(e) => setCorrecaoAnual(Number(e.target.value))} className={inputCls} />
              </DInput>
            </div>
          </div>

          {/* Lance params */}
          <div className="bg-somus-bg-secondary border border-somus-border rounded-lg p-4">
            <h3 className="text-sm font-semibold text-somus-text-primary mb-3">Parâmetros do Crédito Lance</h3>
            <p className="text-[10px] text-somus-text-tertiary mb-3">Usa dados de Crédito p/ Lance se disponíveis, senão os valores abaixo.</p>
            <div className="grid grid-cols-2 gap-3">
              <DInput label="Valor do Empréstimo (R$)">
                <input type="number" step={1000} value={valorLance} onChange={(e) => setValorLance(Number(e.target.value))} className={inputCls} />
              </DInput>
              <DInput label="Prazo (meses)">
                <input type="number" min={1} value={prazoLance} onChange={(e) => setPrazoLance(Number(e.target.value))} className={inputCls} />
              </DInput>
              <DInput label="Taxa Mensal (%)">
                <input type="number" step={0.01} value={taxaLance} onChange={(e) => setTaxaLance(Number(e.target.value))} className={inputCls} />
              </DInput>
              <DInput label="Método">
                <select value={metodoLance} onChange={(e) => setMetodoLance(e.target.value as 'price' | 'sac')} className={inputCls}>
                  <option value="price">Price</option>
                  <option value="sac">SAC</option>
                </select>
              </DInput>
            </div>
          </div>
        </div>

        <button type="button" onClick={onCalculate} disabled={loading}
          className="w-full inline-flex items-center justify-center gap-2 px-4 py-2.5 text-sm font-semibold bg-somus-green text-white rounded-lg hover:bg-somus-green-light transition-colors disabled:opacity-50">
          <Calculator size={16} /> {loading ? 'Calculando...' : 'Calcular Custo Combinado'}
        </button>

        {error && (
          <div className="bg-red-500/10 border border-red-500/30 rounded-lg p-4 flex items-center gap-2">
            <AlertTriangle size={16} className="text-red-400 shrink-0" />
            <span className="text-xs text-red-400">{error}</span>
          </div>
        )}

        {result && (
          <div className="space-y-5">
            {/* Header: Combined TIR */}
            <div className="bg-somus-bg-secondary border border-somus-border rounded-lg p-5 text-center">
              <h3 className="text-xs text-somus-text-secondary uppercase mb-2">TIR Combinada</h3>
              <div className="flex items-center justify-center gap-8">
                <div>
                  <p className="text-2xl font-bold text-somus-gold">{fmtPct(result.tirMensal * 100, 4)}</p>
                  <span className="text-[10px] text-somus-text-tertiary">a.m.</span>
                </div>
                <div className="h-8 w-px bg-somus-border" />
                <div>
                  <p className="text-2xl font-bold text-somus-gold">{fmtPct(result.tirAnual * 100, 2)}</p>
                  <span className="text-[10px] text-somus-text-tertiary">a.a.</span>
                </div>
              </div>
            </div>

            {/* KPI Cards */}
            <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-5 gap-3">
              <div className="bg-somus-bg-secondary border border-somus-border rounded-lg p-3">
                <span className="text-[10px] text-somus-text-secondary uppercase">Total Consórcio</span>
                <p className="text-sm font-bold text-somus-green mt-1">{fmtBRL(totalConsorcio)}</p>
              </div>
              <div className="bg-somus-bg-secondary border border-somus-border rounded-lg p-3">
                <span className="text-[10px] text-somus-text-secondary uppercase">Total Lance</span>
                <p className="text-sm font-bold text-somus-purple mt-1">{fmtBRL(totalLance)}</p>
              </div>
              <div className="bg-somus-bg-secondary border border-somus-border rounded-lg p-3">
                <span className="text-[10px] text-somus-text-secondary uppercase">Total Combinado</span>
                <p className="text-sm font-bold text-somus-gold mt-1">{fmtBRL(result.totalPago)}</p>
              </div>
              <div className="bg-somus-bg-secondary border border-somus-border rounded-lg p-3">
                <span className="text-[10px] text-somus-text-secondary uppercase">TIR Combinado</span>
                <p className="text-sm font-bold text-somus-gold mt-1">{fmtPct(result.tirMensal * 100, 4)} a.m.</p>
              </div>
              <div className="bg-somus-bg-secondary border border-somus-border rounded-lg p-3">
                <span className="text-[10px] text-somus-text-secondary uppercase">CET Combinado</span>
                <p className="text-sm font-bold text-somus-gold mt-1">{fmtPct(result.tirAnual * 100, 2)} a.a.</p>
              </div>
            </div>

            {/* Area chart: stacked flows */}
            {chartData.length > 0 && (
              <div className="bg-somus-bg-secondary border border-somus-border rounded-lg p-5">
                <h3 className="text-sm font-semibold text-somus-text-primary mb-4">Fluxo de Caixa por Componente</h3>
                <ResponsiveContainer width="100%" height={300}>
                  <AreaChart data={chartData}>
                    <CartesianGrid strokeDasharray="3 3" stroke="#1E2A3A" vertical={false} />
                    <XAxis dataKey="mes" tick={{ fontSize: 10, fill: '#5A6577' }} tickLine={false} axisLine={{ stroke: '#1E2A3A' }} />
                    <YAxis tick={{ fontSize: 10, fill: '#5A6577' }} tickLine={false} axisLine={false} tickFormatter={(v: number) => `${(v / 1000).toFixed(0)}k`} />
                    <Tooltip content={<DarkTooltip />} />
                    <Legend iconType="circle" iconSize={8} wrapperStyle={{ fontSize: 11, color: '#8B95A5' }} />
                    <Area type="monotone" dataKey="consorcio" name="Consórcio" stroke="#1A7A3E" fill="#1A7A3E" fillOpacity={0.3} strokeWidth={2} stackId="1" />
                    <Area type="monotone" dataKey="lance" name="Crédito Lance" stroke="#7030A0" fill="#7030A0" fillOpacity={0.3} strokeWidth={2} stackId="1" />
                    <Area type="monotone" dataKey="total" name="Total" stroke="#D4A017" fill="none" fillOpacity={0} strokeWidth={2} strokeDasharray="5 5" />
                  </AreaChart>
                </ResponsiveContainer>
              </div>
            )}

            {/* Cash flow table */}
            <div className="bg-somus-bg-secondary border border-somus-border rounded-lg overflow-hidden">
              <div className="px-4 py-3 border-b border-somus-border flex items-center justify-between">
                <h3 className="text-sm font-semibold text-somus-text-primary">FLUXO DE CAIXA GERAL</h3>
                <span className="text-[10px] text-somus-text-tertiary">{tableData.length} registros</span>
              </div>
              <div className="overflow-x-auto max-h-[400px]">
                <table className="w-full text-xs">
                  <thead className="sticky top-0 bg-somus-bg-tertiary">
                    <tr>
                      <th className="px-3 py-2 text-left text-somus-text-secondary font-medium">Mês</th>
                      <th className="px-3 py-2 text-right font-medium" style={{ color: '#1A7A3E' }}>Op. Consórcio</th>
                      <th className="px-3 py-2 text-right font-medium" style={{ color: '#7030A0' }}>Op. Créd. Lance</th>
                      <th className="px-3 py-2 text-right font-medium" style={{ color: '#D4A017' }}>Total</th>
                    </tr>
                  </thead>
                  <tbody>
                    {tableData.map((r) => (
                      <tr key={r.mes} className="border-b border-somus-border/30">
                        <td className="px-3 py-1.5 text-somus-text-primary font-medium">{r.mes}</td>
                        <td className={`px-3 py-1.5 text-right ${r.consorcio >= 0 ? 'text-emerald-400' : ''}`} style={{ color: r.consorcio < 0 ? '#1A7A3E' : undefined }}>
                          {Math.abs(r.consorcio) > 0.01 ? fmtBRL(r.consorcio) : '—'}
                        </td>
                        <td className={`px-3 py-1.5 text-right`} style={{ color: Math.abs(r.lance) > 0.01 ? '#7030A0' : '#5A6577' }}>
                          {Math.abs(r.lance) > 0.01 ? fmtBRL(r.lance) : '—'}
                        </td>
                        <td className="px-3 py-1.5 text-right font-medium" style={{ color: '#D4A017' }}>
                          {fmtBRL(r.total)}
                        </td>
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

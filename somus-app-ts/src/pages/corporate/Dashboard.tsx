import React, { useState, useMemo } from 'react';
import {
  LayoutDashboard,
  Calculator,
  FileDown,
  FileText,
  Wallet,
  Receipt,
  Percent,
  PiggyBank,
  Target,
  TrendingUp,
  DollarSign,
  Activity,
} from 'lucide-react';
import {
  ResponsiveContainer,
  LineChart,
  Line,
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
import type { FluxoResult, VPLResult, FluxoMensal } from '@engine/nasa-engine';

// ── Format helpers ──────────────────────────────────────────────────────────

function fmtBRL(v: number): string {
  return v.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
}

function fmtPct(v: number, d = 2): string {
  return `${v.toLocaleString('pt-BR', { minimumFractionDigits: d, maximumFractionDigits: d })}%`;
}

// ── Dark Tooltip ────────────────────────────────────────────────────────────

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

// ── KPI Card (inline dark) ──────────────────────────────────────────────────

function DarkKPI({ label, value, sub, color = 'text-somus-text-primary', icon }: {
  label: string;
  value: string;
  sub?: string;
  color?: string;
  icon?: React.ReactNode;
}) {
  return (
    <div className="bg-somus-bg-secondary border border-somus-border rounded-lg p-4 flex flex-col gap-1">
      <div className="flex items-center justify-between">
        <span className="text-[11px] font-medium text-somus-text-secondary uppercase tracking-wider">{label}</span>
        {icon && <span className="text-somus-text-tertiary">{icon}</span>}
      </div>
      <span className={`text-lg font-bold ${color} tracking-tight`}>{value}</span>
      {sub && <span className="text-[10px] text-somus-text-tertiary">{sub}</span>}
    </div>
  );
}

// ── Main Component ──────────────────────────────────────────────────────────

export default function CorporateDashboard() {
  const setPage = useAppStore((s) => s.setPage);
  const [fluxoResult, setFluxoResult] = useState<FluxoResult | null>(null);
  const [vplResult, setVplResult] = useState<VPLResult | null>(null);

  // Auto-calculate with defaults on mount
  const engine = useMemo(() => new NasaEngine(), []);

  const defaultParams = useMemo(() => ({
    valor_credito: 500000,
    prazo_meses: 200,
    taxa_adm_pct: 20,
    fundo_reserva_pct: 3,
    momento_contemplacao: 36,
    seguro_vida_pct: 0.05,
    lance_embutido_pct: 10,
    lance_livre_pct: 20,
    reajuste_pre_pct: 7,
    reajuste_pos_pct: 7,
    reajuste_pre_freq: 'Anual',
    reajuste_pos_freq: 'Anual',
    alm_anual: 12,
    hurdle_anual: 12,
  }), []);

  useMemo(() => {
    try {
      const fr = engine.calcularFluxoCompleto(defaultParams);
      setFluxoResult(fr);
      const vr = engine.calcularVPLHD(defaultParams, fr);
      setVplResult(vr);
    } catch {
      // silent
    }
  }, [engine, defaultParams]);

  const fluxo = fluxoResult?.fluxo ?? [];
  const totais = fluxoResult?.totais;
  const metricas = fluxoResult?.metricas;
  const contemp = 36;

  // Chart: Carta de Crédito vs Saldo Devedor
  const creditVsDebtData = useMemo(() => {
    if (!fluxo.length) return [];
    return fluxo
      .filter((f: FluxoMensal) => f.mes > 0 && f.mes % 3 === 0)
      .map((f: FluxoMensal) => ({
        mes: f.mes,
        carta: Math.round(f.carta_credito_reajustada),
        saldo: Math.round(f.saldo_devedor_reajustado),
      }));
  }, [fluxo]);

  // Chart: VP Acumulado
  const vpAcumData = useMemo(() => {
    if (!vplResult) return [];
    const almM = monthlyFromAnnual(0.12);
    const hurdleM = monthlyFromAnnual(0.12);
    const points: Array<{ mes: number; vpPreT: number; vpPosT: number }> = [];
    let acumPre = 0;
    let acumPos = 0;

    for (const d of vplResult.pv_pre_t_detail ?? []) {
      acumPre += d.pv;
      if (d.mes % 3 === 0 || d.mes === contemp) {
        points.push({ mes: d.mes, vpPreT: Math.round(acumPre), vpPosT: 0 });
      }
    }
    for (const d of vplResult.pv_pos_t_detail ?? []) {
      acumPos += d.pv;
      if (d.mes % 6 === 0) {
        points.push({ mes: d.mes, vpPreT: Math.round(acumPre), vpPosT: Math.round(acumPos) });
      }
    }
    return points;
  }, [vplResult]);

  return (
    <div className="flex flex-col h-full w-full overflow-hidden">
      {/* Header */}
      <header className="sticky top-0 z-20 bg-somus-bg-primary/90 backdrop-blur-md border-b border-somus-border px-6 py-4">
        <div className="flex items-center justify-between">
          <div className="flex items-center gap-3">
            <LayoutDashboard size={22} className="text-somus-green" />
            <div>
              <h1 className="text-lg font-semibold text-somus-text-primary">Corporate Dashboard</h1>
              <p className="text-xs text-somus-text-tertiary">Visão geral do consórcio - NASA HD Engine</p>
            </div>
          </div>
          <div className="flex items-center gap-2">
            <button
              onClick={() => setPage('simulador')}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium bg-somus-green text-white rounded-lg hover:bg-somus-green-light transition-colors"
            >
              <Calculator size={14} /> Calcular
            </button>
            <button className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium bg-somus-bg-secondary border border-somus-border text-somus-text-secondary rounded-lg hover:bg-somus-bg-hover transition-colors">
              <FileDown size={14} /> Exportar Excel
            </button>
            <button
              onClick={() => setPage('gerador-propostas')}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium bg-somus-bg-secondary border border-somus-border text-somus-text-secondary rounded-lg hover:bg-somus-bg-hover transition-colors"
            >
              <FileText size={14} /> Gerar Proposta
            </button>
          </div>
        </div>
      </header>

      <main className="flex-1 overflow-y-auto bg-somus-bg-primary p-6 space-y-5">
        {/* ─── TOP ROW: 8 KPI Cards ───────────────────────────────────── */}
        <div className="grid grid-cols-2 sm:grid-cols-4 xl:grid-cols-8 gap-3">
          <DarkKPI
            label="Valor do Crédito"
            value={fmtBRL(500000)}
            icon={<Wallet size={14} />}
          />
          <DarkKPI
            label="Carta Líquida"
            value={totais ? fmtBRL(totais.carta_liquida) : '—'}
            color="text-somus-green"
            icon={<Wallet size={14} />}
          />
          <DarkKPI
            label="CET Anual"
            value={metricas ? fmtPct(metricas.cet_anual * 100, 2) : '—'}
            color="text-somus-gold"
            icon={<Percent size={14} />}
          />
          <DarkKPI
            label="TIR Anual"
            value={metricas ? fmtPct(metricas.tir_anual * 100, 2) : '—'}
            color="text-somus-gold"
            icon={<TrendingUp size={14} />}
          />
          <DarkKPI
            label="Delta VPL"
            value={vplResult ? fmtBRL(vplResult.delta_vpl) : '—'}
            color={vplResult && vplResult.delta_vpl >= 0 ? 'text-emerald-400' : 'text-red-400'}
            icon={<Activity size={14} />}
          />
          <DarkKPI
            label="Break-even Lance"
            value={vplResult ? fmtPct(vplResult.break_even_lance, 2) : '—'}
            color="text-somus-skyblue"
            icon={<Target size={14} />}
          />
          <DarkKPI
            label="Total Desembolsado"
            value={totais ? fmtBRL(totais.total_pago) : '—'}
            icon={<Receipt size={14} />}
          />
          <DarkKPI
            label="Custo Total %"
            value={metricas ? fmtPct(metricas.custo_total_pct, 2) : '—'}
            icon={<DollarSign size={14} />}
          />
        </div>

        {/* ─── MIDDLE ROW: 2 Charts ──────────────────────────────────── */}
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-5">
          {/* Carta de Crédito vs Saldo Devedor */}
          <div className="bg-somus-bg-secondary border border-somus-border rounded-lg p-5">
            <h3 className="text-sm font-semibold text-somus-text-primary mb-4">Carta de Crédito vs Saldo Devedor</h3>
            <ResponsiveContainer width="100%" height={260}>
              <LineChart data={creditVsDebtData}>
                <CartesianGrid strokeDasharray="3 3" stroke="#1E2A3A" vertical={false} />
                <XAxis dataKey="mes" tick={{ fontSize: 10, fill: '#5A6577' }} tickLine={false} axisLine={{ stroke: '#1E2A3A' }} />
                <YAxis tick={{ fontSize: 10, fill: '#5A6577' }} tickLine={false} axisLine={false} tickFormatter={(v: number) => `${(v / 1000).toFixed(0)}k`} />
                <Tooltip content={<DarkTooltip />} />
                <Legend iconType="circle" iconSize={8} wrapperStyle={{ fontSize: 11, color: '#8B95A5' }} />
                <ReferenceLine x={contemp} stroke="#D4A017" strokeDasharray="5 5" label={{ value: 'T', fill: '#D4A017', fontSize: 10 }} />
                <Line type="monotone" dataKey="carta" name="Carta de Crédito" stroke="#1A7A3E" strokeWidth={2} dot={false} />
                <Line type="monotone" dataKey="saldo" name="Saldo Devedor" stroke="#7030A0" strokeWidth={2} dot={false} />
              </LineChart>
            </ResponsiveContainer>
          </div>

          {/* VP Acumulado */}
          <div className="bg-somus-bg-secondary border border-somus-border rounded-lg p-5">
            <h3 className="text-sm font-semibold text-somus-text-primary mb-4">VP Acumulado</h3>
            <ResponsiveContainer width="100%" height={260}>
              <AreaChart data={vpAcumData}>
                <CartesianGrid strokeDasharray="3 3" stroke="#1E2A3A" vertical={false} />
                <XAxis dataKey="mes" tick={{ fontSize: 10, fill: '#5A6577' }} tickLine={false} axisLine={{ stroke: '#1E2A3A' }} />
                <YAxis tick={{ fontSize: 10, fill: '#5A6577' }} tickLine={false} axisLine={false} tickFormatter={(v: number) => `${(v / 1000).toFixed(0)}k`} />
                <Tooltip content={<DarkTooltip />} />
                <Legend iconType="circle" iconSize={8} wrapperStyle={{ fontSize: 11, color: '#8B95A5' }} />
                <ReferenceLine x={contemp} stroke="#D4A017" strokeDasharray="5 5" />
                <Area type="monotone" dataKey="vpPreT" name="VP Pré-T (ALM)" stroke="#1A7A3E" fill="#1A7A3E" fillOpacity={0.2} strokeWidth={2} />
                <Area type="monotone" dataKey="vpPosT" name="VP Pós-T (Hurdle)" stroke="#EF4444" fill="#EF4444" fillOpacity={0.15} strokeWidth={2} />
              </AreaChart>
            </ResponsiveContainer>
          </div>
        </div>

        {/* ─── BOTTOM ROW: 3 Mini-Tables ─────────────────────────────── */}
        <div className="grid grid-cols-1 lg:grid-cols-3 gap-5">
          {/* Composição de Custos */}
          <div className="bg-somus-bg-secondary border border-somus-border rounded-lg">
            <div className="px-4 py-3 border-b border-somus-border">
              <h3 className="text-sm font-semibold text-somus-text-primary">Composição de Custos</h3>
            </div>
            <table className="w-full text-xs">
              <tbody>
                {[
                  { label: 'Fundo Comum', value: totais?.total_fundo_comum ?? 0 },
                  { label: 'Taxa Administração', value: totais?.total_taxa_adm ?? 0 },
                  { label: 'Fundo Reserva', value: totais?.total_fundo_reserva ?? 0 },
                  { label: 'Seguro', value: totais?.total_seguro ?? 0 },
                  { label: 'Acessórios', value: totais?.total_custos_acessorios ?? 0 },
                ].map((r) => (
                  <tr key={r.label} className="border-b border-somus-border/50">
                    <td className="px-4 py-2 text-somus-text-secondary">{r.label}</td>
                    <td className="px-4 py-2 text-right font-medium text-somus-text-primary">{fmtBRL(r.value)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          {/* Dados dos Lances */}
          <div className="bg-somus-bg-secondary border border-somus-border rounded-lg">
            <div className="px-4 py-3 border-b border-somus-border">
              <h3 className="text-sm font-semibold text-somus-text-primary">Dados dos Lances</h3>
            </div>
            <table className="w-full text-xs">
              <tbody>
                {[
                  { label: 'Lance Embutido', value: totais?.lance_embutido_valor ?? 0 },
                  { label: 'Lance Livre', value: totais?.lance_livre_valor ?? 0 },
                  { label: 'Total Lances', value: (totais?.lance_embutido_valor ?? 0) + (totais?.lance_livre_valor ?? 0) },
                  { label: 'Crédito Líquido', value: totais?.carta_liquida ?? 0 },
                ].map((r) => (
                  <tr key={r.label} className="border-b border-somus-border/50">
                    <td className="px-4 py-2 text-somus-text-secondary">{r.label}</td>
                    <td className="px-4 py-2 text-right font-medium text-somus-text-primary">{fmtBRL(r.value)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          {/* Reajuste */}
          <div className="bg-somus-bg-secondary border border-somus-border rounded-lg">
            <div className="px-4 py-3 border-b border-somus-border">
              <h3 className="text-sm font-semibold text-somus-text-primary">Reajuste</h3>
            </div>
            <table className="w-full text-xs">
              <tbody>
                {[
                  { label: 'Pré-T Taxa', value: '7,00% a.a.' },
                  { label: 'Pré-T Frequência', value: 'Anual' },
                  { label: 'Pós-T Taxa', value: '7,00% a.a.' },
                  { label: 'Pós-T Frequência', value: 'Anual' },
                ].map((r) => (
                  <tr key={r.label} className="border-b border-somus-border/50">
                    <td className="px-4 py-2 text-somus-text-secondary">{r.label}</td>
                    <td className="px-4 py-2 text-right font-medium text-somus-text-primary">{r.value}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>

        {/* ─── NAVIGATION GRID ───────────────────────────────────────── */}
        <div className="grid grid-cols-2 sm:grid-cols-4 lg:grid-cols-5 gap-3">
          {[
            { key: 'simulador', label: 'Dados do Consórcio', icon: <Calculator size={16} />, color: 'text-somus-green' },
            { key: 'fluxo-financeiro', label: 'Fluxo Financeiro', icon: <FileText size={16} />, color: 'text-somus-purple' },
            { key: 'parcelas', label: 'Parcelas', icon: <Receipt size={16} />, color: 'text-somus-orange' },
            { key: 'comparativo-vpl', label: 'Comparativo VPL', icon: <TrendingUp size={16} />, color: 'text-somus-skyblue' },
            { key: 'consorcio-vs-financ', label: 'Cons. vs Financ.', icon: <Activity size={16} />, color: 'text-somus-navy' },
            { key: 'grafico', label: 'Gráfico', icon: <Activity size={16} />, color: 'text-somus-gold' },
            { key: 'resumo-cliente', label: 'Resumo Cliente', icon: <FileText size={16} />, color: 'text-somus-teal' },
            { key: 'cenarios', label: 'Cenários', icon: <PiggyBank size={16} />, color: 'text-somus-text-accent' },
            { key: 'gerador-propostas', label: 'Propostas', icon: <FileText size={16} />, color: 'text-somus-text-warning' },
            { key: 'fluxo-receitas', label: 'Fluxo Receitas', icon: <DollarSign size={16} />, color: 'text-somus-skyblue' },
          ].map((t) => (
            <button
              key={t.key}
              onClick={() => setPage(t.key)}
              className="flex items-center gap-2 bg-somus-bg-secondary border border-somus-border rounded-lg p-3 hover:bg-somus-bg-hover hover:border-somus-border-light transition-colors text-left"
            >
              <span className={t.color}>{t.icon}</span>
              <span className="text-xs font-medium text-somus-text-primary">{t.label}</span>
            </button>
          ))}
        </div>
      </main>
    </div>
  );
}

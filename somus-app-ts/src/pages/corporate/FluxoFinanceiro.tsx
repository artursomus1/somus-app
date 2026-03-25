import React, { useState, useMemo } from 'react';
import { Table, Download, Search, ChevronDown, ChevronRight } from 'lucide-react';
import { useAppStore } from '@/stores/appStore';
import { NasaEngine } from '@engine/index';
import type { FluxoResult, FluxoMensal } from '@engine/nasa-engine';

// ── Format helpers ──────────────────────────────────────────────────────────

function fmtBRL(v: number): string {
  return v.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL', minimumFractionDigits: 2 });
}

function fmtPct4(v: number): string {
  return `${(v * 100).toLocaleString('pt-BR', { minimumFractionDigits: 4, maximumFractionDigits: 4 })}%`;
}

function fmtPct2(v: number): string {
  return `${(v * 100).toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}%`;
}

// ── Column groups ───────────────────────────────────────────────────────────

interface ColGroup {
  label: string;
  color: string;
  keys: string[];
}

const COLUMN_GROUPS: ColGroup[] = [
  { label: 'Amortização FC', color: '#1A7A3E', keys: ['valor_base_fc', 'lance_embutido', 'lance_livre', 'valor_base_final', 'pct_mensal_fc', 'pct_acum_fc', 'amortizacao', 'saldo_principal'] },
  { label: 'Taxa de Administração', color: '#0EA5E9', keys: ['taxa_adm_antecipada', 'pct_ta_mensal', 'pct_ta_acum', 'valor_parcela_ta'] },
  { label: 'Fundo de Reserva', color: '#F59E0B', keys: ['pct_fr_mensal', 'pct_fr_acum', 'pct_fr_base', 'pct_fr_calc', 'fr_saldo', 'valor_fundo_reserva'] },
  { label: 'Reajuste', color: '#ED7D31', keys: ['valor_parcela', 'saldo_devedor', 'peso_parcela', 'pct_reajuste', 'pct_reajuste_acum', 'parcela_apos_reajuste', 'saldo_devedor_reajustado'] },
  { label: 'Seguro + Custos', color: '#EF4444', keys: ['seguro_vida', 'parcela_com_seguro', 'outros_custos'] },
  { label: 'Crédito + Fluxo', color: '#D4A017', keys: ['carta_credito_original', 'carta_credito_reajustada', 'fluxo_caixa', 'fluxo_caixa_tir'] },
];

const COL_LABELS: Record<string, string> = {
  valor_base_fc: 'Base FC',
  lance_embutido: 'L. Embutido',
  lance_livre: 'L. Livre',
  valor_base_final: 'Base Final',
  pct_mensal_fc: '% FC Mês',
  pct_acum_fc: '% FC Acum',
  amortizacao: 'Amortização',
  saldo_principal: 'Saldo Princ.',
  taxa_adm_antecipada: 'TA Antecip.',
  pct_ta_mensal: '% TA Mês',
  pct_ta_acum: '% TA Acum',
  valor_parcela_ta: 'Parcela TA',
  pct_fr_mensal: '% FR Mês',
  pct_fr_acum: '% FR Acum',
  pct_fr_base: '% FR Base',
  pct_fr_calc: '% FR Calc',
  fr_saldo: 'FR Saldo',
  valor_fundo_reserva: 'Valor FR',
  valor_parcela: 'Parcela Base',
  saldo_devedor: 'Saldo Dev.',
  peso_parcela: 'Peso Parc.',
  pct_reajuste: '% Reaj. Mês',
  pct_reajuste_acum: '% Reaj. Acum',
  parcela_apos_reajuste: 'Parc. Reaj.',
  saldo_devedor_reajustado: 'SD Reaj.',
  seguro_vida: 'Seguro',
  parcela_com_seguro: 'Parc. c/ Seg.',
  outros_custos: 'Outros',
  carta_credito_original: 'Carta Orig.',
  carta_credito_reajustada: 'Carta Reaj.',
  fluxo_caixa: 'Fluxo Cx',
  fluxo_caixa_tir: 'Fluxo TIR',
};

const PCT_COLS = new Set(['pct_mensal_fc', 'pct_acum_fc', 'pct_ta_mensal', 'pct_ta_acum', 'pct_fr_mensal', 'pct_fr_acum', 'pct_fr_base', 'pct_fr_calc', 'peso_parcela', 'pct_reajuste', 'pct_reajuste_acum']);

function formatCell(key: string, value: number): string {
  if (PCT_COLS.has(key)) return fmtPct4(value);
  return fmtBRL(value);
}

function cellColor(key: string, value: number): string {
  if (key === 'saldo_devedor') return '#7030A0';
  if (key === 'saldo_devedor_reajustado') return '#C00000';
  if (key === 'fluxo_caixa' || key === 'fluxo_caixa_tir') {
    if (value > 0) return '#D4A017';
    if (value < 0) return '#EF4444';
    return '#5A6577';
  }
  if (value < 0) return '#EF4444';
  return '';
}

// ── Main Component ──────────────────────────────────────────────────────────

export default function FluxoFinanceiro() {
  const setPage = useAppStore((s) => s.setPage);

  const engine = useMemo(() => new NasaEngine(), []);
  const fluxoResult = useMemo<FluxoResult>(() => {
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
  const [searchMin, setSearchMin] = useState(0);
  const [searchMax, setSearchMax] = useState(200);
  const [collapsedGroups, setCollapsedGroups] = useState<Set<number>>(new Set());
  const [compact, setCompact] = useState(true);
  const [visibleGroupFlags, setVisibleGroupFlags] = useState<boolean[]>(COLUMN_GROUPS.map(() => true));

  const toggleGroupVisibility = (idx: number) => {
    setVisibleGroupFlags((prev) => prev.map((v, i) => (i === idx ? !v : v)));
  };

  const toggleGroup = (idx: number) => {
    setCollapsedGroups((prev) => {
      const next = new Set(prev);
      if (next.has(idx)) next.delete(idx); else next.add(idx);
      return next;
    });
  };

  const visibleRows = useMemo(() => {
    return fluxoResult.fluxo.filter((f: FluxoMensal) => f.mes >= searchMin && f.mes <= searchMax);
  }, [fluxoResult, searchMin, searchMax]);

  // Totals row
  const totals = useMemo(() => {
    const t: Record<string, number> = {};
    const sumKeys = ['amortizacao', 'valor_parcela_ta', 'valor_fundo_reserva', 'valor_parcela', 'parcela_apos_reajuste', 'seguro_vida', 'parcela_com_seguro', 'outros_custos', 'fluxo_caixa', 'fluxo_caixa_tir'];
    for (const k of sumKeys) t[k] = 0;
    for (const f of fluxoResult.fluxo) {
      if (f.mes === 0) continue;
      for (const k of sumKeys) t[k] += (f as any)[k] ?? 0;
    }
    return t;
  }, [fluxoResult]);

  function exportCSV() {
    const allKeys = ['mes', 'meses_restantes', ...COLUMN_GROUPS.flatMap((g) => g.keys)];
    const header = allKeys.map((k) => COL_LABELS[k] ?? k).join(';');
    const rows = fluxoResult.fluxo.map((f: FluxoMensal) => allKeys.map((k) => {
      const v = (f as any)[k] ?? 0;
      return typeof v === 'number' ? v.toLocaleString('pt-BR') : v;
    }).join(';'));
    const csv = '\uFEFF' + [header, ...rows].join('\n');
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'fluxo_financeiro.csv';
    a.click();
    URL.revokeObjectURL(url);
  }

  const cellPad = compact ? 'px-1.5 py-0.5' : 'px-3 py-2';
  const fontSize = compact ? 'text-[10px]' : 'text-xs';

  return (
    <div className="flex flex-col h-full w-full overflow-hidden">
      <header className="sticky top-0 z-20 bg-somus-bg-primary/90 backdrop-blur-md border-b border-somus-border px-6 py-3">
        <div className="flex items-center justify-between">
          <div className="flex items-center gap-3">
            <Table size={20} className="text-somus-purple" />
            <div>
              <h1 className="text-lg font-semibold text-somus-text-primary">Fluxo Financeiro do Consórcio</h1>
              <p className="text-xs text-somus-text-tertiary">35 colunas - espelha aba "Fluxo Financeiro" da NASA HD</p>
            </div>
          </div>
          <div className="flex items-center gap-2">
            <button onClick={() => setCompact(!compact)} className="px-2.5 py-1 text-[10px] bg-somus-bg-secondary border border-somus-border text-somus-text-secondary rounded hover:bg-somus-bg-hover transition-colors">
              {compact ? 'Expandir' : 'Compactar'}
            </button>
            <button onClick={exportCSV} className="inline-flex items-center gap-1 px-2.5 py-1 text-[10px] bg-somus-bg-secondary border border-somus-border text-somus-text-secondary rounded hover:bg-somus-bg-hover transition-colors">
              <Download size={12} /> CSV
            </button>
          </div>
        </div>

        {/* Filters */}
        <div className="flex items-center gap-3 mt-2">
          <div className="flex items-center gap-1">
            <Search size={12} className="text-somus-text-tertiary" />
            <span className="text-[10px] text-somus-text-secondary">Mês:</span>
            <input type="number" min={0} max={200} value={searchMin} onChange={(e) => setSearchMin(Number(e.target.value))} className="w-14 px-1.5 py-0.5 text-[10px] bg-somus-bg-input border border-somus-border rounded text-somus-text-primary" />
            <span className="text-[10px] text-somus-text-tertiary">a</span>
            <input type="number" min={0} max={200} value={searchMax} onChange={(e) => setSearchMax(Number(e.target.value))} className="w-14 px-1.5 py-0.5 text-[10px] bg-somus-bg-input border border-somus-border rounded text-somus-text-primary" />
          </div>
          <span className="text-[10px] text-somus-text-tertiary">{visibleRows.length} linhas</span>
          <button onClick={() => { setSearchMin(0); setSearchMax(200); }} className="text-[10px] text-somus-text-tertiary hover:text-somus-text-primary">Limpar</button>
        </div>

        {/* Column group toggles */}
        <div className="flex flex-wrap items-center gap-3 mt-2">
          <span className="text-[10px] text-somus-text-secondary">Grupos:</span>
          {COLUMN_GROUPS.map((g, gi) => (
            <label key={gi} className="inline-flex items-center gap-1 cursor-pointer">
              <input
                type="checkbox"
                checked={visibleGroupFlags[gi]}
                onChange={() => toggleGroupVisibility(gi)}
                className="w-3 h-3 rounded border-somus-border text-somus-green focus:ring-somus-green/40"
              />
              <span className="text-[10px]" style={{ color: g.color }}>{g.label}</span>
            </label>
          ))}
        </div>
      </header>

      <main className="flex-1 overflow-auto bg-somus-bg-primary">
        <div className="overflow-x-auto">
          <table className={`${fontSize} border-collapse min-w-max`}>
            {/* Group Headers */}
            <thead className="sticky top-0 z-10">
              <tr className="bg-somus-bg-tertiary">
                <th className={`${cellPad} text-left text-somus-text-secondary font-medium sticky left-0 bg-somus-bg-tertiary z-20 border-r border-somus-border`} colSpan={2}>
                  Período
                </th>
                {COLUMN_GROUPS.map((g, gi) => {
                  if (!visibleGroupFlags[gi]) return null;
                  const collapsed = collapsedGroups.has(gi);
                  return (
                    <th
                      key={gi}
                      className={`${cellPad} text-center font-medium cursor-pointer border-r border-somus-border/50`}
                      style={{ color: g.color }}
                      colSpan={collapsed ? 1 : g.keys.length}
                      onClick={() => toggleGroup(gi)}
                    >
                      <span className="inline-flex items-center gap-1">
                        {collapsed ? <ChevronRight size={10} /> : <ChevronDown size={10} />}
                        {g.label}
                      </span>
                    </th>
                  );
                })}
              </tr>

              {/* Column Headers */}
              <tr className="bg-somus-bg-secondary border-b border-somus-border">
                <th className={`${cellPad} text-center text-somus-text-secondary font-medium sticky left-0 bg-somus-bg-secondary z-20 w-12`}>Mês</th>
                <th className={`${cellPad} text-center text-somus-text-secondary font-medium`}>Rest.</th>
                {COLUMN_GROUPS.map((g, gi) => {
                  if (!visibleGroupFlags[gi]) return null;
                  if (collapsedGroups.has(gi)) {
                    return <th key={gi} className={`${cellPad} text-center text-somus-text-tertiary font-normal`}>...</th>;
                  }
                  return g.keys.map((k) => (
                    <th key={k} className={`${cellPad} text-right text-somus-text-secondary font-medium whitespace-nowrap`}>
                      {COL_LABELS[k] ?? k}
                    </th>
                  ));
                })}
              </tr>

              {/* Totals Row */}
              <tr className="bg-somus-bg-input border-b-2 border-somus-gold/30">
                <td className={`${cellPad} text-center font-bold text-somus-gold sticky left-0 bg-somus-bg-input z-20`}>TOT</td>
                <td className={`${cellPad} text-center text-somus-text-tertiary`}>—</td>
                {COLUMN_GROUPS.map((g, gi) => {
                  if (!visibleGroupFlags[gi]) return null;
                  if (collapsedGroups.has(gi)) {
                    return <td key={gi} className={cellPad}>--</td>;
                  }
                  return g.keys.map((k) => {
                    const v = totals[k];
                    return (
                      <td key={k} className={`${cellPad} text-right font-semibold text-somus-text-primary`}>
                        {v !== undefined ? formatCell(k, v) : '--'}
                      </td>
                    );
                  });
                })}
              </tr>
            </thead>

            <tbody>
              {visibleRows.map((f: FluxoMensal) => {
                const isContemp = f.mes === contemp;
                const rowBg = isContemp ? 'bg-somus-gold/5' : '';
                return (
                  <tr key={f.mes} className={`border-b border-somus-border/20 hover:bg-somus-bg-hover/50 ${rowBg}`}>
                    <td className={`${cellPad} text-center font-medium sticky left-0 z-10 ${isContemp ? 'bg-somus-gold/10 text-somus-gold' : 'bg-somus-bg-primary text-somus-text-primary'}`}>
                      {f.mes}
                      {isContemp && <span className="ml-0.5 text-[8px]">T</span>}
                    </td>
                    <td className={`${cellPad} text-center text-somus-text-tertiary`}>{f.meses_restantes}</td>
                    {COLUMN_GROUPS.map((g, gi) => {
                      if (!visibleGroupFlags[gi]) return null;
                      if (collapsedGroups.has(gi)) {
                        return <td key={gi} className={`${cellPad} text-center text-somus-text-tertiary`}>...</td>;
                      }
                      return g.keys.map((k) => {
                        const v = (f as any)[k] ?? 0;
                        const cc = cellColor(k, v);
                        return (
                          <td key={k} className={`${cellPad} text-right whitespace-nowrap`} style={cc ? { color: cc } : undefined}>
                            {formatCell(k, v)}
                          </td>
                        );
                      });
                    })}
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </main>
    </div>
  );
}

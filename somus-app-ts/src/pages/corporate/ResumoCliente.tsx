import React, { useMemo } from 'react';
import { FileText } from 'lucide-react';
import { useAppStore } from '@/stores/appStore';
import { NasaEngine } from '@engine/index';
import type { FluxoResult, VPLResult, FluxoMensal } from '@engine/nasa-engine';

// ── Helpers ─────────────────────────────────────────────────────────────────

function fmtBRL(v: number): string {
  return v.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
}

function fmtPct(v: number, d = 2): string {
  return `${v.toLocaleString('pt-BR', { minimumFractionDigits: d, maximumFractionDigits: d })}%`;
}

// ── Panel component ─────────────────────────────────────────────────────────

function Panel({ title, number, children }: { title: string; number: number; children: React.ReactNode }) {
  return (
    <div className="bg-somus-bg-secondary border border-somus-border rounded-lg overflow-hidden">
      <div className="px-4 py-3 border-b border-somus-border flex items-center gap-2">
        <span className="inline-flex items-center justify-center h-5 w-5 rounded bg-somus-green/20 text-somus-green text-[10px] font-bold">{number}</span>
        <h3 className="text-sm font-semibold text-somus-text-primary">{title}</h3>
      </div>
      <div className="p-4">{children}</div>
    </div>
  );
}

function Row({ label, value, highlight, color }: { label: string; value: string; highlight?: boolean; color?: string }) {
  return (
    <div className={`flex items-center justify-between py-1.5 ${highlight ? 'bg-somus-gold/5 -mx-4 px-4' : ''}`}>
      <span className="text-xs text-somus-text-secondary">{label}</span>
      <span className={`text-xs font-semibold ${color || 'text-somus-text-primary'}`}>{value}</span>
    </div>
  );
}

// ── Main Component ──────────────────────────────────────────────────────────

export default function ResumoCliente() {
  const setPage = useAppStore((s) => s.setPage);

  const engine = useMemo(() => new NasaEngine(), []);

  const defaultParams = useMemo(() => ({
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
    alm_anual: 12,
    hurdle_anual: 12,
  }), []);

  const result = useMemo<FluxoResult>(() => engine.calcularFluxoCompleto(defaultParams), [engine, defaultParams]);
  const vplResult = useMemo<VPLResult>(() => engine.calcularVPLHD(defaultParams, result), [engine, defaultParams, result]);

  const credito = 500000;
  const prazo = 200;
  const contemp = 36;
  const taxaAdm = 20;
  const fundoReserva = 3;
  const seguro = 0.05;
  const lanceEmbPct = 10;
  const lanceLivrePct = 20;

  const fluxo = result.fluxo;
  const totais = result.totais;
  const metricas = result.metricas;

  const lanceEmbValor = credito * lanceEmbPct / 100;
  const lanceLivreValor = credito * lanceLivrePct / 100;
  const lanceTotalPct = lanceEmbPct + lanceLivrePct;

  // First payment composition
  const f1 = fluxo[1];
  const fc1 = f1 ? Math.abs(f1.amortizacao) : 0;
  const ta1 = f1 ? f1.valor_parcela_ta : 0;
  const fr1 = f1 ? f1.valor_fundo_reserva : 0;
  const sg1 = f1 ? f1.seguro_vida : 0;
  const total1 = f1 ? f1.parcela_com_seguro : 0;

  // Inspection moments
  const inspMeses = [1, 13, 25, 37];
  const inspData = inspMeses.map((m) => {
    const f = fluxo.find((r: FluxoMensal) => r.mes === m);
    return { mes: m, parcela: f?.parcela_com_seguro ?? 0 };
  });

  // Pre/Post parcela info
  const preContemp = fluxo.filter((f: FluxoMensal) => f.mes > 0 && f.mes <= contemp);
  const posContemp = fluxo.filter((f: FluxoMensal) => f.mes > contemp);
  const parcelaPre = preContemp.length > 0 ? preContemp[0].parcela_com_seguro : 0;
  const parcelaPos = posContemp.length > 0 ? posContemp[0].parcela_com_seguro : 0;

  return (
    <div className="flex flex-col h-full w-full overflow-hidden">
      <header className="sticky top-0 z-20 bg-somus-bg-primary/90 backdrop-blur-md border-b border-somus-border px-6 py-3">
        <div className="flex items-center gap-3">
          <FileText size={20} className="text-somus-teal" />
          <div>
            <h1 className="text-lg font-semibold text-somus-text-primary">Resumo Cliente</h1>
            <p className="text-xs text-somus-text-tertiary">Resumo executivo - espelha aba "Resumo Cliente" da NASA HD</p>
          </div>
        </div>
      </header>

      <main className="flex-1 overflow-y-auto bg-somus-bg-primary p-5">
        <div className="max-w-4xl mx-auto space-y-5">

          {/* Panel 1: Dados da Contratação */}
          <Panel title="Dados da Contratação" number={1}>
            <div className="grid grid-cols-2 lg:grid-cols-5 gap-4">
              {[
                { label: 'Valor do Bem', value: fmtBRL(credito) },
                { label: 'Prazo', value: `${prazo} meses` },
                { label: 'Taxa Adm', value: `${taxaAdm}% (${fmtBRL(credito * taxaAdm / 100)})` },
                { label: 'Fundo Reserva', value: `${fundoReserva}% (${fmtBRL(credito * fundoReserva / 100)})` },
                { label: 'Seguro', value: `${seguro}% (${fmtBRL(sg1)}/mês)` },
              ].map((d) => (
                <div key={d.label}>
                  <span className="text-[10px] text-somus-text-secondary uppercase">{d.label}</span>
                  <p className="text-sm font-semibold text-somus-text-primary mt-0.5">{d.value}</p>
                </div>
              ))}
            </div>
          </Panel>

          {/* Panel 2: Composição da Parcela Inicial */}
          <Panel title="Composição da Parcela Inicial" number={2}>
            <div className="flex items-center gap-2 mb-3">
              <span className="text-[10px] bg-somus-green/20 text-somus-green px-1.5 py-0.5 rounded font-bold">1º mês</span>
            </div>
            <div className="space-y-1">
              <Row label="Fundo Comum (FC)" value={fmtBRL(fc1)} />
              <Row label="Taxa Administração (TA)" value={fmtBRL(ta1)} />
              <Row label="Fundo Reserva (FR)" value={fmtBRL(fr1)} />
              <Row label="Seguro" value={fmtBRL(sg1)} />
              <div className="h-px bg-somus-border my-1" />
              <Row label="Total Parcela" value={fmtBRL(total1)} highlight color="text-somus-green" />
            </div>
          </Panel>

          {/* Panel 3: Contemplação e Lances */}
          <Panel title="Contemplação e Lances Projetados" number={3}>
            <div className="grid grid-cols-2 lg:grid-cols-4 gap-4">
              <div>
                <span className="text-[10px] text-somus-text-secondary uppercase">Contemplação</span>
                <p className="text-sm font-semibold text-somus-gold mt-0.5">Mês {contemp}</p>
              </div>
              <div>
                <span className="text-[10px] text-somus-text-secondary uppercase">Lance Embutido</span>
                <p className="text-sm font-semibold text-somus-text-primary mt-0.5">{fmtPct(lanceEmbPct)} ({fmtBRL(lanceEmbValor)})</p>
              </div>
              <div>
                <span className="text-[10px] text-somus-text-secondary uppercase">Lance Livre</span>
                <p className="text-sm font-semibold text-somus-text-primary mt-0.5">{fmtPct(lanceLivrePct)} ({fmtBRL(lanceLivreValor)})</p>
              </div>
              <div>
                <span className="text-[10px] text-somus-text-secondary uppercase">Total Lance</span>
                <p className="text-sm font-semibold text-somus-text-primary mt-0.5">{fmtPct(lanceTotalPct)} ({fmtBRL(lanceEmbValor + lanceLivreValor)})</p>
              </div>
            </div>
          </Panel>

          {/* Panel 4: Valor das Parcelas em Outros Momentos */}
          <Panel title="Valor das Parcelas em Outros Momentos" number={4}>
            <div className="overflow-x-auto">
              <table className="w-full text-xs">
                <thead>
                  <tr className="border-b border-somus-border">
                    <th className="px-3 py-2 text-left text-somus-text-secondary font-medium">Mês</th>
                    <th className="px-3 py-2 text-right text-somus-text-secondary font-medium">Valor da Parcela</th>
                    <th className="px-3 py-2 text-center text-somus-text-secondary font-medium">Fase</th>
                  </tr>
                </thead>
                <tbody>
                  {inspData.map((d) => (
                    <tr key={d.mes} className={`border-b border-somus-border/30 ${d.mes === contemp + 1 ? 'bg-somus-gold/5' : ''}`}>
                      <td className="px-3 py-2 text-somus-text-primary font-medium">{d.mes}</td>
                      <td className="px-3 py-2 text-right text-somus-text-primary font-semibold">{fmtBRL(d.parcela)}</td>
                      <td className="px-3 py-2 text-center">
                        <span className={`text-[10px] px-1.5 py-0.5 rounded ${d.mes <= contemp ? 'bg-somus-green/20 text-somus-green' : 'bg-somus-skyblue/20 text-somus-skyblue'}`}>
                          {d.mes <= contemp ? 'Pré-T' : 'Pós-T'}
                        </span>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </Panel>

          {/* Panel 5: Parcelas Pré/Pós + Reajuste */}
          <Panel title="Pré/Pós Contemplação + Reajuste" number={5}>
            <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
              <div className="space-y-1">
                <h4 className="text-xs font-semibold text-somus-text-primary mb-2">Parcelas</h4>
                <Row label="Parcela Pré-Contemplação (base)" value={fmtBRL(parcelaPre)} color="text-somus-green" />
                <Row label="Parcela Pós-Contemplação (base)" value={fmtBRL(parcelaPos)} color="text-somus-skyblue" />
                <Row label="Parcela Média" value={fmtBRL(metricas.parcela_media)} />
                <Row label="Parcela Máxima" value={fmtBRL(metricas.parcela_maxima)} />
                <Row label="Parcela Mínima" value={fmtBRL(metricas.parcela_minima)} />
                <div className="h-px bg-somus-border my-1" />
                <Row label="Total Desembolsado" value={fmtBRL(totais.total_pago)} highlight />
                <Row label="Carta Líquida" value={fmtBRL(totais.carta_liquida)} color="text-somus-green" />
                <Row label="Custo Total %" value={fmtPct(metricas.custo_total_pct)} />
              </div>

              <div className="space-y-1">
                <h4 className="text-xs font-semibold text-somus-text-primary mb-2">Reajuste</h4>
                <Row label="Pré-T Taxa" value="7,00% a.a." />
                <Row label="Pré-T Frequência" value="Anual" />
                <Row label="Pós-T Taxa" value="7,00% a.a." />
                <Row label="Pós-T Frequência" value="Anual" />
                <div className="h-px bg-somus-border my-1" />
                <h4 className="text-xs font-semibold text-somus-text-primary mt-3 mb-2">Métricas</h4>
                <Row label="TIR Mensal" value={fmtPct(metricas.tir_mensal * 100, 4)} color="text-somus-gold" />
                <Row label="CET Anual" value={fmtPct(metricas.cet_anual * 100, 2)} color="text-somus-gold" />
                <Row label="Delta VPL" value={fmtBRL(vplResult.delta_vpl)} color={vplResult.delta_vpl >= 0 ? 'text-emerald-400' : 'text-red-400'} />
                <Row label="Break-even Lance" value={fmtPct(vplResult.break_even_lance, 2)} color="text-somus-skyblue" />
              </div>
            </div>
          </Panel>

        </div>
      </main>
    </div>
  );
}

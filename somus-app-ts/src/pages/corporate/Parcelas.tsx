import React, { useMemo } from 'react';
import { List } from 'lucide-react';
import { useAppStore } from '@/stores/appStore';
import { NasaEngine } from '@engine/index';
import type { FluxoResult, FluxoMensal } from '@engine/nasa-engine';

// ── Helpers ─────────────────────────────────────────────────────────────────

function fmtBRL(v: number): string {
  return v.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
}

function fmtPct(v: number, d = 2): string {
  return `${v.toLocaleString('pt-BR', { minimumFractionDigits: d, maximumFractionDigits: d })}%`;
}

// ── Main Component ──────────────────────────────────────────────────────────

export default function Parcelas() {
  const setPage = useAppStore((s) => s.setPage);

  const engine = useMemo(() => new NasaEngine(), []);
  const result = useMemo<FluxoResult>(() => {
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
  const credito = 500000;
  const fluxo = result.fluxo.filter((f: FluxoMensal) => f.mes > 0);

  // Summary stats
  const parcelas = fluxo.map((f: FluxoMensal) => f.parcela_com_seguro);
  const primeira = parcelas[0] ?? 0;
  const ultima = parcelas[parcelas.length - 1] ?? 0;
  const media = parcelas.length > 0 ? parcelas.reduce((a: number, b: number) => a + b, 0) / parcelas.length : 0;
  const maxParcela = Math.max(...parcelas);
  const minParcela = Math.min(...parcelas.filter((p: number) => p > 0));

  // Multi-column layout (3 columns repeating: # | % | R$)
  const COLS_PER_GROUP = 3;
  const GROUPS = 4; // 4 groups side by side
  const rowsPerGroup = Math.ceil(fluxo.length / GROUPS);

  return (
    <div className="flex flex-col h-full w-full overflow-hidden">
      <header className="sticky top-0 z-20 bg-somus-bg-primary/90 backdrop-blur-md border-b border-somus-border px-6 py-3">
        <div className="flex items-center gap-3">
          <List size={20} className="text-somus-orange" />
          <div>
            <h1 className="text-lg font-semibold text-somus-text-primary">Parcelas</h1>
            <p className="text-xs text-somus-text-tertiary">Cronograma de pagamentos - espelha aba "Parcelas" da NASA HD</p>
          </div>
        </div>
      </header>

      <main className="flex-1 overflow-y-auto bg-somus-bg-primary p-5 space-y-5">
        {/* Summary Stats */}
        <div className="grid grid-cols-2 sm:grid-cols-5 gap-3">
          {[
            { label: 'Primeira Parcela', value: fmtBRL(primeira) },
            { label: 'Última Parcela', value: fmtBRL(ultima) },
            { label: 'Parcela Média', value: fmtBRL(media) },
            { label: 'Parcela Máxima', value: fmtBRL(maxParcela) },
            { label: 'Parcela Mínima', value: fmtBRL(minParcela) },
          ].map((s) => (
            <div key={s.label} className="bg-somus-bg-secondary border border-somus-border rounded-lg p-3">
              <span className="text-[10px] text-somus-text-secondary uppercase tracking-wider">{s.label}</span>
              <p className="text-sm font-bold text-somus-text-primary mt-1">{s.value}</p>
            </div>
          ))}
        </div>

        {/* Multi-column parcela table */}
        <div className="bg-somus-bg-secondary border border-somus-border rounded-lg overflow-hidden">
          <div className="overflow-x-auto">
            <table className="w-full text-[10px]">
              <thead>
                <tr className="bg-somus-bg-tertiary border-b border-somus-border">
                  {Array.from({ length: GROUPS }).map((_, gi) => (
                    <React.Fragment key={gi}>
                      <th className="px-2 py-2 text-center text-somus-text-secondary font-medium border-r border-somus-border/30">#</th>
                      <th className="px-2 py-2 text-right text-somus-text-secondary font-medium">% Crédito</th>
                      <th className={`px-2 py-2 text-right text-somus-text-secondary font-medium ${gi < GROUPS - 1 ? 'border-r-2 border-somus-border' : ''}`}>Valor (R$)</th>
                    </React.Fragment>
                  ))}
                </tr>
              </thead>
              <tbody>
                {Array.from({ length: rowsPerGroup }).map((_, ri) => (
                  <tr key={ri} className="border-b border-somus-border/20 hover:bg-somus-bg-hover/50">
                    {Array.from({ length: GROUPS }).map((_, gi) => {
                      const idx = gi * rowsPerGroup + ri;
                      const f = idx < fluxo.length ? fluxo[idx] : null;
                      if (!f) {
                        return (
                          <React.Fragment key={gi}>
                            <td className="px-2 py-1 border-r border-somus-border/30" />
                            <td className="px-2 py-1" />
                            <td className={`px-2 py-1 ${gi < GROUPS - 1 ? 'border-r-2 border-somus-border' : ''}`} />
                          </React.Fragment>
                        );
                      }

                      const mes = f.mes;
                      const isContemp = mes === contemp;
                      const isPreT = mes <= contemp;
                      const pct = credito > 0 ? (f.parcela_com_seguro / credito) * 100 : 0;

                      return (
                        <React.Fragment key={gi}>
                          <td className={`px-2 py-1 text-center font-medium border-r border-somus-border/30 ${isContemp ? 'bg-somus-gold/10 text-somus-gold font-bold' : 'text-somus-text-primary'}`}>
                            {mes}
                            {isContemp && <span className="text-[8px] ml-0.5">T</span>}
                          </td>
                          <td className={`px-2 py-1 text-right ${isContemp ? 'bg-somus-gold/10' : ''}`}>
                            <span className="text-somus-text-secondary">{fmtPct(pct, 4)}</span>
                          </td>
                          <td className={`px-2 py-1 text-right font-medium ${gi < GROUPS - 1 ? 'border-r-2 border-somus-border' : ''} ${isContemp ? 'bg-somus-gold/10 text-somus-gold' : ''}`}>
                            <span className={isPreT && !isContemp ? 'text-somus-green' : isContemp ? 'text-somus-gold' : 'text-somus-text-primary'}>
                              {fmtBRL(f.parcela_com_seguro)}
                            </span>
                          </td>
                        </React.Fragment>
                      );
                    })}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>

        {/* Phase indicator */}
        <div className="flex items-center gap-4 text-xs text-somus-text-secondary">
          <div className="flex items-center gap-1.5">
            <span className="h-3 w-3 rounded-sm bg-somus-green/30 border border-somus-green/50" />
            <span>Pré-contemplação (meses 1 a {contemp})</span>
          </div>
          <div className="flex items-center gap-1.5">
            <span className="h-3 w-3 rounded-sm bg-somus-gold/30 border border-somus-gold/50" />
            <span>Contemplação (mês {contemp})</span>
          </div>
          <div className="flex items-center gap-1.5">
            <span className="h-3 w-3 rounded-sm bg-somus-bg-tertiary border border-somus-border" />
            <span>Pós-contemplação (meses {contemp + 1} a 200)</span>
          </div>
        </div>

        {/* Correction factor */}
        <div className="bg-somus-bg-secondary border border-somus-border rounded-lg p-4">
          <h3 className="text-xs font-semibold text-somus-text-primary mb-2">Fator de Correção por Período</h3>
          <div className="overflow-x-auto max-h-[200px]">
            <table className="w-full text-[10px]">
              <thead className="sticky top-0 bg-somus-bg-tertiary">
                <tr>
                  <th className="px-2 py-1.5 text-left text-somus-text-secondary font-medium">Mês</th>
                  <th className="px-2 py-1.5 text-right text-somus-text-secondary font-medium">Fator Reajuste</th>
                  <th className="px-2 py-1.5 text-right text-somus-text-secondary font-medium">Parcela Base</th>
                  <th className="px-2 py-1.5 text-right text-somus-text-secondary font-medium">Parcela Reajustada</th>
                </tr>
              </thead>
              <tbody>
                {fluxo.filter((f: FluxoMensal) => f.mes % 12 === 0 || f.mes === 1 || f.mes === contemp).map((f: FluxoMensal) => (
                  <tr key={f.mes} className={`border-b border-somus-border/20 ${f.mes === contemp ? 'bg-somus-gold/5' : ''}`}>
                    <td className="px-2 py-1 text-somus-text-primary font-medium">{f.mes}</td>
                    <td className="px-2 py-1 text-right text-somus-text-secondary">{f.fator_reajuste.toLocaleString('pt-BR', { minimumFractionDigits: 6 })}</td>
                    <td className="px-2 py-1 text-right text-somus-text-primary">{fmtBRL(f.valor_parcela)}</td>
                    <td className="px-2 py-1 text-right text-somus-text-primary font-medium">{fmtBRL(f.parcela_apos_reajuste)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      </main>
    </div>
  );
}

import React, { useState, useCallback } from 'react';
import {
  Calculator,
  Layers,
  Plus,
  Trash2,
} from 'lucide-react';
import { consolidarCotas } from '@engine/index';
import type { GrupoInput, ConsolidacaoResult } from '@engine/consolidador-cotas';

// ── Format helpers ──────────────────────────────────────────────────────────

function fmtBRL(v: number): string {
  return v.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
}

function fmtPct(v: number, d = 2): string {
  return `${v.toLocaleString('pt-BR', { minimumFractionDigits: d, maximumFractionDigits: d })}%`;
}

const inputCls = 'w-full px-2.5 py-1.5 text-sm bg-somus-bg-input border border-somus-border rounded-md text-somus-text-primary focus:ring-1 focus:ring-somus-green/50 focus:border-somus-green outline-none';

interface GrupoRow {
  id: number;
  grupo: string;
  valor: number;
  prazo: number;
  taxaAdm: number;
  fundoReserva: number;
}

const emptyRow = (id: number): GrupoRow => ({
  id,
  grupo: `Grupo ${id}`,
  valor: 500000,
  prazo: 200,
  taxaAdm: 20,
  fundoReserva: 3,
});

// ── Main Component ──────────────────────────────────────────────────────────

export default function ConsolidacaoCotas() {
  const [nextId, setNextId] = useState(3);
  const [grupos, setGrupos] = useState<GrupoRow[]>([emptyRow(1), emptyRow(2)]);
  const [result, setResult] = useState<ConsolidacaoResult | null>(null);
  const [loading, setLoading] = useState(false);

  const addRow = useCallback(() => {
    setGrupos((prev) => [...prev, emptyRow(nextId)]);
    setNextId((prev) => prev + 1);
  }, [nextId]);

  const removeRow = useCallback((id: number) => {
    setGrupos((prev) => prev.filter((g) => g.id !== id));
  }, []);

  const updateRow = useCallback((id: number, field: keyof GrupoRow, value: any) => {
    setGrupos((prev) =>
      prev.map((g) => (g.id === id ? { ...g, [field]: value } : g))
    );
  }, []);

  function onCalculate() {
    setLoading(true);
    try {
      const inputs: GrupoInput[] = grupos.map((g) => ({
        grupo: g.grupo,
        valor: g.valor,
        prazo: g.prazo,
        taxaAdm: g.taxaAdm,
        fundoReserva: g.fundoReserva,
      }));
      const res = consolidarCotas(inputs);
      setResult(res);
    } finally {
      setLoading(false);
    }
  }

  // Calculate contribution weights
  const totalCredito = grupos.reduce((s, g) => s + (g.valor || 0), 0);

  return (
    <div className="flex flex-col h-full w-full overflow-hidden">
      <header className="sticky top-0 z-20 bg-somus-bg-primary/90 backdrop-blur-md border-b border-somus-border px-6 py-3">
        <div className="flex items-center gap-3">
          <Layers size={20} className="text-somus-skyblue" />
          <div>
            <h1 className="text-lg font-semibold text-somus-text-primary">Consolidação de Cotas</h1>
            <p className="text-xs text-somus-text-tertiary">Consolidando Cotas - espelha aba "Consolidando Cotas" da NASA HD</p>
          </div>
        </div>
      </header>

      <main className="flex-1 overflow-y-auto bg-somus-bg-primary p-5 space-y-5">
        {/* Input Table */}
        <div className="bg-somus-bg-secondary border border-somus-border rounded-lg overflow-hidden">
          <div className="px-4 py-3 border-b border-somus-border flex items-center justify-between">
            <h3 className="text-sm font-semibold text-somus-text-primary">Grupos de Consórcio</h3>
            <button type="button" onClick={addRow}
              className="inline-flex items-center gap-1 px-3 py-1.5 text-xs font-medium bg-somus-green/20 text-somus-green border border-somus-green/30 rounded-lg hover:bg-somus-green/30 transition-colors">
              <Plus size={12} /> Adicionar Grupo
            </button>
          </div>
          <div className="overflow-x-auto">
            <table className="w-full text-xs">
              <thead className="bg-somus-bg-tertiary">
                <tr>
                  <th className="px-3 py-2 text-left text-somus-text-secondary font-medium">Grupo</th>
                  <th className="px-3 py-2 text-right text-somus-text-secondary font-medium">Valor (R$)</th>
                  <th className="px-3 py-2 text-right text-somus-text-secondary font-medium">Prazo (meses)</th>
                  <th className="px-3 py-2 text-right text-somus-text-secondary font-medium">Tx Adm (%)</th>
                  <th className="px-3 py-2 text-right text-somus-text-secondary font-medium">Fdo Reserva (%)</th>
                  <th className="px-3 py-2 text-center text-somus-text-secondary font-medium w-16"></th>
                </tr>
              </thead>
              <tbody>
                {grupos.map((g) => (
                  <tr key={g.id} className="border-b border-somus-border/30">
                    <td className="px-3 py-1.5">
                      <input type="text" value={g.grupo} onChange={(e) => updateRow(g.id, 'grupo', e.target.value)} className={inputCls} />
                    </td>
                    <td className="px-3 py-1.5">
                      <input type="number" step={1000} value={g.valor} onChange={(e) => updateRow(g.id, 'valor', Number(e.target.value))} className={`${inputCls} text-right`} />
                    </td>
                    <td className="px-3 py-1.5">
                      <input type="number" min={1} value={g.prazo} onChange={(e) => updateRow(g.id, 'prazo', Number(e.target.value))} className={`${inputCls} text-right`} />
                    </td>
                    <td className="px-3 py-1.5">
                      <input type="number" step={0.1} value={g.taxaAdm} onChange={(e) => updateRow(g.id, 'taxaAdm', Number(e.target.value))} className={`${inputCls} text-right`} />
                    </td>
                    <td className="px-3 py-1.5">
                      <input type="number" step={0.1} value={g.fundoReserva} onChange={(e) => updateRow(g.id, 'fundoReserva', Number(e.target.value))} className={`${inputCls} text-right`} />
                    </td>
                    <td className="px-3 py-1.5 text-center">
                      {grupos.length > 1 && (
                        <button type="button" onClick={() => removeRow(g.id)}
                          className="p-1.5 rounded-md hover:bg-red-500/20 text-somus-text-tertiary hover:text-red-400 transition-colors">
                          <Trash2 size={14} />
                        </button>
                      )}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>

        {/* Calculate */}
        <button type="button" onClick={onCalculate} disabled={loading || grupos.length === 0}
          className="w-full inline-flex items-center justify-center gap-2 px-4 py-2.5 text-sm font-semibold bg-somus-green text-white rounded-lg hover:bg-somus-green-light transition-colors disabled:opacity-50">
          <Calculator size={16} /> {loading ? 'Calculando...' : 'Consolidar Cotas'}
        </button>

        {result && (
          <div className="space-y-5">
            {/* Two panels side by side */}
            <div className="grid grid-cols-1 lg:grid-cols-2 gap-5">
              {/* Left: Cálculo Original */}
              <div className="bg-somus-bg-secondary border border-somus-border rounded-lg p-5">
                <h3 className="text-sm font-semibold text-somus-text-primary mb-4">Cálculo Original (Soma Simples)</h3>
                <div className="space-y-3">
                  {[
                    { label: 'Total Crédito', value: fmtBRL(totalCredito), color: 'text-somus-green' },
                    { label: 'Prazo Médio Simples', value: `${(grupos.reduce((s, g) => s + g.prazo, 0) / grupos.length).toFixed(0)} meses`, color: '' },
                    { label: 'Tx Adm Média Simples', value: fmtPct(grupos.reduce((s, g) => s + g.taxaAdm, 0) / grupos.length), color: '' },
                    { label: 'Fdo Reserva Média Simples', value: fmtPct(grupos.reduce((s, g) => s + g.fundoReserva, 0) / grupos.length), color: '' },
                    { label: 'Quantidade de Grupos', value: `${grupos.length}`, color: '' },
                  ].map((r) => (
                    <div key={r.label} className="flex items-center justify-between py-2 border-b border-somus-border/30 text-xs">
                      <span className="text-somus-text-secondary">{r.label}</span>
                      <span className={`font-semibold ${r.color || 'text-somus-text-primary'}`}>{r.value}</span>
                    </div>
                  ))}
                </div>
              </div>

              {/* Right: Dados Consolidados e Ajustados */}
              <div className="bg-somus-bg-secondary border border-somus-border rounded-lg p-5">
                <h3 className="text-sm font-semibold text-somus-text-primary mb-4">Dados Consolidados e Ajustados</h3>
                <p className="text-[10px] text-somus-text-tertiary mb-3">Médias ponderadas pelo valor do crédito de cada grupo</p>
                <div className="space-y-3">
                  {[
                    { label: 'Total Crédito', value: fmtBRL(result.totalCredito), color: 'text-somus-green' },
                    { label: 'Prazo Médio Ponderado', value: `${result.prazoMedio.toFixed(1)} meses`, color: 'text-somus-skyblue' },
                    { label: 'Tx Adm Média Ponderada', value: fmtPct(result.taxaAdmMedia), color: 'text-somus-gold' },
                    { label: 'Fdo Reserva Média Ponderada', value: fmtPct(result.fundoReservaMedia), color: 'text-somus-gold' },
                    { label: 'Quantidade de Grupos', value: `${grupos.length}`, color: '' },
                  ].map((r) => (
                    <div key={r.label} className="flex items-center justify-between py-2 border-b border-somus-border/30 text-xs">
                      <span className="text-somus-text-secondary">{r.label}</span>
                      <span className={`font-semibold ${r.color || 'text-somus-text-primary'}`}>{r.value}</span>
                    </div>
                  ))}
                </div>
              </div>
            </div>

            {/* Detail table: contribution weight */}
            <div className="bg-somus-bg-secondary border border-somus-border rounded-lg overflow-hidden">
              <div className="px-4 py-3 border-b border-somus-border">
                <h3 className="text-sm font-semibold text-somus-text-primary">Peso de Cada Grupo na Consolidação</h3>
              </div>
              <div className="overflow-x-auto">
                <table className="w-full text-xs">
                  <thead className="bg-somus-bg-tertiary">
                    <tr>
                      <th className="px-3 py-2 text-left text-somus-text-secondary font-medium">Grupo</th>
                      <th className="px-3 py-2 text-right text-somus-text-secondary font-medium">Valor (R$)</th>
                      <th className="px-3 py-2 text-right text-somus-text-secondary font-medium">Peso (%)</th>
                      <th className="px-3 py-2 text-right text-somus-text-secondary font-medium">Prazo</th>
                      <th className="px-3 py-2 text-right text-somus-text-secondary font-medium">Tx Adm</th>
                      <th className="px-3 py-2 text-right text-somus-text-secondary font-medium">Fdo Reserva</th>
                    </tr>
                  </thead>
                  <tbody>
                    {grupos.map((g) => {
                      const peso = totalCredito > 0 ? (g.valor / totalCredito) * 100 : 0;
                      return (
                        <tr key={g.id} className="border-b border-somus-border/30">
                          <td className="px-3 py-2 text-somus-text-primary font-medium">{g.grupo}</td>
                          <td className="px-3 py-2 text-right text-somus-text-primary">{fmtBRL(g.valor)}</td>
                          <td className="px-3 py-2 text-right text-somus-skyblue font-semibold">{fmtPct(peso)}</td>
                          <td className="px-3 py-2 text-right text-somus-text-secondary">{g.prazo}</td>
                          <td className="px-3 py-2 text-right text-somus-text-secondary">{fmtPct(g.taxaAdm)}</td>
                          <td className="px-3 py-2 text-right text-somus-text-secondary">{fmtPct(g.fundoReserva)}</td>
                        </tr>
                      );
                    })}
                    <tr className="bg-somus-bg-tertiary font-semibold">
                      <td className="px-3 py-2 text-somus-text-primary">TOTAL</td>
                      <td className="px-3 py-2 text-right text-somus-green">{fmtBRL(totalCredito)}</td>
                      <td className="px-3 py-2 text-right text-somus-text-primary">100,00%</td>
                      <td className="px-3 py-2 text-right text-somus-skyblue">{result.prazoMedio.toFixed(1)}</td>
                      <td className="px-3 py-2 text-right text-somus-gold">{fmtPct(result.taxaAdmMedia)}</td>
                      <td className="px-3 py-2 text-right text-somus-gold">{fmtPct(result.fundoReservaMedia)}</td>
                    </tr>
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

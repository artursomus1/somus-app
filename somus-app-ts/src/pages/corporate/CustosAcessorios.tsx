import React, { useState, useCallback, useMemo, useEffect } from 'react';
import {
  Receipt,
  Plus,
  Trash2,
} from 'lucide-react';
import { saveData, loadData } from '@/services/storage';

// ── Format helpers ──────────────────────────────────────────────────────────

function fmtBRL(v: number): string {
  return v.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
}

const inputCls = 'w-full px-2.5 py-1.5 text-sm bg-somus-bg-input border border-somus-border rounded-md text-somus-text-primary focus:ring-1 focus:ring-somus-green/50 focus:border-somus-green outline-none';

const STORAGE_KEY = 'somus-custos-acessorios';

// ── Types ──────────────────────────────────────────────────────────────────

interface CustoItem {
  id: number;
  descricao: string;
  valor: number;
  momento: number;
  tipo: 'Fixo' | 'Percentual' | 'Recorrente';
}

// ── Main Component ──────────────────────────────────────────────────────────

export default function CustosAcessorios() {
  const [custos, setCustos] = useState<CustoItem[]>(() =>
    loadData<CustoItem[]>(STORAGE_KEY, [])
  );
  const [nextId, setNextId] = useState(() => {
    const loaded = loadData<CustoItem[]>(STORAGE_KEY, []);
    return loaded.length > 0 ? Math.max(...loaded.map((c) => c.id)) + 1 : 1;
  });

  // Form state for adding a new cost
  const [newDescricao, setNewDescricao] = useState('');
  const [newValor, setNewValor] = useState(0);
  const [newMomento, setNewMomento] = useState(0);
  const [newTipo, setNewTipo] = useState<'Fixo' | 'Percentual' | 'Recorrente'>('Fixo');

  // Persist on change
  useEffect(() => {
    saveData(STORAGE_KEY, custos);
  }, [custos]);

  const addCusto = useCallback(() => {
    if (!newDescricao.trim() || newValor <= 0) return;
    const item: CustoItem = {
      id: nextId,
      descricao: newDescricao.trim(),
      valor: newValor,
      momento: newMomento,
      tipo: newTipo,
    };
    setCustos((prev) => [...prev, item]);
    setNextId((prev) => prev + 1);
    setNewDescricao('');
    setNewValor(0);
    setNewMomento(0);
  }, [nextId, newDescricao, newValor, newMomento, newTipo]);

  const removeCusto = useCallback((id: number) => {
    setCustos((prev) => prev.filter((c) => c.id !== id));
  }, []);

  // Metrics
  const totalCustos = useMemo(() => custos.reduce((s, c) => s + c.valor, 0), [custos]);
  const quantidade = custos.length;
  const mediaPorMomento = useMemo(() => {
    if (custos.length === 0) return 0;
    const momentos = new Set(custos.map((c) => c.momento));
    return totalCustos / momentos.size;
  }, [custos, totalCustos]);

  // Summary pivot: grouped by momento
  const pivotData = useMemo(() => {
    const map = new Map<number, { total: number; items: CustoItem[] }>();
    for (const c of custos) {
      const existing = map.get(c.momento);
      if (existing) {
        existing.total += c.valor;
        existing.items.push(c);
      } else {
        map.set(c.momento, { total: c.valor, items: [c] });
      }
    }
    return Array.from(map.entries()).sort((a, b) => a[0] - b[0]);
  }, [custos]);

  return (
    <div className="flex flex-col h-full w-full overflow-hidden">
      <header className="sticky top-0 z-20 bg-somus-bg-primary/90 backdrop-blur-md border-b border-somus-border px-6 py-3">
        <div className="flex items-center gap-3">
          <Receipt size={20} className="text-somus-gold" />
          <div>
            <h1 className="text-lg font-semibold text-somus-text-primary">Custos Acessórios</h1>
            <p className="text-xs text-somus-text-tertiary">Gerenciamento de custos acessórios - espelha aba "Custos Acessórios" da NASA HD</p>
          </div>
        </div>
      </header>

      <main className="flex-1 overflow-y-auto bg-somus-bg-primary p-5 space-y-5">
        {/* KPI Cards */}
        <div className="grid grid-cols-3 gap-3">
          <div className="bg-somus-bg-secondary border border-somus-border rounded-lg p-4">
            <span className="text-[10px] text-somus-text-secondary uppercase">Total Custos</span>
            <p className="text-lg font-bold text-somus-text-primary mt-1">{fmtBRL(totalCustos)}</p>
          </div>
          <div className="bg-somus-bg-secondary border border-somus-border rounded-lg p-4">
            <span className="text-[10px] text-somus-text-secondary uppercase">Quantidade</span>
            <p className="text-lg font-bold text-somus-text-primary mt-1">{quantidade}</p>
          </div>
          <div className="bg-somus-bg-secondary border border-somus-border rounded-lg p-4">
            <span className="text-[10px] text-somus-text-secondary uppercase">Média por Momento</span>
            <p className="text-lg font-bold text-somus-gold mt-1">{fmtBRL(mediaPorMomento)}</p>
          </div>
        </div>

        {/* Add form (inline) */}
        <div className="bg-somus-bg-secondary border border-somus-border rounded-lg p-4">
          <h3 className="text-sm font-semibold text-somus-text-primary mb-3">Adicionar Custo</h3>
          <div className="grid grid-cols-2 sm:grid-cols-5 gap-3 items-end">
            <div>
              <label className="block text-[10px] font-medium text-somus-text-secondary uppercase tracking-wider mb-1">Descrição</label>
              <input type="text" value={newDescricao} onChange={(e) => setNewDescricao(e.target.value)}
                placeholder="Ex: TAC, Comissão..."
                className={inputCls} />
            </div>
            <div>
              <label className="block text-[10px] font-medium text-somus-text-secondary uppercase tracking-wider mb-1">Valor (R$)</label>
              <input type="number" step={100} value={newValor} onChange={(e) => setNewValor(Number(e.target.value))} className={inputCls} />
            </div>
            <div>
              <label className="block text-[10px] font-medium text-somus-text-secondary uppercase tracking-wider mb-1">Momento (mês)</label>
              <input type="number" min={0} value={newMomento} onChange={(e) => setNewMomento(Number(e.target.value))} className={inputCls} />
            </div>
            <div>
              <label className="block text-[10px] font-medium text-somus-text-secondary uppercase tracking-wider mb-1">Tipo</label>
              <select value={newTipo} onChange={(e) => setNewTipo(e.target.value as any)} className={inputCls}>
                <option value="Fixo">Fixo</option>
                <option value="Percentual">Percentual</option>
                <option value="Recorrente">Recorrente</option>
              </select>
            </div>
            <button type="button" onClick={addCusto}
              className="inline-flex items-center justify-center gap-1 px-4 py-1.5 text-sm font-semibold bg-somus-green text-white rounded-lg hover:bg-somus-green-light transition-colors disabled:opacity-50"
              disabled={!newDescricao.trim() || newValor <= 0}>
              <Plus size={14} /> Adicionar
            </button>
          </div>
        </div>

        {/* Costs table */}
        <div className="bg-somus-bg-secondary border border-somus-border rounded-lg overflow-hidden">
          <div className="px-4 py-3 border-b border-somus-border flex items-center justify-between">
            <h3 className="text-sm font-semibold text-somus-text-primary">Lista de Custos Acessórios</h3>
            <span className="text-[10px] text-somus-text-tertiary">{quantidade} itens</span>
          </div>
          <div className="overflow-x-auto max-h-[350px]">
            <table className="w-full text-xs">
              <thead className="sticky top-0 bg-somus-bg-tertiary">
                <tr>
                  <th className="px-3 py-2 text-left text-somus-text-secondary font-medium">Descrição</th>
                  <th className="px-3 py-2 text-right text-somus-text-secondary font-medium">Valor</th>
                  <th className="px-3 py-2 text-center text-somus-text-secondary font-medium">Momento (mês)</th>
                  <th className="px-3 py-2 text-center text-somus-text-secondary font-medium">Tipo</th>
                  <th className="px-3 py-2 text-center text-somus-text-secondary font-medium w-16"></th>
                </tr>
              </thead>
              <tbody>
                {custos.length === 0 && (
                  <tr>
                    <td colSpan={5} className="px-3 py-8 text-center text-somus-text-tertiary">
                      Nenhum custo acessório cadastrado. Use o formulário acima para adicionar.
                    </td>
                  </tr>
                )}
                {custos.map((c) => (
                  <tr key={c.id} className="border-b border-somus-border/30">
                    <td className="px-3 py-2 text-somus-text-primary font-medium">{c.descricao}</td>
                    <td className="px-3 py-2 text-right text-somus-text-primary">{fmtBRL(c.valor)}</td>
                    <td className="px-3 py-2 text-center text-somus-text-secondary">{c.momento}</td>
                    <td className="px-3 py-2 text-center">
                      <span className={`text-[10px] font-bold px-1.5 py-0.5 rounded ${
                        c.tipo === 'Fixo' ? 'bg-somus-green/20 text-somus-green' :
                        c.tipo === 'Percentual' ? 'bg-somus-gold/20 text-somus-gold' :
                        'bg-somus-purple/20 text-somus-purple'
                      }`}>
                        {c.tipo}
                      </span>
                    </td>
                    <td className="px-3 py-2 text-center">
                      <button type="button" onClick={() => removeCusto(c.id)}
                        className="p-1.5 rounded-md hover:bg-red-500/20 text-somus-text-tertiary hover:text-red-400 transition-colors">
                        <Trash2 size={14} />
                      </button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>

        {/* Summary pivot: by momento */}
        {pivotData.length > 0 && (
          <div className="bg-somus-bg-secondary border border-somus-border rounded-lg overflow-hidden">
            <div className="px-4 py-3 border-b border-somus-border">
              <h3 className="text-sm font-semibold text-somus-text-primary">Resumo por Momento</h3>
            </div>
            <div className="overflow-x-auto">
              <table className="w-full text-xs">
                <thead className="bg-somus-bg-tertiary">
                  <tr>
                    <th className="px-3 py-2 text-left text-somus-text-secondary font-medium">Momento (mês)</th>
                    <th className="px-3 py-2 text-right text-somus-text-secondary font-medium">Total</th>
                    <th className="px-3 py-2 text-right text-somus-text-secondary font-medium">Itens</th>
                    <th className="px-3 py-2 text-left text-somus-text-secondary font-medium">Descrições</th>
                  </tr>
                </thead>
                <tbody>
                  {pivotData.map(([momento, data]) => (
                    <tr key={momento} className="border-b border-somus-border/30">
                      <td className="px-3 py-2 text-somus-text-primary font-medium">{momento}</td>
                      <td className="px-3 py-2 text-right text-somus-gold font-semibold">{fmtBRL(data.total)}</td>
                      <td className="px-3 py-2 text-right text-somus-text-secondary">{data.items.length}</td>
                      <td className="px-3 py-2 text-somus-text-secondary">{data.items.map((i) => i.descricao).join(', ')}</td>
                    </tr>
                  ))}
                  <tr className="bg-somus-bg-tertiary font-semibold">
                    <td className="px-3 py-2 text-somus-text-primary">TOTAL</td>
                    <td className="px-3 py-2 text-right text-somus-green">{fmtBRL(totalCustos)}</td>
                    <td className="px-3 py-2 text-right text-somus-text-primary">{quantidade}</td>
                    <td className="px-3 py-2 text-somus-text-tertiary">{pivotData.length} momentos</td>
                  </tr>
                </tbody>
              </table>
            </div>
          </div>
        )}
      </main>
    </div>
  );
}

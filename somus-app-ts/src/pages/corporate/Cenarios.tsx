import React, { useState, useCallback } from 'react';
import {
  ArrowLeft,
  Layers,
  Plus,
  Trash2,
  Download,
  Upload,
  GitCompareArrows,
  Check,
  X,
} from 'lucide-react';
import { PageLayout } from '@components/PageLayout';
import { Card } from '@components/Card';
import { Button } from '@components/Button';
import { FormField } from '@components/FormField';
import { CurrencyInput } from '@components/CurrencyInput';
import { PercentInput } from '@components/PercentInput';
import { KPICard } from '@components/KPICard';
import { Modal } from '@components/Modal';
import { Select } from '@components/Select';
import { ChartCard, CHART_COLORS } from '@components/ChartCard';
import { useAppStore } from '@/stores/appStore';
import { usePersistedState } from '@/hooks/usePersistedState';
import { cn } from '@/utils/cn';
import { calcularFluxoConsorcio } from '@engine/index';

// ── Types ───────────────────────────────────────────────────────────────────

interface SavedScenario {
  id: number;
  nome: string;
  params: Record<string, any>;
  parcelaF1: number;
  parcelaF2: number;
  totalPago: number;
  cartaLiquida: number;
  tirMensal: number;
  cetAnual: number;
  lanceLivreValor: number;
  lanceEmbutidoValor: number;
  timestamp: string;
}

// ── Format helpers ──────────────────────────────────────────────────────────

function fmtBRL(v: number): string {
  return v.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
}

function fmtPct(v: number, d = 2): string {
  return `${v.toLocaleString('pt-BR', { minimumFractionDigits: d, maximumFractionDigits: d })}%`;
}

// ── Defaults ────────────────────────────────────────────────────────────────

const PRAZO_OPTIONS = [36, 48, 60, 72, 84, 96, 108, 120, 144, 156, 168, 180, 200, 216, 240, 360, 420]
  .map((p) => ({ value: String(p), label: `${p} meses` }));

const PARCELA_RED_OPTIONS = [
  { value: '100', label: '100%' },
  { value: '70', label: '70%' },
  { value: '50', label: '50%' },
];

// ── Save Modal ──────────────────────────────────────────────────────────────

function SaveModal({ open, onClose, onSave }: {
  open: boolean;
  onClose: () => void;
  onSave: (nome: string, params: Record<string, any>) => void;
}) {
  const [nome, setNome] = useState('');
  const [valorCarta, setValorCarta] = useState(500000);
  const [prazoMeses, setPrazoMeses] = useState(200);
  const [taxaAdm, setTaxaAdm] = useState(20);
  const [fundoReserva, setFundoReserva] = useState(3);
  const [seguro, setSeguro] = useState(0.05);
  const [correcaoAnual, setCorrecaoAnual] = useState(7);
  const [prazoContemp, setPrazoContemp] = useState(3);
  const [parcelaRedPct, setParcelaRedPct] = useState(70);
  const [lanceLivrePct, setLanceLivrePct] = useState(20);
  const [lanceEmbutidoPct, setLanceEmbutidoPct] = useState(10);

  const inputClass = 'w-full px-3 py-2 text-sm border border-somus-border rounded-lg focus:ring-2 focus:ring-somus-green/40 focus:border-somus-green outline-none bg-somus-bg-input text-somus-text-primary';

  return (
    <Modal
      open={open}
      onClose={onClose}
      title="Salvar Cenario"
      size="lg"
      footer={
        <>
          <Button variant="secondary" onClick={onClose}>Cancelar</Button>
          <Button
            variant="primary"
            disabled={!nome.trim()}
            onClick={() => {
              if (!nome.trim()) return;
              onSave(nome.trim(), {
                valor_carta: valorCarta,
                prazo_meses: prazoMeses,
                taxa_adm: taxaAdm,
                fundo_reserva: fundoReserva,
                seguro,
                correcao_anual: correcaoAnual,
                prazo_contemp: prazoContemp,
                parcela_red_pct: parcelaRedPct,
                lance_livre_pct: lanceLivrePct,
                lance_embutido_pct: lanceEmbutidoPct,
              });
            }}
          >
            Salvar
          </Button>
        </>
      }
    >
      <div className="space-y-4">
        <FormField label="Nome do Cenario" required>
          <input value={nome} onChange={(e) => setNome(e.target.value)} className={inputClass} placeholder="Ex: Cenario Base 500k" />
        </FormField>
        <FormField label="Valor da Carta (R$)">
          <CurrencyInput value={valorCarta} onChange={(v) => setValorCarta(v)} />
        </FormField>
        <div className="grid grid-cols-2 gap-3">
          <FormField label="Prazo (meses)">
            <Select options={PRAZO_OPTIONS} value={String(prazoMeses)} onChange={(e) => setPrazoMeses(Number(e.target.value))} />
          </FormField>
          <FormField label="Contemplacao (mes)">
            <input type="number" min={1} value={prazoContemp} onChange={(e) => setPrazoContemp(Number(e.target.value))} className={inputClass} />
          </FormField>
        </div>
        <div className="grid grid-cols-2 gap-3">
          <FormField label="Taxa Adm (%)">
            <PercentInput value={taxaAdm} onChange={(v) => setTaxaAdm(v)} />
          </FormField>
          <FormField label="Fdo Reserva (%)">
            <PercentInput value={fundoReserva} onChange={(v) => setFundoReserva(v)} />
          </FormField>
        </div>
        <div className="grid grid-cols-2 gap-3">
          <FormField label="Seguro (%)">
            <PercentInput value={seguro} onChange={(v) => setSeguro(v)} decimals={4} />
          </FormField>
          <FormField label="Correcao Anual (%)">
            <PercentInput value={correcaoAnual} onChange={(v) => setCorrecaoAnual(v)} />
          </FormField>
        </div>
        <div className="grid grid-cols-2 gap-3">
          <FormField label="Lance Livre (%)">
            <PercentInput value={lanceLivrePct} onChange={(v) => setLanceLivrePct(v)} />
          </FormField>
          <FormField label="Lance Embutido (%)">
            <PercentInput value={lanceEmbutidoPct} onChange={(v) => setLanceEmbutidoPct(v)} />
          </FormField>
        </div>
        <FormField label="Parcela Reduzida (%)">
          <Select options={PARCELA_RED_OPTIONS} value={String(parcelaRedPct)} onChange={(e) => setParcelaRedPct(Number(e.target.value))} />
        </FormField>
      </div>
    </Modal>
  );
}

// ── Scenario Card ───────────────────────────────────────────────────────────

function ScenarioCard({ scenario, selected, onToggleSelect, onDelete }: {
  scenario: SavedScenario;
  selected: boolean;
  onToggleSelect: () => void;
  onDelete: () => void;
}) {
  return (
    <div
      className={cn(
        'rounded-lg border-2 p-5 transition-colors cursor-pointer',
        selected ? 'border-somus-green bg-somus-green-bg' : 'border-somus-gray-200 bg-white hover:border-somus-gray-300',
      )}
      onClick={onToggleSelect}
    >
      <div className="flex items-start justify-between mb-3">
        <div className="flex items-center gap-2">
          <div className={cn(
            'h-5 w-5 rounded-md border-2 flex items-center justify-center',
            selected ? 'bg-somus-green border-somus-green' : 'border-somus-gray-300',
          )}>
            {selected && <Check className="h-3.5 w-3.5 text-white" />}
          </div>
          <h3 className="font-semibold text-somus-gray-900">{scenario.nome}</h3>
        </div>
        <button onClick={(e) => { e.stopPropagation(); onDelete(); }} className="p-1 hover:bg-red-50 rounded-md">
          <Trash2 className="h-4 w-4 text-red-400" />
        </button>
      </div>
      <div className="grid grid-cols-2 gap-x-4 gap-y-2 text-sm">
        <div><span className="text-somus-gray-500">Carta Liq.:</span> <span className="ml-1 font-medium">{fmtBRL(scenario.cartaLiquida)}</span></div>
        <div><span className="text-somus-gray-500">Total Pago:</span> <span className="ml-1 font-medium">{fmtBRL(scenario.totalPago)}</span></div>
        <div><span className="text-somus-gray-500">CET a.a.:</span> <span className="ml-1 font-medium">{fmtPct(scenario.cetAnual * 100, 2)}</span></div>
        <div><span className="text-somus-gray-500">Parcela F1:</span> <span className="ml-1 font-medium">{fmtBRL(scenario.parcelaF1)}</span></div>
        <div><span className="text-somus-gray-500">Parcela F2:</span> <span className="ml-1 font-medium">{fmtBRL(scenario.parcelaF2)}</span></div>
      </div>
      <p className="text-xs text-somus-gray-400 mt-3">
        Salvo em {new Date(scenario.timestamp).toLocaleDateString('pt-BR')}
      </p>
    </div>
  );
}

// ── Main Component ──────────────────────────────────────────────────────────

export default function Cenarios() {
  const setPage = useAppStore((s) => s.setPage);
  const [cenarios, setCenarios] = usePersistedState<SavedScenario[]>('cenarios', []);
  const [selectedIds, setSelectedIds] = useState<Set<number>>(new Set());
  const [modalOpen, setModalOpen] = useState(false);
  const [comparing, setComparing] = useState(false);

  const persist = useCallback((list: SavedScenario[]) => {
    setCenarios(list);
  }, [setCenarios]);

  function handleSave(nome: string, params: Record<string, any>) {
    if (cenarios.length >= 10) {
      alert('Maximo de 10 cenarios atingido.');
      return;
    }

    const res = calcularFluxoConsorcio(params) as Record<string, any>;
    const metricas = res.metricas ?? {};

    const newScenario: SavedScenario = {
      id: Date.now(),
      nome,
      params,
      parcelaF1: res.parcela_f1_base ?? 0,
      parcelaF2: res.parcela_f2_base ?? 0,
      totalPago: res.total_pago ?? 0,
      cartaLiquida: res.carta_liquida ?? 0,
      tirMensal: metricas.tir_mensal ?? 0,
      cetAnual: metricas.cet_anual ?? 0,
      lanceLivreValor: res.lance_livre_valor ?? 0,
      lanceEmbutidoValor: res.lance_embutido_valor ?? 0,
      timestamp: new Date().toISOString(),
    };

    persist([...cenarios, newScenario]);
    setModalOpen(false);
  }

  function handleDelete(id: number) {
    persist(cenarios.filter((c) => c.id !== id));
    setSelectedIds((prev) => { const n = new Set(prev); n.delete(id); return n; });
  }

  function toggleSelect(id: number) {
    setSelectedIds((prev) => {
      const n = new Set(prev);
      if (n.has(id)) n.delete(id); else if (n.size < 3) n.add(id);
      return n;
    });
  }

  function handleExport() {
    const blob = new Blob([JSON.stringify(cenarios, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `cenarios_somus_${new Date().toISOString().split('T')[0]}.json`;
    a.click();
    URL.revokeObjectURL(url);
  }

  function handleImport() {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = '.json';
    input.onchange = (e) => {
      const file = (e.target as HTMLInputElement).files?.[0];
      if (!file) return;
      const reader = new FileReader();
      reader.onload = () => {
        try {
          const imported = JSON.parse(reader.result as string) as SavedScenario[];
          if (Array.isArray(imported)) persist([...cenarios, ...imported].slice(0, 10));
        } catch { alert('Arquivo JSON invalido.'); }
      };
      reader.readAsText(file);
    };
    input.click();
  }

  const selected = cenarios.filter((c) => selectedIds.has(c.id));

  const compChartData = selected.map((s) => ({
    nome: s.nome.length > 12 ? s.nome.slice(0, 12) + '...' : s.nome,
    totalPago: s.totalPago,
    cartaLiquida: s.cartaLiquida,
    parcelaF1: s.parcelaF1,
    parcelaF2: s.parcelaF2,
  }));

  return (
    <PageLayout title="Cenarios" subtitle="Gerencie ate 10 cenarios salvos">
      <div className="max-w-7xl mx-auto">
        <div className="flex items-center justify-between mb-6">
          <button onClick={() => setPage('dashboard')} className="inline-flex items-center gap-1.5 text-sm text-somus-gray-500 hover:text-somus-gray-700 transition-colors">
            <ArrowLeft className="h-4 w-4" /> Voltar ao Dashboard
          </button>
          <div className="flex items-center gap-2">
            <Button variant="secondary" size="sm" icon={<Upload className="h-4 w-4" />} onClick={handleImport}>Importar</Button>
            <Button variant="secondary" size="sm" icon={<Download className="h-4 w-4" />} onClick={handleExport} disabled={cenarios.length === 0}>Exportar</Button>
            <Button variant="primary" size="sm" icon={<Plus className="h-4 w-4" />} onClick={() => setModalOpen(true)} disabled={cenarios.length >= 10}>Novo Cenario</Button>
          </div>
        </div>

        {/* Info bar */}
        <div className="flex items-center justify-between mb-5">
          <p className="text-sm text-somus-gray-500">
            {cenarios.length}/10 cenarios
            {selectedIds.size > 0 && ` | ${selectedIds.size} selecionado(s)`}
          </p>
          {selectedIds.size >= 2 && (
            <Button variant="primary" size="sm" icon={<GitCompareArrows className="h-4 w-4" />} onClick={() => setComparing(!comparing)}>
              {comparing ? 'Fechar Comparacao' : 'Comparar Selecionados'}
            </Button>
          )}
        </div>

        {/* Scenario Grid */}
        {cenarios.length === 0 ? (
          <div className="flex flex-col items-center justify-center py-20 bg-white rounded-lg border-2 border-dashed border-somus-gray-200">
            <Layers className="h-12 w-12 text-somus-gray-300 mb-4" />
            <p className="text-sm text-somus-gray-400 mb-4">Nenhum cenario salvo</p>
            <Button variant="primary" icon={<Plus className="h-4 w-4" />} onClick={() => setModalOpen(true)}>Criar Primeiro Cenario</Button>
          </div>
        ) : (
          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4 mb-8">
            {cenarios.map((s) => (
              <ScenarioCard key={s.id} scenario={s} selected={selectedIds.has(s.id)} onToggleSelect={() => toggleSelect(s.id)} onDelete={() => handleDelete(s.id)} />
            ))}
          </div>
        )}

        {/* ─── Comparison View ─────────────────────────────────────── */}
        {comparing && selected.length >= 2 && (
          <div className="space-y-6">
            <h2 className="text-lg font-bold text-somus-gray-900">Comparacao de Cenarios</h2>

            {/* Comparison Table */}
            <Card title="Metricas Comparativas" padding="none">
              <div className="overflow-x-auto">
                <table className="w-full text-sm">
                  <thead>
                    <tr className="bg-somus-gray-50 border-b border-somus-gray-200">
                      <th className="text-left px-5 py-3 font-medium text-somus-gray-600">Metrica</th>
                      {selected.map((s) => (
                        <th key={s.id} className="text-right px-5 py-3 font-medium text-somus-green">{s.nome}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {[
                      { label: 'Carta Liquida', fn: (s: SavedScenario) => fmtBRL(s.cartaLiquida) },
                      { label: 'Total Pago', fn: (s: SavedScenario) => fmtBRL(s.totalPago) },
                      { label: 'Parcela F1', fn: (s: SavedScenario) => fmtBRL(s.parcelaF1) },
                      { label: 'Parcela F2', fn: (s: SavedScenario) => fmtBRL(s.parcelaF2) },
                      { label: 'CET Mensal', fn: (s: SavedScenario) => fmtPct(s.tirMensal * 100, 4) },
                      { label: 'CET Anual', fn: (s: SavedScenario) => fmtPct(s.cetAnual * 100, 2) },
                      { label: 'Lance Livre', fn: (s: SavedScenario) => fmtBRL(s.lanceLivreValor) },
                      { label: 'Lance Embutido', fn: (s: SavedScenario) => fmtBRL(s.lanceEmbutidoValor) },
                    ].map((row) => (
                      <tr key={row.label} className="border-b border-somus-gray-100 hover:bg-somus-gray-50">
                        <td className="px-5 py-3 text-somus-gray-700">{row.label}</td>
                        {selected.map((s) => (
                          <td key={s.id} className="text-right px-5 py-3 font-medium">{row.fn(s)}</td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </Card>

            {/* Charts */}
            <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
              <ChartCard
                title="Total Pago vs Carta Liquida"
                type="bar"
                data={compChartData}
                series={[
                  { dataKey: 'totalPago', name: 'Total Pago', color: '#EF4444' },
                  { dataKey: 'cartaLiquida', name: 'Carta Liquida', color: '#059669' },
                ]}
                xAxisKey="nome"
                height={260}
                valueFormatter={(v) => fmtBRL(v)}
              />
              <ChartCard
                title="Parcelas F1 vs F2"
                type="bar"
                data={compChartData}
                series={[
                  { dataKey: 'parcelaF1', name: 'Parcela F1', color: '#2563EB' },
                  { dataKey: 'parcelaF2', name: 'Parcela F2', color: '#7C3AED' },
                ]}
                xAxisKey="nome"
                height={260}
                valueFormatter={(v) => fmtBRL(v)}
              />
            </div>
          </div>
        )}
      </div>

      <SaveModal open={modalOpen} onClose={() => setModalOpen(false)} onSave={handleSave} />
    </PageLayout>
  );
}

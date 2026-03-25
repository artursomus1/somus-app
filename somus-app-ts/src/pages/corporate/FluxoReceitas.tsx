import React, { useState, useEffect, useCallback } from 'react';
import {
  ArrowLeft,
  Plus,
  Search,
  ChevronDown,
  ChevronRight,
  Edit3,
  Trash2,
  DollarSign,
} from 'lucide-react';
import { PageLayout } from '@components/PageLayout';
import { Card } from '@components/Card';
import { Button } from '@components/Button';
import { FormField } from '@components/FormField';
import { CurrencyInput } from '@components/CurrencyInput';
import { StatusBadge } from '@components/StatusBadge';
import { KPICard } from '@components/KPICard';
import { Modal } from '@components/Modal';
import { Select } from '@components/Select';
import { useAppStore } from '@/stores/appStore';
import { usePersistedState } from '@/hooks/usePersistedState';
import { cn } from '@/utils/cn';
import type { Operacao, PagamentoOperacao } from '@/types';

function generateId(): string {
  return `op_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;
}

function fmtBRL(v: number): string {
  return v.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
}

// ── Status mapping ──────────────────────────────────────────────────────────

function mapOpStatus(status: Operacao['status']): 'ativa' | 'pendente' | 'encerrada' {
  return status === 'contemplada' ? 'pendente' : status;
}

function mapPgStatus(status: PagamentoOperacao['status']): 'ativa' | 'pendente' | 'atrasado' {
  return status === 'pago' ? 'ativa' : status;
}

// ── Add/Edit Modal ──────────────────────────────────────────────────────────

const TIPOS_OPERACAO = ['Imovel', 'Automovel', 'Servico', 'Caminhao', 'Maquina Agricola', 'Outro']
  .map((t) => ({ value: t, label: t }));

const STATUS_OPTIONS = [
  { value: 'ativa', label: 'Ativa' },
  { value: 'contemplada', label: 'Contemplada' },
  { value: 'encerrada', label: 'Encerrada' },
];

function OperacaoModal({
  open,
  onClose,
  onSave,
  editingOp,
}: {
  open: boolean;
  onClose: () => void;
  onSave: (op: Operacao) => void;
  editingOp?: Operacao | null;
}) {
  const [clienteNome, setClienteNome] = useState('');
  const [assessor, setAssessor] = useState('');
  const [tipo, setTipo] = useState('Imovel');
  const [valorCarta, setValorCarta] = useState(0);
  const [status, setStatus] = useState<Operacao['status']>('ativa');
  const [dataInicio, setDataInicio] = useState('');
  const [numParcelas, setNumParcelas] = useState(12);
  const [valorParcela, setValorParcela] = useState(0);

  useEffect(() => {
    if (editingOp) {
      setClienteNome(editingOp.clienteNome);
      setAssessor(editingOp.assessor);
      setTipo(editingOp.tipo);
      setValorCarta(editingOp.valorCarta);
      setStatus(editingOp.status);
      setDataInicio(editingOp.dataInicio);
    } else {
      setClienteNome('');
      setAssessor('');
      setTipo('Imovel');
      setValorCarta(0);
      setStatus('ativa');
      setDataInicio(new Date().toISOString().split('T')[0]);
      setNumParcelas(12);
      setValorParcela(0);
    }
  }, [editingOp, open]);

  const inputClass = 'w-full px-3 py-2 text-sm border border-somus-border rounded-lg focus:ring-2 focus:ring-somus-green/40 focus:border-somus-green outline-none bg-somus-bg-input text-somus-text-primary';

  function handleSave() {
    const pagamentos: PagamentoOperacao[] = editingOp
      ? editingOp.pagamentos
      : Array.from({ length: numParcelas }, (_, i) => ({
          mes: i + 1,
          valor: valorParcela,
          data: '',
          status: 'pendente' as const,
        }));

    onSave({
      id: editingOp?.id || generateId(),
      clienteNome, assessor, tipo, valorCarta, status, dataInicio, pagamentos,
    });
  }

  return (
    <Modal
      open={open}
      onClose={onClose}
      title={editingOp ? 'Editar Operacao' : 'Nova Operacao'}
      footer={
        <>
          <Button variant="secondary" onClick={onClose}>Cancelar</Button>
          <Button variant="primary" onClick={handleSave}>{editingOp ? 'Salvar' : 'Criar Operacao'}</Button>
        </>
      }
    >
      <div className="space-y-4">
        <FormField label="Nome do Cliente" required>
          <input value={clienteNome} onChange={(e) => setClienteNome(e.target.value)} className={inputClass} />
        </FormField>
        <FormField label="Assessor" required>
          <input value={assessor} onChange={(e) => setAssessor(e.target.value)} className={inputClass} />
        </FormField>
        <div className="grid grid-cols-2 gap-4">
          <FormField label="Tipo">
            <Select options={TIPOS_OPERACAO} value={tipo} onChange={(e) => setTipo(e.target.value)} />
          </FormField>
          <FormField label="Status">
            <Select options={STATUS_OPTIONS} value={status} onChange={(e) => setStatus(e.target.value as Operacao['status'])} />
          </FormField>
        </div>
        <FormField label="Valor da Carta (R$)" required>
          <CurrencyInput value={valorCarta} onChange={(v) => setValorCarta(v)} />
        </FormField>
        <FormField label="Data Inicio">
          <input type="date" value={dataInicio} onChange={(e) => setDataInicio(e.target.value)} className={inputClass} />
        </FormField>
        {!editingOp && (
          <div className="grid grid-cols-2 gap-4">
            <FormField label="Num. Parcelas">
              <input type="number" min={1} value={numParcelas} onChange={(e) => setNumParcelas(Number(e.target.value))} className={inputClass} />
            </FormField>
            <FormField label="Valor Parcela (R$)">
              <CurrencyInput value={valorParcela} onChange={(v) => setValorParcela(v)} />
            </FormField>
          </div>
        )}
      </div>
    </Modal>
  );
}

// ── Expandable Row ──────────────────────────────────────────────────────────

function OperacaoRow({ op, onEdit, onDelete, onTogglePagamento }: {
  op: Operacao;
  onEdit: () => void;
  onDelete: () => void;
  onTogglePagamento: (mes: number) => void;
}) {
  const [expanded, setExpanded] = useState(false);
  const pagos = op.pagamentos.filter((p) => p.status === 'pago').length;
  const total = op.pagamentos.length;
  const totalPago = op.pagamentos.filter((p) => p.status === 'pago').reduce((s, p) => s + p.valor, 0);

  return (
    <>
      <tr className="border-b border-somus-border/30 hover:bg-somus-bg-hover">
        <td className="px-4 py-3">
          <button onClick={() => setExpanded(!expanded)} className="p-0.5 hover:bg-somus-bg-tertiary rounded">
            {expanded ? <ChevronDown className="h-4 w-4 text-somus-text-secondary" /> : <ChevronRight className="h-4 w-4 text-somus-text-secondary" />}
          </button>
        </td>
        <td className="px-4 py-3 font-medium text-somus-text-primary">{op.clienteNome}</td>
        <td className="px-4 py-3 text-somus-text-secondary">{op.assessor}</td>
        <td className="px-4 py-3 text-somus-text-secondary">{op.tipo}</td>
        <td className="px-4 py-3 font-medium text-right">{fmtBRL(op.valorCarta)}</td>
        <td className="px-4 py-3">
          <StatusBadge status={mapOpStatus(op.status)} label={op.status.charAt(0).toUpperCase() + op.status.slice(1)} />
        </td>
        <td className="px-4 py-3 text-somus-text-secondary text-sm">{pagos}/{total}</td>
        <td className="px-4 py-3 text-right font-medium">{fmtBRL(totalPago)}</td>
        <td className="px-4 py-3">
          <div className="flex items-center gap-1">
            <button onClick={onEdit} className="p-1.5 hover:bg-somus-bg-tertiary rounded-md" title="Editar"><Edit3 className="h-3.5 w-3.5 text-somus-text-secondary" /></button>
            <button onClick={onDelete} className="p-1.5 hover:bg-red-50 rounded-md" title="Excluir"><Trash2 className="h-3.5 w-3.5 text-red-400" /></button>
          </div>
        </td>
      </tr>
      {expanded && (
        <tr>
          <td colSpan={9} className="px-4 py-0">
            <div className="bg-somus-bg-hover rounded-lg my-2 p-4">
              <h4 className="text-sm font-semibold text-somus-text-primary mb-3">Pagamentos</h4>
              <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-4 lg:grid-cols-6 gap-2">
                {op.pagamentos.map((pg) => (
                  <button
                    key={pg.mes}
                    onClick={() => onTogglePagamento(pg.mes)}
                    className={cn(
                      'flex flex-col items-center p-2.5 rounded-lg border text-xs transition-colors',
                      pg.status === 'pago' ? 'bg-emerald-50 border-emerald-200 text-emerald-700' :
                      pg.status === 'atrasado' ? 'bg-red-50 border-red-200 text-red-700' :
                      'bg-somus-bg-secondary border-somus-border text-somus-text-secondary hover:bg-somus-bg-tertiary',
                    )}
                  >
                    <span className="font-bold text-sm">M{pg.mes}</span>
                    <span className="mt-0.5">{fmtBRL(pg.valor)}</span>
                    <StatusBadge
                      status={mapPgStatus(pg.status)}
                      label={pg.status === 'pago' ? 'Pago' : pg.status === 'atrasado' ? 'Atrasado' : 'Pendente'}
                      className="mt-1 text-[9px]"
                    />
                  </button>
                ))}
              </div>
            </div>
          </td>
        </tr>
      )}
    </>
  );
}

// ── Main Component ──────────────────────────────────────────────────────────

export default function FluxoReceitas() {
  const setPage = useAppStore((s) => s.setPage);
  const [operacoes, setOperacoes] = usePersistedState<Operacao[]>('operacoes', []);
  const [searchQuery, setSearchQuery] = useState('');
  const [modalOpen, setModalOpen] = useState(false);
  const [editingOp, setEditingOp] = useState<Operacao | null>(null);
  const [filterStatus, setFilterStatus] = useState('all');

  const persist = useCallback((ops: Operacao[]) => {
    setOperacoes(ops);
  }, [setOperacoes]);

  function handleSave(op: Operacao) {
    const idx = operacoes.findIndex((o) => o.id === op.id);
    const updated = idx >= 0 ? operacoes.map((o, i) => i === idx ? op : o) : [...operacoes, op];
    persist(updated);
    setModalOpen(false);
    setEditingOp(null);
  }

  function handleDelete(id: string) { persist(operacoes.filter((o) => o.id !== id)); }

  function handleTogglePagamento(opId: string, mes: number) {
    persist(operacoes.map((op) => {
      if (op.id !== opId) return op;
      return {
        ...op,
        pagamentos: op.pagamentos.map((pg) => {
          if (pg.mes !== mes) return pg;
          const next: PagamentoOperacao['status'] = pg.status === 'pendente' ? 'pago' : pg.status === 'pago' ? 'atrasado' : 'pendente';
          return { ...pg, status: next, data: next === 'pago' ? new Date().toISOString().split('T')[0] : pg.data };
        }),
      };
    }));
  }

  const filtered = operacoes.filter((op) => {
    const matchSearch = !searchQuery || op.clienteNome.toLowerCase().includes(searchQuery.toLowerCase()) || op.assessor.toLowerCase().includes(searchQuery.toLowerCase());
    const matchStatus = filterStatus === 'all' || op.status === filterStatus;
    return matchSearch && matchStatus;
  });

  const totalCartas = operacoes.reduce((s, o) => s + o.valorCarta, 0);
  const totalRecebido = operacoes.reduce((s, o) => s + o.pagamentos.filter((p) => p.status === 'pago').reduce((ps, p) => ps + p.valor, 0), 0);
  const ativas = operacoes.filter((o) => o.status === 'ativa').length;

  const FILTER_OPTIONS = [
    { value: 'all', label: 'Todos os status' },
    { value: 'ativa', label: 'Ativa' },
    { value: 'contemplada', label: 'Contemplada' },
    { value: 'encerrada', label: 'Encerrada' },
  ];

  return (
    <PageLayout title="Fluxo de Receitas" subtitle="Acompanhe operacoes e pagamentos">
      <div className="max-w-7xl mx-auto">
        <div className="flex items-center justify-between mb-6">
          <button onClick={() => setPage('dashboard')} className="inline-flex items-center gap-1.5 text-sm text-somus-text-secondary hover:text-somus-text-primary transition-colors">
            <ArrowLeft className="h-4 w-4" /> Voltar ao Dashboard
          </button>
          <Button variant="primary" icon={<Plus className="h-4 w-4" />} onClick={() => { setEditingOp(null); setModalOpen(true); }}>
            Nova Operacao
          </Button>
        </div>

        {/* Summary KPIs */}
        <div className="grid grid-cols-1 sm:grid-cols-3 gap-4 mb-6">
          <KPICard title="Total em Cartas" value={fmtBRL(totalCartas)} icon={<DollarSign className="h-5 w-5" />} />
          <KPICard title="Total Recebido" value={fmtBRL(totalRecebido)} icon={<DollarSign className="h-5 w-5" />} variant="green" />
          <KPICard title="Operacoes Ativas" value={String(ativas)} variant="navy" />
        </div>

        {/* Filters */}
        <div className="flex items-center gap-4 mb-4">
          <div className="relative flex-1 max-w-md">
            <Search className="absolute left-3 top-1/2 -translate-y-1/2 h-4 w-4 text-somus-text-tertiary" />
            <input
              type="text"
              placeholder="Buscar cliente ou assessor..."
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
              className="w-full pl-9 pr-3 py-2 text-sm border border-somus-border rounded-lg focus:ring-2 focus:ring-somus-green/40 focus:border-somus-green outline-none bg-somus-bg-input text-somus-text-primary placeholder:text-somus-text-tertiary"
            />
          </div>
          <div className="w-48">
            <Select options={FILTER_OPTIONS} value={filterStatus} onChange={(e) => setFilterStatus(e.target.value)} />
          </div>
        </div>

        {/* Operations Table */}
        <Card padding="none">
          {filtered.length === 0 ? (
            <div className="flex flex-col items-center justify-center py-16">
              <DollarSign className="h-10 w-10 text-somus-text-tertiary mb-3" />
              <p className="text-sm text-somus-text-tertiary">Nenhuma operacao encontrada</p>
              <Button variant="primary" size="sm" className="mt-4" icon={<Plus className="h-4 w-4" />} onClick={() => { setEditingOp(null); setModalOpen(true); }}>
                Adicionar Operacao
              </Button>
            </div>
          ) : (
            <div className="overflow-x-auto">
              <table className="w-full text-sm">
                <thead>
                  <tr className="bg-somus-bg-hover border-b border-somus-border">
                    <th className="w-10 px-4 py-3" />
                    <th className="text-left px-4 py-3 font-medium text-somus-text-secondary">Cliente</th>
                    <th className="text-left px-4 py-3 font-medium text-somus-text-secondary">Assessor</th>
                    <th className="text-left px-4 py-3 font-medium text-somus-text-secondary">Tipo</th>
                    <th className="text-right px-4 py-3 font-medium text-somus-text-secondary">Valor Carta</th>
                    <th className="text-left px-4 py-3 font-medium text-somus-text-secondary">Status</th>
                    <th className="text-left px-4 py-3 font-medium text-somus-text-secondary">Pgtos</th>
                    <th className="text-right px-4 py-3 font-medium text-somus-text-secondary">Total Pago</th>
                    <th className="w-20 px-4 py-3" />
                  </tr>
                </thead>
                <tbody>
                  {filtered.map((op) => (
                    <OperacaoRow
                      key={op.id}
                      op={op}
                      onEdit={() => { setEditingOp(op); setModalOpen(true); }}
                      onDelete={() => handleDelete(op.id)}
                      onTogglePagamento={(mes) => handleTogglePagamento(op.id, mes)}
                    />
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </Card>
      </div>

      <OperacaoModal
        open={modalOpen}
        onClose={() => { setModalOpen(false); setEditingOp(null); }}
        onSave={handleSave}
        editingOp={editingOp}
      />
    </PageLayout>
  );
}

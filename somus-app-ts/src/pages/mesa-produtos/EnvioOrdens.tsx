import React, { useState, useMemo, useCallback } from 'react';
import { Card } from '@components/Card';
import { Button } from '@components/Button';
import {
  Upload,
  Send,
  FileText,
  Check,
  Clock,
  AlertCircle,
  Search,
  Users,
} from 'lucide-react';
import { cn, formatCurrency } from '@/lib/utils';
import { readExcelFromFile } from '@/services/excel-reader';
import { createDrafts } from '@/services/outlook';

// ── Types ────────────────────────────────────────────────────────────────────

interface Ordem {
  id: string;
  cliente: string;
  ativo: string;
  tipo: 'Compra' | 'Venda';
  quantidade: number;
  preco: number;
  assessor: string;
  status: 'pendente' | 'enviada' | 'executada';
}

// ── Mock Data ────────────────────────────────────────────────────────────────

const MOCK_ORDENS: Ordem[] = [
  { id: '1', cliente: 'Joao Mendes', ativo: 'CDB XP IPCA+7%', tipo: 'Compra', quantidade: 100, preco: 5000, assessor: 'Carlos Silva', status: 'pendente' },
  { id: '2', cliente: 'Maria Costa', ativo: 'LCI Itau 97% CDI', tipo: 'Compra', quantidade: 50, preco: 6000, assessor: 'Carlos Silva', status: 'pendente' },
  { id: '3', cliente: 'Pedro Lima', ativo: 'Debenture VALE25', tipo: 'Venda', quantidade: 200, preco: 3750, assessor: 'Ana Santos', status: 'pendente' },
  { id: '4', cliente: 'Ana Souza', ativo: 'CRA Suzano IPCA+6', tipo: 'Compra', quantidade: 80, preco: 5250, assessor: 'Ana Santos', status: 'pendente' },
  { id: '5', cliente: 'Lucas Ferreira', ativo: 'CRI Cyrela CDI+2%', tipo: 'Compra', quantidade: 150, preco: 4000, assessor: 'Pedro Costa', status: 'pendente' },
  { id: '6', cliente: 'Fernanda Dias', ativo: 'LCA BB 95% CDI', tipo: 'Venda', quantidade: 300, preco: 667, assessor: 'Pedro Costa', status: 'pendente' },
  { id: '7', cliente: 'Bruno Neves', ativo: 'CDB Safra Pre 13%', tipo: 'Compra', quantidade: 120, preco: 2917, assessor: 'Maria Oliveira', status: 'pendente' },
  { id: '8', cliente: 'Camila Rocha', ativo: 'Tesouro IPCA+ 2035', tipo: 'Compra', quantidade: 60, preco: 8333, assessor: 'Joao Souza', status: 'pendente' },
];

const ASSESSORES = [...new Set(MOCK_ORDENS.map((o) => o.assessor))];

// ── Component ────────────────────────────────────────────────────────────────

export default function EnvioOrdens() {
  const [ordens, setOrdens] = useState<Ordem[]>(MOCK_ORDENS);
  const [selectedIds, setSelectedIds] = useState<Set<string>>(new Set());
  const [searchTerm, setSearchTerm] = useState('');
  const [assessorFilter, setAssessorFilter] = useState('TODOS');
  const [uploading, setUploading] = useState(false);
  const [generating, setGenerating] = useState(false);

  const filteredOrdens = useMemo(() => {
    return ordens.filter((o) => {
      if (assessorFilter !== 'TODOS' && o.assessor !== assessorFilter) return false;
      if (searchTerm) {
        const term = searchTerm.toLowerCase();
        return (
          o.cliente.toLowerCase().includes(term) ||
          o.ativo.toLowerCase().includes(term) ||
          o.assessor.toLowerCase().includes(term)
        );
      }
      return true;
    });
  }, [ordens, searchTerm, assessorFilter]);

  const toggleSelect = (id: string) => {
    setSelectedIds((prev) => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  };

  const toggleAll = () => {
    if (selectedIds.size === filteredOrdens.length) {
      setSelectedIds(new Set());
    } else {
      setSelectedIds(new Set(filteredOrdens.map((o) => o.id)));
    }
  };

  const handleUpload = useCallback(async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setUploading(true);
    try {
      const rows = await readExcelFromFile(file);
      const parsed: Ordem[] = rows.map((row, i) => ({
        id: String(i + 1),
        cliente: String(row['Cliente'] || ''),
        ativo: String(row['Ativo'] || ''),
        tipo: (String(row['Tipo'] || 'Compra') as 'Compra' | 'Venda'),
        quantidade: Number(row['Quantidade'] || 0),
        preco: Number(row['Preco'] || 0),
        assessor: String(row['Assessor'] || ''),
        status: 'pendente' as const,
      }));
      if (parsed.length > 0) setOrdens(parsed);
    } catch (err) {
      console.error('Erro ao ler arquivo:', err);
    } finally {
      setUploading(false);
    }
  }, []);

  const handleGerarRascunhos = useCallback(async () => {
    setGenerating(true);
    try {
      const selected = ordens.filter((o) => selectedIds.has(o.id));
      const byAssessor: Record<string, Ordem[]> = {};
      selected.forEach((o) => {
        if (!byAssessor[o.assessor]) byAssessor[o.assessor] = [];
        byAssessor[o.assessor].push(o);
      });

      const emails = Object.entries(byAssessor).map(([assessor, ords]) => ({
        to: `${assessor.toLowerCase().replace(/\s/g, '.')}@somus.com`,
        subject: `Ordens para Execucao - ${new Date().toLocaleDateString('pt-BR')}`,
        body: `<div style="font-family: DM Sans, sans-serif;">
          <h2 style="color: #004D33;">Ordens para Execucao</h2>
          <table style="width: 100%; border-collapse: collapse; font-size: 13px;">
            <tr style="background: #F0F0E0;">
              <th style="padding: 8px; text-align: left;">Cliente</th>
              <th style="padding: 8px; text-align: left;">Ativo</th>
              <th style="padding: 8px; text-align: left;">Tipo</th>
              <th style="padding: 8px; text-align: right;">Qtd</th>
              <th style="padding: 8px; text-align: right;">Preco</th>
            </tr>
            ${ords.map((o) => `
              <tr style="border-bottom: 1px solid #E5E7EB;">
                <td style="padding: 8px;">${o.cliente}</td>
                <td style="padding: 8px;">${o.ativo}</td>
                <td style="padding: 8px;">${o.tipo}</td>
                <td style="padding: 8px; text-align: right;">${o.quantidade}</td>
                <td style="padding: 8px; text-align: right;">${formatCurrency(o.preco)}</td>
              </tr>
            `).join('')}
          </table>
          <p style="color: #888; font-size: 12px; margin-top: 20px;">Somus Capital - Mesa de Produtos</p>
        </div>`,
      }));

      await createDrafts(emails);

      setOrdens((prev) =>
        prev.map((o) =>
          selectedIds.has(o.id) ? { ...o, status: 'enviada' as const } : o
        )
      );
      setSelectedIds(new Set());
    } catch (err) {
      console.error('Erro ao gerar rascunhos:', err);
    } finally {
      setGenerating(false);
    }
  }, [ordens, selectedIds]);

  const statusIcon = (status: Ordem['status']) => {
    switch (status) {
      case 'pendente':
        return <Clock className="h-4 w-4 text-amber-400" />;
      case 'enviada':
        return <Send className="h-4 w-4 text-blue-400" />;
      case 'executada':
        return <Check className="h-4 w-4 text-somus-green-400" />;
    }
  };

  const statusLabel: Record<Ordem['status'], string> = {
    pendente: 'Pendente',
    enviada: 'Enviada',
    executada: 'Executada',
  };

  const statusColor: Record<Ordem['status'], string> = {
    pendente: 'bg-amber-500/10 text-amber-400',
    enviada: 'bg-blue-500/10 text-blue-400',
    executada: 'bg-somus-green-500/10 text-somus-green-400',
  };

  return (
    <div className="space-y-6">
      {/* Header */}
      <div className="flex items-center justify-between">
        <div>
          <h1 className="text-2xl font-bold text-somus-gray-900">
            Envio de Ordens
          </h1>
          <p className="text-sm text-somus-gray-500 mt-1">
            Upload e envio de ordens para assessores
          </p>
        </div>
        <div className="flex items-center gap-3">
          <label className="cursor-pointer">
            <input
              type="file"
              accept=".xlsx,.xls"
              onChange={handleUpload}
              className="hidden"
            />
            <Button
              variant="secondary"
              icon={<Upload className="h-4 w-4" />}
              loading={uploading}
              onClick={() => {}}
            >
              Upload Excel
            </Button>
          </label>
          <Button
            variant="primary"
            icon={<Send className="h-4 w-4" />}
            loading={generating}
            onClick={handleGerarRascunhos}
            disabled={selectedIds.size === 0}
          >
            Gerar Rascunhos ({selectedIds.size})
          </Button>
        </div>
      </div>

      {/* Filters */}
      <Card padding="sm">
        <div className="flex flex-wrap items-center gap-4 p-2">
          <div className="relative">
            <Search className="absolute left-3 top-1/2 -translate-y-1/2 h-4 w-4 text-somus-gray-400" />
            <input
              type="text"
              placeholder="Buscar..."
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
              className="text-sm border border-somus-gray-300 rounded-lg pl-9 pr-3 py-1.5 w-56 focus:ring-2 focus:ring-somus-green/40 focus:outline-none"
            />
          </div>
          <select
            value={assessorFilter}
            onChange={(e) => setAssessorFilter(e.target.value)}
            className="text-sm border border-somus-gray-300 rounded-lg px-3 py-1.5 focus:ring-2 focus:ring-somus-green/40 focus:outline-none"
          >
            <option value="TODOS">Todos assessores</option>
            {ASSESSORES.map((a) => (
              <option key={a} value={a}>{a}</option>
            ))}
          </select>
        </div>
      </Card>

      {/* Table */}
      <Card title={`Ordens (${filteredOrdens.length})`}>
        <div className="mt-4 overflow-x-auto">
          <table className="w-full text-sm">
            <thead>
              <tr className="border-b border-somus-gray-200">
                <th className="py-3 px-4 w-10">
                  <input
                    type="checkbox"
                    checked={selectedIds.size === filteredOrdens.length && filteredOrdens.length > 0}
                    onChange={toggleAll}
                    className="w-4 h-4 rounded border-somus-gray-300 text-somus-green focus:ring-somus-green/40"
                  />
                </th>
                <th className="text-left py-3 px-4 font-semibold text-somus-gray-600">Cliente</th>
                <th className="text-left py-3 px-4 font-semibold text-somus-gray-600">Ativo</th>
                <th className="text-left py-3 px-4 font-semibold text-somus-gray-600">Tipo</th>
                <th className="text-right py-3 px-4 font-semibold text-somus-gray-600">Qtd</th>
                <th className="text-right py-3 px-4 font-semibold text-somus-gray-600">Preco</th>
                <th className="text-left py-3 px-4 font-semibold text-somus-gray-600">Assessor</th>
                <th className="text-left py-3 px-4 font-semibold text-somus-gray-600">Status</th>
              </tr>
            </thead>
            <tbody>
              {filteredOrdens.map((o) => (
                <tr
                  key={o.id}
                  className={cn(
                    'border-b border-somus-gray-100 hover:bg-somus-gray-50 transition-colors',
                    selectedIds.has(o.id) && 'bg-somus-green/5'
                  )}
                >
                  <td className="py-3 px-4">
                    <input
                      type="checkbox"
                      checked={selectedIds.has(o.id)}
                      onChange={() => toggleSelect(o.id)}
                      className="w-4 h-4 rounded border-somus-gray-300 text-somus-green focus:ring-somus-green/40"
                    />
                  </td>
                  <td className="py-3 px-4 font-medium text-somus-gray-900">{o.cliente}</td>
                  <td className="py-3 px-4 text-somus-gray-700">{o.ativo}</td>
                  <td className="py-3 px-4">
                    <span
                      className={cn(
                        'inline-block px-2 py-0.5 text-xs rounded-full font-medium',
                        o.tipo === 'Compra'
                          ? 'bg-somus-green-500/10 text-somus-green-400'
                          : 'bg-red-500/10 text-red-400'
                      )}
                    >
                      {o.tipo}
                    </span>
                  </td>
                  <td className="py-3 px-4 text-right text-somus-gray-700">{o.quantidade}</td>
                  <td className="py-3 px-4 text-right font-medium text-somus-gray-900">
                    {formatCurrency(o.preco)}
                  </td>
                  <td className="py-3 px-4 text-somus-gray-600">{o.assessor}</td>
                  <td className="py-3 px-4">
                    <span className={cn('inline-flex items-center gap-1 px-2 py-0.5 text-xs rounded-full font-medium', statusColor[o.status])}>
                      {statusIcon(o.status)}
                      {statusLabel[o.status]}
                    </span>
                  </td>
                </tr>
              ))}
              {filteredOrdens.length === 0 && (
                <tr>
                  <td colSpan={8} className="py-12 text-center text-somus-gray-400">
                    Nenhuma ordem encontrada
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </Card>
    </div>
  );
}

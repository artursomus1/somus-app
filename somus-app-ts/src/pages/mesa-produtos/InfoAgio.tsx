import React, { useState, useMemo, useCallback } from 'react';
import { Card } from '@components/Card';
import { Button } from '@components/Button';
import {
  Upload,
  Search,
  ArrowUpDown,
  TrendingUp,
  TrendingDown,
  Filter,
} from 'lucide-react';
import { cn, formatCurrency, formatPercent } from '@/lib/utils';
import { readExcelFromFile } from '@/services/excel-reader';

// ── Types ────────────────────────────────────────────────────────────────────

interface AgioItem {
  id: string;
  cliente: string;
  ativo: string;
  tipo: string;
  agioPct: number;
  valor: number;
  data: string;
}

type SortField = 'cliente' | 'ativo' | 'tipo' | 'agioPct' | 'valor' | 'data';
type SortDir = 'asc' | 'desc';

// ── Mock Data ────────────────────────────────────────────────────────────────

const MOCK_DATA: AgioItem[] = [
  { id: '1', cliente: 'Joao Mendes', ativo: 'CDB XP IPCA+7%', tipo: 'CDB', agioPct: 2.35, valor: 500000, data: '2026-03-20' },
  { id: '2', cliente: 'Maria Costa', ativo: 'LCI Itau 97% CDI', tipo: 'LCI', agioPct: -1.5, valor: 300000, data: '2026-03-19' },
  { id: '3', cliente: 'Carlos Ferreira', ativo: 'Debenture VALE25', tipo: 'Debenture', agioPct: 3.1, valor: 750000, data: '2026-03-18' },
  { id: '4', cliente: 'Ana Souza', ativo: 'CRA Suzano IPCA+6', tipo: 'CRA', agioPct: -0.8, valor: 420000, data: '2026-03-17' },
  { id: '5', cliente: 'Pedro Lima', ativo: 'CRI Cyrela CDI+2%', tipo: 'CRI', agioPct: 1.75, valor: 600000, data: '2026-03-16' },
  { id: '6', cliente: 'Fernanda Alves', ativo: 'LCA BB 95% CDI', tipo: 'LCA', agioPct: -2.1, valor: 200000, data: '2026-03-15' },
  { id: '7', cliente: 'Lucas Santos', ativo: 'CDB Safra Pre 13%', tipo: 'CDB', agioPct: 0.45, valor: 350000, data: '2026-03-14' },
  { id: '8', cliente: 'Juliana Rocha', ativo: 'Debenture PETR26', tipo: 'Debenture', agioPct: -3.2, valor: 900000, data: '2026-03-13' },
  { id: '9', cliente: 'Bruno Neves', ativo: 'CDB BTG IPCA+6.5', tipo: 'CDB', agioPct: 1.9, valor: 480000, data: '2026-03-12' },
  { id: '10', cliente: 'Camila Dias', ativo: 'LCI Bradesco 93%', tipo: 'LCI', agioPct: -0.3, valor: 150000, data: '2026-03-11' },
  { id: '11', cliente: 'Ricardo Gomes', ativo: 'CRA JBS IPCA+5.5', tipo: 'CRA', agioPct: 4.2, valor: 1200000, data: '2026-03-10' },
  { id: '12', cliente: 'Tatiana Mello', ativo: 'CRI LOG CP CDI+3', tipo: 'CRI', agioPct: -1.8, valor: 550000, data: '2026-03-09' },
];

const TIPOS = [...new Set(MOCK_DATA.map((d) => d.tipo))];

// ── Component ────────────────────────────────────────────────────────────────

export default function InfoAgio() {
  const [data, setData] = useState<AgioItem[]>(MOCK_DATA);
  const [searchTerm, setSearchTerm] = useState('');
  const [tipoFilter, setTipoFilter] = useState('TODOS');
  const [agioFilter, setAgioFilter] = useState<'TODOS' | 'AGIO' | 'DESAGIO'>('TODOS');
  const [sortField, setSortField] = useState<SortField>('data');
  const [sortDir, setSortDir] = useState<SortDir>('desc');
  const [uploading, setUploading] = useState(false);

  const handleSort = (field: SortField) => {
    if (sortField === field) {
      setSortDir((d) => (d === 'asc' ? 'desc' : 'asc'));
    } else {
      setSortField(field);
      setSortDir('asc');
    }
  };

  const filteredData = useMemo(() => {
    let result = [...data];

    // Search
    if (searchTerm) {
      const term = searchTerm.toLowerCase();
      result = result.filter(
        (d) =>
          d.cliente.toLowerCase().includes(term) ||
          d.ativo.toLowerCase().includes(term) ||
          d.tipo.toLowerCase().includes(term)
      );
    }

    // Tipo
    if (tipoFilter !== 'TODOS') {
      result = result.filter((d) => d.tipo === tipoFilter);
    }

    // Agio/Desagio
    if (agioFilter === 'AGIO') {
      result = result.filter((d) => d.agioPct > 0);
    } else if (agioFilter === 'DESAGIO') {
      result = result.filter((d) => d.agioPct < 0);
    }

    // Sort
    result.sort((a, b) => {
      let cmp = 0;
      if (sortField === 'agioPct' || sortField === 'valor') {
        cmp = a[sortField] - b[sortField];
      } else {
        cmp = String(a[sortField]).localeCompare(String(b[sortField]));
      }
      return sortDir === 'asc' ? cmp : -cmp;
    });

    return result;
  }, [data, searchTerm, tipoFilter, agioFilter, sortField, sortDir]);

  const handleUpload = useCallback(async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setUploading(true);
    try {
      const rows = await readExcelFromFile(file);
      const parsed: AgioItem[] = rows.map((row, i) => ({
        id: String(i + 1),
        cliente: String(row['Cliente'] || row['cliente'] || ''),
        ativo: String(row['Ativo'] || row['ativo'] || ''),
        tipo: String(row['Tipo'] || row['tipo'] || ''),
        agioPct: Number(row['Agio/Desagio %'] || row['agioPct'] || row['Agio'] || 0),
        valor: Number(row['Valor'] || row['valor'] || 0),
        data: String(row['Data'] || row['data'] || ''),
      }));
      if (parsed.length > 0) setData(parsed);
    } catch (err) {
      console.error('Erro ao ler arquivo:', err);
    } finally {
      setUploading(false);
    }
  }, []);

  const formatDate = (dateStr: string) => {
    if (!dateStr) return '';
    const d = new Date(dateStr + 'T00:00:00');
    return d.toLocaleDateString('pt-BR');
  };

  // Stats
  const totalAgio = filteredData.filter((d) => d.agioPct > 0).length;
  const totalDesagio = filteredData.filter((d) => d.agioPct < 0).length;
  const avgAgio =
    filteredData.length > 0
      ? filteredData.reduce((s, d) => s + d.agioPct, 0) / filteredData.length
      : 0;

  const SortHeader = ({ field, children }: { field: SortField; children: React.ReactNode }) => (
    <th
      className="text-left py-3 px-4 font-semibold text-somus-gray-600 cursor-pointer select-none hover:text-somus-green transition-colors"
      onClick={() => handleSort(field)}
    >
      <div className="flex items-center gap-1">
        {children}
        <ArrowUpDown className={cn('h-3 w-3', sortField === field ? 'text-somus-green' : 'text-somus-gray-300')} />
      </div>
    </th>
  );

  return (
    <div className="space-y-6">
      {/* Header */}
      <div className="flex items-center justify-between">
        <div>
          <h1 className="text-2xl font-bold text-somus-gray-900">
            Tabela de Agio / Desagio
          </h1>
          <p className="text-sm text-somus-gray-500 mt-1">
            Visualizacao de premios e descontos em ativos de renda fixa
          </p>
        </div>
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
      </div>

      {/* Stats */}
      <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
        <div className="bg-somus-bg-secondary rounded-lg border border-somus-gray-200 p-4">
          <div className="flex items-center gap-2 mb-2">
            <TrendingUp className="h-4 w-4 text-somus-green-400" />
            <span className="text-sm text-somus-gray-500">Com Agio</span>
          </div>
          <div className="text-2xl font-bold text-somus-green-400">{totalAgio}</div>
        </div>
        <div className="bg-somus-bg-secondary rounded-lg border border-somus-gray-200 p-4">
          <div className="flex items-center gap-2 mb-2">
            <TrendingDown className="h-4 w-4 text-red-400" />
            <span className="text-sm text-somus-gray-500">Com Desagio</span>
          </div>
          <div className="text-2xl font-bold text-red-400">{totalDesagio}</div>
        </div>
        <div className="bg-somus-bg-secondary rounded-lg border border-somus-gray-200 p-4">
          <div className="flex items-center gap-2 mb-2">
            <ArrowUpDown className="h-4 w-4 text-somus-gray-400" />
            <span className="text-sm text-somus-gray-500">Media Agio</span>
          </div>
          <div className={cn('text-2xl font-bold', avgAgio >= 0 ? 'text-somus-green-400' : 'text-red-400')}>
            {formatPercent(avgAgio)}
          </div>
        </div>
      </div>

      {/* Filters */}
      <Card padding="sm">
        <div className="flex flex-wrap items-center gap-4 p-2">
          <div className="flex items-center gap-2">
            <Filter className="h-4 w-4 text-somus-gray-400" />
          </div>
          <div className="relative">
            <Search className="absolute left-3 top-1/2 -translate-y-1/2 h-4 w-4 text-somus-gray-400" />
            <input
              type="text"
              placeholder="Buscar cliente, ativo..."
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
              className="text-sm border border-somus-gray-300 rounded-lg pl-9 pr-3 py-1.5 w-56 focus:ring-2 focus:ring-somus-green/40 focus:outline-none"
            />
          </div>
          <select
            value={tipoFilter}
            onChange={(e) => setTipoFilter(e.target.value)}
            className="text-sm border border-somus-gray-300 rounded-lg px-3 py-1.5 focus:ring-2 focus:ring-somus-green/40 focus:outline-none"
          >
            <option value="TODOS">Todos os tipos</option>
            {TIPOS.map((t) => (
              <option key={t} value={t}>{t}</option>
            ))}
          </select>
          <select
            value={agioFilter}
            onChange={(e) => setAgioFilter(e.target.value as any)}
            className="text-sm border border-somus-gray-300 rounded-lg px-3 py-1.5 focus:ring-2 focus:ring-somus-green/40 focus:outline-none"
          >
            <option value="TODOS">Agio e Desagio</option>
            <option value="AGIO">Somente Agio</option>
            <option value="DESAGIO">Somente Desagio</option>
          </select>
        </div>
      </Card>

      {/* Table */}
      <Card title={`Dados (${filteredData.length} registros)`}>
        <div className="mt-4 overflow-x-auto">
          <table className="w-full text-sm">
            <thead>
              <tr className="border-b border-somus-gray-200">
                <SortHeader field="cliente">Cliente</SortHeader>
                <SortHeader field="ativo">Ativo</SortHeader>
                <SortHeader field="tipo">Tipo</SortHeader>
                <SortHeader field="agioPct">Agio/Desagio %</SortHeader>
                <SortHeader field="valor">Valor</SortHeader>
                <SortHeader field="data">Data</SortHeader>
              </tr>
            </thead>
            <tbody>
              {filteredData.map((item) => (
                <tr
                  key={item.id}
                  className="border-b border-somus-gray-100 hover:bg-somus-gray-50 transition-colors"
                >
                  <td className="py-3 px-4 font-medium text-somus-gray-900">{item.cliente}</td>
                  <td className="py-3 px-4 text-somus-gray-700">{item.ativo}</td>
                  <td className="py-3 px-4">
                    <span className="inline-block px-2 py-0.5 text-xs rounded-full bg-somus-gray-100 text-somus-gray-700 font-medium">
                      {item.tipo}
                    </span>
                  </td>
                  <td className="py-3 px-4">
                    <span
                      className={cn(
                        'inline-flex items-center gap-1 px-2 py-0.5 text-xs rounded-full font-semibold',
                        item.agioPct > 0
                          ? 'bg-somus-green-500/10 text-somus-green-400'
                          : item.agioPct < 0
                          ? 'bg-red-500/10 text-red-400'
                          : 'bg-somus-gray-100 text-somus-gray-600'
                      )}
                    >
                      {item.agioPct > 0 ? (
                        <TrendingUp className="h-3 w-3" />
                      ) : item.agioPct < 0 ? (
                        <TrendingDown className="h-3 w-3" />
                      ) : null}
                      {formatPercent(item.agioPct)}
                    </span>
                  </td>
                  <td className="py-3 px-4 text-right font-medium text-somus-gray-900">
                    {formatCurrency(item.valor)}
                  </td>
                  <td className="py-3 px-4 text-somus-gray-600">{formatDate(item.data)}</td>
                </tr>
              ))}
              {filteredData.length === 0 && (
                <tr>
                  <td colSpan={6} className="py-12 text-center text-somus-gray-400">
                    Nenhum registro encontrado
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

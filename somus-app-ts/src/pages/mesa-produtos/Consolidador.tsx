import React, { useState, useMemo, useCallback } from 'react';
import { Card } from '@components/Card';
import { Button } from '@components/Button';
import {
  Upload,
  Download,
  FileSpreadsheet,
  Trash2,
  Combine,
  Search,
  Users,
  DollarSign,
  Briefcase,
} from 'lucide-react';
import { cn, formatCurrency } from '@/lib/utils';
import { readExcelFromFile, writeExcelFromJSON } from '@/services/excel-reader';

// ── Types ────────────────────────────────────────────────────────────────────

interface PortfolioItem {
  cliente: string;
  cpfCnpj: string;
  assessor: string;
  ativo: string;
  tipo: string;
  quantidade: number;
  precoMedio: number;
  valorAtual: number;
  resultado: number;
}

interface ConsolidatedClient {
  cliente: string;
  cpfCnpj: string;
  assessor: string;
  totalAtivos: number;
  valorTotal: number;
  resultadoTotal: number;
  items: PortfolioItem[];
}

// ── Mock Data ────────────────────────────────────────────────────────────────

const MOCK_PORTFOLIO: PortfolioItem[] = [
  { cliente: 'Joao Mendes', cpfCnpj: '123.456.789-00', assessor: 'Carlos Silva', ativo: 'CDB XP IPCA+7%', tipo: 'Renda Fixa', quantidade: 1, precoMedio: 500000, valorAtual: 520000, resultado: 20000 },
  { cliente: 'Joao Mendes', cpfCnpj: '123.456.789-00', assessor: 'Carlos Silva', ativo: 'Fundo XP Macro', tipo: 'Fundos', quantidade: 1000, precoMedio: 15.5, valorAtual: 16.2, resultado: 700 },
  { cliente: 'Maria Costa', cpfCnpj: '987.654.321-00', assessor: 'Carlos Silva', ativo: 'LCI Itau 97%', tipo: 'Renda Fixa', quantidade: 1, precoMedio: 300000, valorAtual: 310000, resultado: 10000 },
  { cliente: 'Maria Costa', cpfCnpj: '987.654.321-00', assessor: 'Carlos Silva', ativo: 'PETR4', tipo: 'Renda Variavel', quantidade: 500, precoMedio: 32.5, valorAtual: 35.8, resultado: 1650 },
  { cliente: 'Pedro Lima', cpfCnpj: '456.789.123-00', assessor: 'Ana Santos', ativo: 'Debenture VALE25', tipo: 'Renda Fixa', quantidade: 1, precoMedio: 750000, valorAtual: 780000, resultado: 30000 },
  { cliente: 'Pedro Lima', cpfCnpj: '456.789.123-00', assessor: 'Ana Santos', ativo: 'CRA Suzano', tipo: 'Renda Fixa', quantidade: 1, precoMedio: 420000, valorAtual: 430000, resultado: 10000 },
  { cliente: 'Ana Souza', cpfCnpj: '321.654.987-00', assessor: 'Pedro Costa', ativo: 'CDB Safra Pre', tipo: 'Renda Fixa', quantidade: 1, precoMedio: 200000, valorAtual: 208000, resultado: 8000 },
  { cliente: 'Ana Souza', cpfCnpj: '321.654.987-00', assessor: 'Pedro Costa', ativo: 'Tesouro Selic', tipo: 'Renda Fixa', quantidade: 5, precoMedio: 14200, valorAtual: 14500, resultado: 1500 },
  { cliente: 'Lucas Ferreira', cpfCnpj: '654.321.987-00', assessor: 'Maria Oliveira', ativo: 'VALE3', tipo: 'Renda Variavel', quantidade: 300, precoMedio: 68.5, valorAtual: 72.3, resultado: 1140 },
  { cliente: 'Lucas Ferreira', cpfCnpj: '654.321.987-00', assessor: 'Maria Oliveira', ativo: 'Fundo Verde', tipo: 'Fundos', quantidade: 2000, precoMedio: 45.0, valorAtual: 47.5, resultado: 5000 },
];

// ── Component ────────────────────────────────────────────────────────────────

export default function Consolidador() {
  const [portfolio, setPortfolio] = useState<PortfolioItem[]>(MOCK_PORTFOLIO);
  const [searchTerm, setSearchTerm] = useState('');
  const [uploading, setUploading] = useState(false);
  const [expandedClients, setExpandedClients] = useState<Set<string>>(new Set());

  const consolidated = useMemo(() => {
    const map: Record<string, ConsolidatedClient> = {};
    portfolio.forEach((item) => {
      const key = item.cpfCnpj;
      if (!map[key]) {
        map[key] = {
          cliente: item.cliente,
          cpfCnpj: item.cpfCnpj,
          assessor: item.assessor,
          totalAtivos: 0,
          valorTotal: 0,
          resultadoTotal: 0,
          items: [],
        };
      }
      map[key].items.push(item);
      map[key].totalAtivos++;
      map[key].valorTotal += item.valorAtual;
      map[key].resultadoTotal += item.resultado;
    });
    return Object.values(map);
  }, [portfolio]);

  const filteredConsolidated = useMemo(() => {
    if (!searchTerm) return consolidated;
    const term = searchTerm.toLowerCase();
    return consolidated.filter(
      (c) =>
        c.cliente.toLowerCase().includes(term) ||
        c.assessor.toLowerCase().includes(term) ||
        c.cpfCnpj.includes(term)
    );
  }, [consolidated, searchTerm]);

  const totalClientes = filteredConsolidated.length;
  const totalValor = filteredConsolidated.reduce((s, c) => s + c.valorTotal, 0);
  const totalResultado = filteredConsolidated.reduce((s, c) => s + c.resultadoTotal, 0);

  const toggleExpand = (cpf: string) => {
    setExpandedClients((prev) => {
      const next = new Set(prev);
      if (next.has(cpf)) next.delete(cpf);
      else next.add(cpf);
      return next;
    });
  };

  const handleUpload = useCallback(async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setUploading(true);
    try {
      const rows = await readExcelFromFile(file);
      const parsed: PortfolioItem[] = rows.map((row) => ({
        cliente: String(row['Cliente'] || ''),
        cpfCnpj: String(row['CPF/CNPJ'] || row['cpfCnpj'] || ''),
        assessor: String(row['Assessor'] || ''),
        ativo: String(row['Ativo'] || ''),
        tipo: String(row['Tipo'] || ''),
        quantidade: Number(row['Quantidade'] || 0),
        precoMedio: Number(row['Preco Medio'] || row['precoMedio'] || 0),
        valorAtual: Number(row['Valor Atual'] || row['valorAtual'] || 0),
        resultado: Number(row['Resultado'] || row['resultado'] || 0),
      }));
      if (parsed.length > 0) setPortfolio(parsed);
    } catch (err) {
      console.error('Erro ao ler arquivo:', err);
    } finally {
      setUploading(false);
    }
  }, []);

  const handleExport = useCallback(() => {
    const exportData = filteredConsolidated.map((c) => ({
      Cliente: c.cliente,
      'CPF/CNPJ': c.cpfCnpj,
      Assessor: c.assessor,
      'Total Ativos': c.totalAtivos,
      'Valor Total': c.valorTotal,
      'Resultado Total': c.resultadoTotal,
    }));
    writeExcelFromJSON(exportData, 'Portfolio_Consolidado.xlsx');
  }, [filteredConsolidated]);

  return (
    <div className="space-y-6">
      {/* Header */}
      <div className="flex items-center justify-between">
        <div>
          <h1 className="text-2xl font-bold text-somus-gray-900">
            Consolidador de Carteiras
          </h1>
          <p className="text-sm text-somus-gray-500 mt-1">
            Consolidacao de portfolios por cliente
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
              Upload Carteira
            </Button>
          </label>
          <Button
            variant="secondary"
            icon={<Download className="h-4 w-4" />}
            onClick={handleExport}
          >
            Exportar
          </Button>
        </div>
      </div>

      {/* KPIs */}
      <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
        <div className="bg-white rounded-lg shadow-sm border border-somus-gray-200 p-5">
          <div className="flex items-center gap-2 mb-2">
            <Users className="h-5 w-5 text-somus-green" />
            <span className="text-sm text-somus-gray-500">Total Clientes</span>
          </div>
          <div className="text-2xl font-bold text-somus-gray-900">{totalClientes}</div>
        </div>
        <div className="bg-white rounded-lg shadow-sm border border-somus-gray-200 p-5">
          <div className="flex items-center gap-2 mb-2">
            <Briefcase className="h-5 w-5 text-somus-green" />
            <span className="text-sm text-somus-gray-500">Valor Total</span>
          </div>
          <div className="text-2xl font-bold text-somus-gray-900">{formatCurrency(totalValor)}</div>
        </div>
        <div className="bg-white rounded-lg shadow-sm border border-somus-gray-200 p-5">
          <div className="flex items-center gap-2 mb-2">
            <DollarSign className="h-5 w-5 text-somus-green" />
            <span className="text-sm text-somus-gray-500">Resultado Total</span>
          </div>
          <div className={cn('text-2xl font-bold', totalResultado >= 0 ? 'text-emerald-600' : 'text-red-500')}>
            {formatCurrency(totalResultado)}
          </div>
        </div>
      </div>

      {/* Search */}
      <div className="relative">
        <Search className="absolute left-3 top-1/2 -translate-y-1/2 h-4 w-4 text-somus-gray-400" />
        <input
          type="text"
          placeholder="Buscar por cliente, assessor ou CPF/CNPJ..."
          value={searchTerm}
          onChange={(e) => setSearchTerm(e.target.value)}
          className="w-full text-sm border border-somus-gray-300 rounded-lg pl-9 pr-3 py-2 focus:ring-2 focus:ring-somus-green/40 focus:outline-none"
        />
      </div>

      {/* Consolidated Table */}
      <Card title="Carteiras Consolidadas">
        <div className="mt-4 space-y-2">
          {filteredConsolidated.map((client) => {
            const isExpanded = expandedClients.has(client.cpfCnpj);
            return (
              <div key={client.cpfCnpj} className="border border-somus-gray-200 rounded-lg overflow-hidden">
                {/* Client Row */}
                <button
                  onClick={() => toggleExpand(client.cpfCnpj)}
                  className="w-full flex items-center justify-between py-3 px-4 bg-white hover:bg-somus-gray-50 transition-colors"
                >
                  <div className="flex items-center gap-4">
                    <div className={cn(
                      'w-5 h-5 flex items-center justify-center transition-transform text-somus-gray-400',
                      isExpanded && 'rotate-90'
                    )}>
                      <svg viewBox="0 0 16 16" fill="currentColor" className="w-3 h-3">
                        <path d="M6 3l5 5-5 5V3z" />
                      </svg>
                    </div>
                    <div className="text-left">
                      <div className="text-sm font-semibold text-somus-gray-900">
                        {client.cliente}
                      </div>
                      <div className="text-xs text-somus-gray-400">
                        {client.cpfCnpj} | {client.assessor}
                      </div>
                    </div>
                  </div>
                  <div className="flex items-center gap-8">
                    <div className="text-right">
                      <div className="text-xs text-somus-gray-400">Ativos</div>
                      <div className="text-sm font-medium text-somus-gray-700">{client.totalAtivos}</div>
                    </div>
                    <div className="text-right">
                      <div className="text-xs text-somus-gray-400">Valor</div>
                      <div className="text-sm font-medium text-somus-gray-900">{formatCurrency(client.valorTotal)}</div>
                    </div>
                    <div className="text-right">
                      <div className="text-xs text-somus-gray-400">Resultado</div>
                      <div className={cn('text-sm font-medium', client.resultadoTotal >= 0 ? 'text-emerald-600' : 'text-red-500')}>
                        {formatCurrency(client.resultadoTotal)}
                      </div>
                    </div>
                  </div>
                </button>

                {/* Items */}
                {isExpanded && (
                  <div className="border-t border-somus-gray-200 bg-somus-gray-50">
                    <table className="w-full text-sm">
                      <thead>
                        <tr className="border-b border-somus-gray-200">
                          <th className="text-left py-2 px-4 font-semibold text-somus-gray-500 text-xs">Ativo</th>
                          <th className="text-left py-2 px-4 font-semibold text-somus-gray-500 text-xs">Tipo</th>
                          <th className="text-right py-2 px-4 font-semibold text-somus-gray-500 text-xs">Qtd</th>
                          <th className="text-right py-2 px-4 font-semibold text-somus-gray-500 text-xs">Preco Medio</th>
                          <th className="text-right py-2 px-4 font-semibold text-somus-gray-500 text-xs">Valor Atual</th>
                          <th className="text-right py-2 px-4 font-semibold text-somus-gray-500 text-xs">Resultado</th>
                        </tr>
                      </thead>
                      <tbody>
                        {client.items.map((item, idx) => (
                          <tr key={idx} className="border-b border-somus-gray-100">
                            <td className="py-2 px-4 text-somus-gray-700">{item.ativo}</td>
                            <td className="py-2 px-4">
                              <span className="inline-block px-2 py-0.5 text-xs rounded-full bg-somus-gray-200 text-somus-gray-600">
                                {item.tipo}
                              </span>
                            </td>
                            <td className="py-2 px-4 text-right text-somus-gray-600">{item.quantidade}</td>
                            <td className="py-2 px-4 text-right text-somus-gray-600">{formatCurrency(item.precoMedio)}</td>
                            <td className="py-2 px-4 text-right font-medium text-somus-gray-900">{formatCurrency(item.valorAtual)}</td>
                            <td className={cn('py-2 px-4 text-right font-medium', item.resultado >= 0 ? 'text-emerald-600' : 'text-red-500')}>
                              {formatCurrency(item.resultado)}
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                )}
              </div>
            );
          })}

          {filteredConsolidated.length === 0 && (
            <div className="text-center py-12 text-somus-gray-400">
              Nenhum cliente encontrado
            </div>
          )}
        </div>
      </Card>
    </div>
  );
}

import React, { useState, useMemo, useCallback } from 'react';
import { Card } from '@components/Card';
import { Button } from '@components/Button';
import {
  Upload,
  Download,
  Filter,
  DollarSign,
} from 'lucide-react';
import { formatCurrency } from '@/lib/utils';
import { readExcelFromFile, writeExcelFromJSON } from '@/services/excel-reader';

// ── Types ────────────────────────────────────────────────────────────────────

interface ReceitaRow {
  assessor: string;
  equipe: string;
  rendaFixa: number;
  rendaVariavel: number;
  fundos: number;
  coe: number;
  previdencia: number;
  seguros: number;
  cambio: number;
  consorcio: number;
  total: number;
}

// ── Mock Data ────────────────────────────────────────────────────────────────

const EQUIPES = ['SP', 'LEBLON', 'PRODUTOS', 'CORPORATE', 'BACKOFFICE'];
const PRODUTOS = ['rendaFixa', 'rendaVariavel', 'fundos', 'coe', 'previdencia', 'seguros', 'cambio', 'consorcio'] as const;
const PRODUTO_LABELS: Record<string, string> = {
  rendaFixa: 'Renda Fixa',
  rendaVariavel: 'Renda Variavel',
  fundos: 'Fundos',
  coe: 'COE',
  previdencia: 'Previdencia',
  seguros: 'Seguros',
  cambio: 'Cambio',
  consorcio: 'Consorcio',
};

function generateMockData(): ReceitaRow[] {
  const assessores = [
    { nome: 'Carlos Silva', equipe: 'SP' },
    { nome: 'Ana Santos', equipe: 'SP' },
    { nome: 'Roberto Alves', equipe: 'SP' },
    { nome: 'Pedro Costa', equipe: 'LEBLON' },
    { nome: 'Maria Oliveira', equipe: 'LEBLON' },
    { nome: 'Joao Souza', equipe: 'PRODUTOS' },
    { nome: 'Fernanda Lima', equipe: 'CORPORATE' },
    { nome: 'Lucas Almeida', equipe: 'CORPORATE' },
    { nome: 'Juliana Rocha', equipe: 'BACKOFFICE' },
    { nome: 'Bruno Ferreira', equipe: 'BACKOFFICE' },
  ];

  return assessores.map((a) => {
    const row: any = {
      assessor: a.nome,
      equipe: a.equipe,
    };
    let total = 0;
    PRODUTOS.forEach((p) => {
      const val = Math.round(Math.random() * 200000 + 10000);
      row[p] = val;
      total += val;
    });
    row.total = total;
    return row as ReceitaRow;
  });
}

const MESES = [
  'Janeiro', 'Fevereiro', 'Marco', 'Abril', 'Maio', 'Junho',
  'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro',
];

// ── Component ────────────────────────────────────────────────────────────────

export default function CtrlReceita() {
  const [data, setData] = useState<ReceitaRow[]>(generateMockData);
  const [equipeFilter, setEquipeFilter] = useState('TODAS');
  const [mesFilter, setMesFilter] = useState(new Date().getMonth());
  const [anoFilter, setAnoFilter] = useState(new Date().getFullYear());
  const [uploading, setUploading] = useState(false);

  const filteredData = useMemo(() => {
    if (equipeFilter === 'TODAS') return data;
    return data.filter((d) => d.equipe === equipeFilter);
  }, [data, equipeFilter]);

  // Totals
  const totals = useMemo(() => {
    const t: Record<string, number> = {};
    PRODUTOS.forEach((p) => {
      t[p] = filteredData.reduce((s, r) => s + (r as any)[p], 0);
    });
    t.total = filteredData.reduce((s, r) => s + r.total, 0);
    return t;
  }, [filteredData]);

  // Subtotals by equipe
  const subtotalsByEquipe = useMemo(() => {
    const map: Record<string, Record<string, number>> = {};
    filteredData.forEach((row) => {
      if (!map[row.equipe]) {
        map[row.equipe] = {};
        PRODUTOS.forEach((p) => (map[row.equipe][p] = 0));
        map[row.equipe].total = 0;
      }
      PRODUTOS.forEach((p) => {
        map[row.equipe][p] += (row as any)[p];
      });
      map[row.equipe].total += row.total;
    });
    return map;
  }, [filteredData]);

  const handleUpload = useCallback(async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setUploading(true);
    try {
      const rows = await readExcelFromFile(file);
      const parsed: ReceitaRow[] = rows.map((row) => {
        const r: any = {
          assessor: String(row['Assessor'] || ''),
          equipe: String(row['Equipe'] || 'SP'),
        };
        let total = 0;
        PRODUTOS.forEach((p) => {
          const label = PRODUTO_LABELS[p];
          const val = Number(row[label] || row[p] || 0);
          r[p] = val;
          total += val;
        });
        r.total = total;
        return r as ReceitaRow;
      });
      if (parsed.length > 0) setData(parsed);
    } catch (err) {
      console.error('Erro ao ler arquivo:', err);
    } finally {
      setUploading(false);
    }
  }, []);

  const handleExport = useCallback(() => {
    const exportData = filteredData.map((row) => {
      const r: Record<string, any> = {
        Assessor: row.assessor,
        Equipe: row.equipe,
      };
      PRODUTOS.forEach((p) => {
        r[PRODUTO_LABELS[p]] = (row as any)[p];
      });
      r['Total'] = row.total;
      return r;
    });
    writeExcelFromJSON(exportData, `Controle_Receita_${MESES[mesFilter]}_${anoFilter}.xlsx`);
  }, [filteredData, mesFilter, anoFilter]);

  const anos = Array.from({ length: 5 }, (_, i) => new Date().getFullYear() - i);

  return (
    <div className="space-y-6">
      {/* Header */}
      <div className="flex items-center justify-between">
        <div>
          <h1 className="text-2xl font-bold text-somus-gray-900">
            Controle de Receita
          </h1>
          <p className="text-sm text-somus-gray-500 mt-1">
            Receita mensal por assessor e produto
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
            variant="secondary"
            icon={<Download className="h-4 w-4" />}
            onClick={handleExport}
          >
            Exportar Excel
          </Button>
        </div>
      </div>

      {/* Filters */}
      <Card padding="sm">
        <div className="flex flex-wrap items-center gap-4 p-2">
          <div className="flex items-center gap-2">
            <Filter className="h-4 w-4 text-somus-gray-400" />
            <span className="text-sm font-medium text-somus-gray-600">Filtros:</span>
          </div>
          <select
            value={mesFilter}
            onChange={(e) => setMesFilter(Number(e.target.value))}
            className="text-sm border border-somus-gray-300 rounded-lg px-3 py-1.5 focus:ring-2 focus:ring-somus-green/40 focus:outline-none"
          >
            {MESES.map((m, i) => (
              <option key={i} value={i}>{m}</option>
            ))}
          </select>
          <select
            value={anoFilter}
            onChange={(e) => setAnoFilter(Number(e.target.value))}
            className="text-sm border border-somus-gray-300 rounded-lg px-3 py-1.5 focus:ring-2 focus:ring-somus-green/40 focus:outline-none"
          >
            {anos.map((a) => (
              <option key={a} value={a}>{a}</option>
            ))}
          </select>
          <select
            value={equipeFilter}
            onChange={(e) => setEquipeFilter(e.target.value)}
            className="text-sm border border-somus-gray-300 rounded-lg px-3 py-1.5 focus:ring-2 focus:ring-somus-green/40 focus:outline-none"
          >
            <option value="TODAS">Todas as equipes</option>
            {EQUIPES.map((eq) => (
              <option key={eq} value={eq}>{eq}</option>
            ))}
          </select>
        </div>
      </Card>

      {/* Total KPI */}
      <div className="bg-white rounded-lg shadow-sm border border-somus-gray-200 p-5">
        <div className="flex items-center gap-3 mb-1">
          <div className="p-2 rounded-lg bg-somus-green/10 text-somus-green">
            <DollarSign className="h-5 w-5" />
          </div>
          <div>
            <div className="text-sm text-somus-gray-500">Receita Total - {MESES[mesFilter]} {anoFilter}</div>
            <div className="text-2xl font-bold text-somus-gray-900">{formatCurrency(totals.total)}</div>
          </div>
        </div>
      </div>

      {/* Pivot Table */}
      <Card title="Receita por Assessor e Produto">
        <div className="mt-4 overflow-x-auto">
          <table className="w-full text-sm">
            <thead>
              <tr className="border-b border-somus-gray-200">
                <th className="text-left py-3 px-3 font-semibold text-somus-gray-600 sticky left-0 bg-white min-w-[160px]">
                  Assessor
                </th>
                <th className="text-left py-3 px-2 font-semibold text-somus-gray-600 min-w-[80px]">
                  Equipe
                </th>
                {PRODUTOS.map((p) => (
                  <th
                    key={p}
                    className="text-right py-3 px-2 font-semibold text-somus-gray-600 min-w-[110px]"
                  >
                    {PRODUTO_LABELS[p]}
                  </th>
                ))}
                <th className="text-right py-3 px-3 font-semibold text-somus-green min-w-[120px]">
                  Total
                </th>
              </tr>
            </thead>
            <tbody>
              {filteredData.map((row) => (
                <tr
                  key={row.assessor}
                  className="border-b border-somus-gray-100 hover:bg-somus-gray-50 transition-colors"
                >
                  <td className="py-2.5 px-3 font-medium text-somus-gray-900 sticky left-0 bg-white">
                    {row.assessor}
                  </td>
                  <td className="py-2.5 px-2">
                    <span className="inline-block px-2 py-0.5 text-xs rounded-full bg-somus-green/10 text-somus-green font-medium">
                      {row.equipe}
                    </span>
                  </td>
                  {PRODUTOS.map((p) => (
                    <td key={p} className="py-2.5 px-2 text-right text-somus-gray-700">
                      {formatCurrency((row as any)[p])}
                    </td>
                  ))}
                  <td className="py-2.5 px-3 text-right font-bold text-somus-green">
                    {formatCurrency(row.total)}
                  </td>
                </tr>
              ))}

              {/* Subtotals by equipe */}
              {Object.entries(subtotalsByEquipe).map(([equipe, vals]) => (
                <tr key={`sub-${equipe}`} className="bg-somus-gray-50 border-b border-somus-gray-200">
                  <td className="py-2.5 px-3 font-bold text-somus-gray-700 sticky left-0 bg-somus-gray-50">
                    Subtotal {equipe}
                  </td>
                  <td className="py-2.5 px-2"></td>
                  {PRODUTOS.map((p) => (
                    <td key={p} className="py-2.5 px-2 text-right font-semibold text-somus-gray-700">
                      {formatCurrency(vals[p])}
                    </td>
                  ))}
                  <td className="py-2.5 px-3 text-right font-bold text-somus-green">
                    {formatCurrency(vals.total)}
                  </td>
                </tr>
              ))}

              {/* Grand Total */}
              <tr className="bg-somus-green/5 border-t-2 border-somus-green">
                <td className="py-3 px-3 font-bold text-somus-green sticky left-0 bg-somus-green/5">
                  TOTAL GERAL
                </td>
                <td className="py-3 px-2"></td>
                {PRODUTOS.map((p) => (
                  <td key={p} className="py-3 px-2 text-right font-bold text-somus-green">
                    {formatCurrency(totals[p])}
                  </td>
                ))}
                <td className="py-3 px-3 text-right font-bold text-somus-green text-base">
                  {formatCurrency(totals.total)}
                </td>
              </tr>
            </tbody>
          </table>
        </div>
      </Card>
    </div>
  );
}

import React, { useState, useMemo, useCallback } from 'react';
import { Card } from '@components/Card';
import { Button } from '@components/Button';
import {
  Upload,
  FileText,
  Filter,
  Download,
  Eye,
  Calendar,
  Search,
} from 'lucide-react';
import { cn, formatCurrency } from '@/lib/utils';
import { readExcelFromFile } from '@/services/excel-reader';
import { generateFluxoRFPDF } from '@/services/pdf-generator';

// ── Types ────────────────────────────────────────────────────────────────────

interface EventoRF {
  id: string;
  ativo: string;
  tipoEvento: string;
  data: string;
  assessor: string;
  valor: number;
  descricao?: string;
}

// ── Mock Data ────────────────────────────────────────────────────────────────

const MOCK_EVENTOS: EventoRF[] = [
  { id: '1', ativo: 'CDB Banco XP 2026', tipoEvento: 'Vencimento', data: '2026-04-15', assessor: 'Carlos Silva', valor: 250000 },
  { id: '2', ativo: 'LCI Itau 95% CDI', tipoEvento: 'Vencimento', data: '2026-04-20', assessor: 'Carlos Silva', valor: 180000 },
  { id: '3', ativo: 'Debenture VALE25', tipoEvento: 'Cupom', data: '2026-04-01', assessor: 'Ana Santos', valor: 45000 },
  { id: '4', ativo: 'CRA Suzano 2026', tipoEvento: 'Amortizacao', data: '2026-04-10', assessor: 'Ana Santos', valor: 120000 },
  { id: '5', ativo: 'LCA BB 97% CDI', tipoEvento: 'Vencimento', data: '2026-05-05', assessor: 'Pedro Costa', valor: 300000 },
  { id: '6', ativo: 'CDB Safra IPCA+6', tipoEvento: 'Vencimento', data: '2026-05-15', assessor: 'Pedro Costa', valor: 420000 },
  { id: '7', ativo: 'Debenture PETR26', tipoEvento: 'Cupom', data: '2026-04-25', assessor: 'Maria Oliveira', valor: 35000 },
  { id: '8', ativo: 'CRI Cyrela 2026', tipoEvento: 'Amortizacao', data: '2026-04-30', assessor: 'Joao Souza', valor: 90000 },
  { id: '9', ativo: 'Tesouro IPCA+ 2026', tipoEvento: 'Cupom', data: '2026-05-01', assessor: 'Fernanda Lima', valor: 15000 },
  { id: '10', ativo: 'LCI Bradesco 93% CDI', tipoEvento: 'Vencimento', data: '2026-05-20', assessor: 'Lucas Almeida', valor: 500000 },
];

const ASSESSORES_LIST = [...new Set(MOCK_EVENTOS.map((e) => e.assessor))];
const TIPOS_EVENTO = [...new Set(MOCK_EVENTOS.map((e) => e.tipoEvento))];

// ── Component ────────────────────────────────────────────────────────────────

export default function FluxoRF() {
  const [eventos, setEventos] = useState<EventoRF[]>(MOCK_EVENTOS);
  const [assessorFilter, setAssessorFilter] = useState('TODOS');
  const [tipoFilter, setTipoFilter] = useState('TODOS');
  const [dataInicio, setDataInicio] = useState('');
  const [dataFim, setDataFim] = useState('');
  const [searchTerm, setSearchTerm] = useState('');
  const [uploading, setUploading] = useState(false);
  const [generating, setGenerating] = useState(false);
  const [generatedPDFs, setGeneratedPDFs] = useState<string[]>([]);
  const [previewUrl, setPreviewUrl] = useState<string | null>(null);

  const filteredEventos = useMemo(() => {
    return eventos.filter((e) => {
      if (assessorFilter !== 'TODOS' && e.assessor !== assessorFilter) return false;
      if (tipoFilter !== 'TODOS' && e.tipoEvento !== tipoFilter) return false;
      if (dataInicio && e.data < dataInicio) return false;
      if (dataFim && e.data > dataFim) return false;
      if (searchTerm) {
        const term = searchTerm.toLowerCase();
        return (
          e.ativo.toLowerCase().includes(term) ||
          e.assessor.toLowerCase().includes(term)
        );
      }
      return true;
    });
  }, [eventos, assessorFilter, tipoFilter, dataInicio, dataFim, searchTerm]);

  const handleUpload = useCallback(async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setUploading(true);
    try {
      const data = await readExcelFromFile(file);
      const parsed: EventoRF[] = data.map((row, i) => ({
        id: String(i + 1),
        ativo: String(row['Ativo'] || row['ativo'] || ''),
        tipoEvento: String(row['Tipo Evento'] || row['tipoEvento'] || row['Tipo'] || ''),
        data: String(row['Data'] || row['data'] || ''),
        assessor: String(row['Assessor'] || row['assessor'] || ''),
        valor: Number(row['Valor'] || row['valor'] || 0),
        descricao: String(row['Descricao'] || row['descricao'] || ''),
      }));
      setEventos(parsed.length > 0 ? parsed : MOCK_EVENTOS);
    } catch (err) {
      console.error('Erro ao ler arquivo:', err);
    } finally {
      setUploading(false);
    }
  }, []);

  const handleGerarPDFs = useCallback(async () => {
    setGenerating(true);
    setGeneratedPDFs([]);
    try {
      const byAssessor: Record<string, EventoRF[]> = {};
      filteredEventos.forEach((e) => {
        if (!byAssessor[e.assessor]) byAssessor[e.assessor] = [];
        byAssessor[e.assessor].push(e);
      });

      const names: string[] = [];
      for (const [assessor, evts] of Object.entries(byAssessor)) {
        const blob = await generateFluxoRFPDF(
          assessor,
          evts.map((e) => ({
            ativo: e.ativo,
            tipoEvento: e.tipoEvento,
            data: e.data,
            valor: e.valor,
          }))
        );
        // Simula download/preview
        const url = URL.createObjectURL(blob);
        names.push(assessor);
        // Seta preview do ultimo gerado
        setPreviewUrl(url);
      }
      setGeneratedPDFs(names);
    } catch (err) {
      console.error('Erro ao gerar PDFs:', err);
    } finally {
      setGenerating(false);
    }
  }, [filteredEventos]);

  const formatDate = (dateStr: string) => {
    if (!dateStr) return '';
    const d = new Date(dateStr + 'T00:00:00');
    return d.toLocaleDateString('pt-BR');
  };

  return (
    <div className="space-y-6">
      {/* Header */}
      <div className="flex items-center justify-between">
        <div>
          <h1 className="text-2xl font-bold text-somus-text-primary">
            Fluxo Renda Fixa
          </h1>
          <p className="text-sm text-somus-text-secondary mt-1">
            Agenda de eventos de renda fixa por assessor
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
            icon={<FileText className="h-4 w-4" />}
            loading={generating}
            onClick={handleGerarPDFs}
            disabled={filteredEventos.length === 0}
          >
            Gerar PDFs
          </Button>
        </div>
      </div>

      {/* Filters */}
      <Card padding="sm">
        <div className="flex flex-wrap items-center gap-4 p-2">
          <div className="flex items-center gap-2">
            <Filter className="h-4 w-4 text-somus-text-tertiary" />
            <span className="text-sm font-medium text-somus-text-secondary">Filtros:</span>
          </div>

          <div className="relative">
            <Search className="absolute left-3 top-1/2 -translate-y-1/2 h-4 w-4 text-somus-text-tertiary" />
            <input
              type="text"
              placeholder="Buscar ativo ou assessor..."
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
              className="text-sm border border-somus-border rounded-lg pl-9 pr-3 py-1.5 w-56 bg-somus-bg-input text-somus-text-primary placeholder:text-somus-text-tertiary focus:ring-2 focus:ring-somus-green/40 focus:outline-none"
            />
          </div>

          <select
            value={assessorFilter}
            onChange={(e) => setAssessorFilter(e.target.value)}
            className="text-sm border border-somus-border rounded-lg px-3 py-1.5 bg-somus-bg-input text-somus-text-primary focus:ring-2 focus:ring-somus-green/40 focus:outline-none"
          >
            <option value="TODOS">Todos assessores</option>
            {ASSESSORES_LIST.map((a) => (
              <option key={a} value={a}>{a}</option>
            ))}
          </select>

          <select
            value={tipoFilter}
            onChange={(e) => setTipoFilter(e.target.value)}
            className="text-sm border border-somus-border rounded-lg px-3 py-1.5 bg-somus-bg-input text-somus-text-primary focus:ring-2 focus:ring-somus-green/40 focus:outline-none"
          >
            <option value="TODOS">Todos os tipos</option>
            {TIPOS_EVENTO.map((t) => (
              <option key={t} value={t}>{t}</option>
            ))}
          </select>

          <div className="flex items-center gap-2">
            <Calendar className="h-4 w-4 text-somus-text-tertiary" />
            <input
              type="date"
              value={dataInicio}
              onChange={(e) => setDataInicio(e.target.value)}
              className="text-sm border border-somus-border rounded-lg px-3 py-1.5 bg-somus-bg-input text-somus-text-primary focus:ring-2 focus:ring-somus-green/40 focus:outline-none"
            />
            <span className="text-somus-text-tertiary">a</span>
            <input
              type="date"
              value={dataFim}
              onChange={(e) => setDataFim(e.target.value)}
              className="text-sm border border-somus-border rounded-lg px-3 py-1.5 bg-somus-bg-input text-somus-text-primary focus:ring-2 focus:ring-somus-green/40 focus:outline-none"
            />
          </div>
        </div>
      </Card>

      {/* DataTable */}
      <Card title={`Eventos (${filteredEventos.length})`}>
        <div className="mt-4 overflow-x-auto">
          <table className="w-full text-sm">
            <thead>
              <tr className="border-b border-somus-border">
                <th className="text-left py-3 px-4 font-semibold text-somus-text-secondary">Ativo</th>
                <th className="text-left py-3 px-4 font-semibold text-somus-text-secondary">Tipo Evento</th>
                <th className="text-left py-3 px-4 font-semibold text-somus-text-secondary">Data</th>
                <th className="text-left py-3 px-4 font-semibold text-somus-text-secondary">Assessor</th>
                <th className="text-right py-3 px-4 font-semibold text-somus-text-secondary">Valor</th>
              </tr>
            </thead>
            <tbody>
              {filteredEventos.map((evt) => (
                <tr
                  key={evt.id}
                  className="border-b border-somus-border/30 hover:bg-somus-bg-hover transition-colors"
                >
                  <td className="py-3 px-4 font-medium text-somus-text-primary">{evt.ativo}</td>
                  <td className="py-3 px-4">
                    <span
                      className={cn(
                        'inline-block px-2 py-0.5 text-xs rounded-full font-medium',
                        evt.tipoEvento === 'Vencimento'
                          ? 'bg-red-500/10 text-red-400'
                          : evt.tipoEvento === 'Cupom'
                          ? 'bg-amber-500/10 text-amber-400'
                          : 'bg-blue-500/10 text-blue-400'
                      )}
                    >
                      {evt.tipoEvento}
                    </span>
                  </td>
                  <td className="py-3 px-4 text-somus-text-secondary">{formatDate(evt.data)}</td>
                  <td className="py-3 px-4 text-somus-text-secondary">{evt.assessor}</td>
                  <td className="py-3 px-4 text-right font-medium text-somus-text-primary">
                    {formatCurrency(evt.valor)}
                  </td>
                </tr>
              ))}
              {filteredEventos.length === 0 && (
                <tr>
                  <td colSpan={5} className="py-12 text-center text-somus-text-tertiary">
                    Nenhum evento encontrado
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </Card>

      {/* Generated PDFs */}
      {generatedPDFs.length > 0 && (
        <Card title="PDFs Gerados">
          <div className="mt-4 space-y-2">
            {generatedPDFs.map((name) => (
              <div
                key={name}
                className="flex items-center justify-between py-2 px-4 bg-somus-bg-hover rounded-lg"
              >
                <div className="flex items-center gap-3">
                  <FileText className="h-5 w-5 text-somus-green" />
                  <span className="text-sm font-medium text-somus-text-primary">
                    Fluxo RF - {name}.pdf
                  </span>
                </div>
                <div className="flex items-center gap-2">
                  <Button variant="ghost" size="sm" icon={<Eye className="h-4 w-4" />}>
                    Visualizar
                  </Button>
                  <Button variant="ghost" size="sm" icon={<Download className="h-4 w-4" />}>
                    Baixar
                  </Button>
                </div>
              </div>
            ))}
          </div>
        </Card>
      )}

      {/* PDF Preview */}
      {previewUrl && (
        <Card title="Preview do PDF">
          <div className="mt-4">
            <iframe
              src={previewUrl}
              className="w-full h-[600px] border border-somus-border rounded-lg"
              title="Preview PDF"
            />
          </div>
        </Card>
      )}
    </div>
  );
}

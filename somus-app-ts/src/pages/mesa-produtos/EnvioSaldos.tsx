import React, { useState, useMemo, useCallback } from 'react';
import { Card } from '@components/Card';
import { Button } from '@components/Button';
import {
  Upload,
  Send,
  Search,
  Filter,
  Eye,
  Mail,
  DollarSign,
  Users,
} from 'lucide-react';
import { cn, formatCurrency } from '@/lib/utils';
import { readExcelFromFile } from '@/services/excel-reader';
import { createDrafts } from '@/services/outlook';

// ── Types ────────────────────────────────────────────────────────────────────

interface SaldoItem {
  id: string;
  assessor: string;
  assessorEmail: string;
  cliente: string;
  saldoDisponivel: number;
  saldoAplicado: number;
  vencimentosRF: number;
  vencimentosFundos: number;
  dataRef: string;
}

// ── Mock Data ────────────────────────────────────────────────────────────────

const MOCK_SALDOS: SaldoItem[] = [
  { id: '1', assessor: 'Carlos Silva', assessorEmail: 'carlos@somus.com', cliente: 'Joao Mendes', saldoDisponivel: 85000, saldoAplicado: 520000, vencimentosRF: 250000, vencimentosFundos: 50000, dataRef: '2026-03-24' },
  { id: '2', assessor: 'Carlos Silva', assessorEmail: 'carlos@somus.com', cliente: 'Maria Costa', saldoDisponivel: 42000, saldoAplicado: 310000, vencimentosRF: 180000, vencimentosFundos: 0, dataRef: '2026-03-24' },
  { id: '3', assessor: 'Carlos Silva', assessorEmail: 'carlos@somus.com', cliente: 'Roberto Alves', saldoDisponivel: 120000, saldoAplicado: 750000, vencimentosRF: 0, vencimentosFundos: 120000, dataRef: '2026-03-24' },
  { id: '4', assessor: 'Ana Santos', assessorEmail: 'ana@somus.com', cliente: 'Pedro Lima', saldoDisponivel: 15000, saldoAplicado: 1200000, vencimentosRF: 420000, vencimentosFundos: 200000, dataRef: '2026-03-24' },
  { id: '5', assessor: 'Ana Santos', assessorEmail: 'ana@somus.com', cliente: 'Fernanda Dias', saldoDisponivel: 230000, saldoAplicado: 450000, vencimentosRF: 100000, vencimentosFundos: 0, dataRef: '2026-03-24' },
  { id: '6', assessor: 'Pedro Costa', assessorEmail: 'pedro@somus.com', cliente: 'Ana Souza', saldoDisponivel: 5000, saldoAplicado: 280000, vencimentosRF: 90000, vencimentosFundos: 30000, dataRef: '2026-03-24' },
  { id: '7', assessor: 'Pedro Costa', assessorEmail: 'pedro@somus.com', cliente: 'Lucas Ferreira', saldoDisponivel: 67000, saldoAplicado: 900000, vencimentosRF: 0, vencimentosFundos: 0, dataRef: '2026-03-24' },
  { id: '8', assessor: 'Maria Oliveira', assessorEmail: 'maria@somus.com', cliente: 'Bruno Neves', saldoDisponivel: 340000, saldoAplicado: 1500000, vencimentosRF: 600000, vencimentosFundos: 150000, dataRef: '2026-03-24' },
];

const ASSESSORES = [...new Set(MOCK_SALDOS.map((s) => s.assessor))];

// ── Component ────────────────────────────────────────────────────────────────

export default function EnvioSaldos() {
  const [saldos, setSaldos] = useState<SaldoItem[]>(MOCK_SALDOS);
  const [searchTerm, setSearchTerm] = useState('');
  const [assessorFilter, setAssessorFilter] = useState('TODOS');
  const [uploading, setUploading] = useState(false);
  const [generating, setGenerating] = useState(false);
  const [generated, setGenerated] = useState(false);
  const [previewAssessor, setPreviewAssessor] = useState<string | null>(null);

  const filteredSaldos = useMemo(() => {
    return saldos.filter((s) => {
      if (assessorFilter !== 'TODOS' && s.assessor !== assessorFilter) return false;
      if (searchTerm) {
        const term = searchTerm.toLowerCase();
        return (
          s.cliente.toLowerCase().includes(term) ||
          s.assessor.toLowerCase().includes(term)
        );
      }
      return true;
    });
  }, [saldos, searchTerm, assessorFilter]);

  const totalDisponivel = filteredSaldos.reduce((s, r) => s + r.saldoDisponivel, 0);
  const totalAplicado = filteredSaldos.reduce((s, r) => s + r.saldoAplicado, 0);

  const handleUpload = useCallback(async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setUploading(true);
    try {
      const rows = await readExcelFromFile(file);
      const parsed: SaldoItem[] = rows.map((row, i) => ({
        id: String(i + 1),
        assessor: String(row['Assessor'] || ''),
        assessorEmail: String(row['Email'] || `${String(row['Assessor'] || '').toLowerCase().replace(/\s/g, '.')}@somus.com`),
        cliente: String(row['Cliente'] || ''),
        saldoDisponivel: Number(row['Saldo Disponivel'] || 0),
        saldoAplicado: Number(row['Saldo Aplicado'] || 0),
        vencimentosRF: Number(row['Vencimentos RF'] || 0),
        vencimentosFundos: Number(row['Vencimentos Fundos'] || 0),
        dataRef: String(row['Data Ref'] || new Date().toISOString().split('T')[0]),
      }));
      if (parsed.length > 0) setSaldos(parsed);
    } catch (err) {
      console.error('Erro ao ler arquivo:', err);
    } finally {
      setUploading(false);
    }
  }, []);

  const handleGerarRascunhos = useCallback(async () => {
    setGenerating(true);
    try {
      const byAssessor: Record<string, SaldoItem[]> = {};
      filteredSaldos.forEach((s) => {
        if (!byAssessor[s.assessor]) byAssessor[s.assessor] = [];
        byAssessor[s.assessor].push(s);
      });

      const emails = Object.entries(byAssessor).map(([assessor, items]) => {
        const email = items[0].assessorEmail;
        return {
          to: email,
          subject: `Saldos e Vencimentos - ${new Date().toLocaleDateString('pt-BR')}`,
          body: `<div style="font-family: DM Sans, sans-serif; max-width: 700px;">
            <div style="background: #004D33; color: white; padding: 20px; border-radius: 8px 8px 0 0;">
              <h2 style="margin: 0;">Saldos e Vencimentos</h2>
              <p style="margin: 5px 0 0; opacity: 0.8; font-size: 13px;">${new Date().toLocaleDateString('pt-BR')}</p>
            </div>
            <div style="padding: 20px; border: 1px solid #E5E7EB; border-top: none; border-radius: 0 0 8px 8px;">
              <p>Ola ${assessor.split(' ')[0]},</p>
              <p>Segue o resumo dos saldos e vencimentos dos seus clientes:</p>
              <table style="width: 100%; border-collapse: collapse; font-size: 13px; margin-top: 15px;">
                <tr style="background: #F0F0E0;">
                  <th style="padding: 10px; text-align: left; border: 1px solid #E5E7EB;">Cliente</th>
                  <th style="padding: 10px; text-align: right; border: 1px solid #E5E7EB;">Saldo Disponivel</th>
                  <th style="padding: 10px; text-align: right; border: 1px solid #E5E7EB;">Venc. RF</th>
                  <th style="padding: 10px; text-align: right; border: 1px solid #E5E7EB;">Venc. Fundos</th>
                </tr>
                ${items.map((item) => `
                  <tr>
                    <td style="padding: 8px; border: 1px solid #E5E7EB;">${item.cliente}</td>
                    <td style="padding: 8px; text-align: right; border: 1px solid #E5E7EB; color: ${item.saldoDisponivel > 50000 ? '#059669' : '#111'};">${formatCurrency(item.saldoDisponivel)}</td>
                    <td style="padding: 8px; text-align: right; border: 1px solid #E5E7EB;">${item.vencimentosRF > 0 ? formatCurrency(item.vencimentosRF) : '-'}</td>
                    <td style="padding: 8px; text-align: right; border: 1px solid #E5E7EB;">${item.vencimentosFundos > 0 ? formatCurrency(item.vencimentosFundos) : '-'}</td>
                  </tr>
                `).join('')}
              </table>
              <p style="color: #888; font-size: 12px; margin-top: 20px;">Somus Capital - Mesa de Produtos</p>
            </div>
          </div>`,
        };
      });

      await createDrafts(emails);
      setGenerated(true);
      setTimeout(() => setGenerated(false), 4000);
    } catch (err) {
      console.error('Erro ao gerar rascunhos:', err);
    } finally {
      setGenerating(false);
    }
  }, [filteredSaldos]);

  const previewItems = previewAssessor
    ? filteredSaldos.filter((s) => s.assessor === previewAssessor)
    : [];

  return (
    <div className="space-y-6">
      {/* Header */}
      <div className="flex items-center justify-between">
        <div>
          <h1 className="text-2xl font-bold text-somus-text-primary">Envio de Saldos</h1>
          <p className="text-sm text-somus-text-secondary mt-1">
            Envio diario de saldos e vencimentos para assessores
          </p>
        </div>
        <div className="flex items-center gap-3">
          <label className="cursor-pointer">
            <input type="file" accept=".xlsx,.xls" onChange={handleUpload} className="hidden" />
            <Button variant="secondary" icon={<Upload className="h-4 w-4" />} loading={uploading} onClick={() => {}}>
              Upload Base
            </Button>
          </label>
          <Button
            variant="primary"
            icon={<Mail className="h-4 w-4" />}
            loading={generating}
            onClick={handleGerarRascunhos}
            disabled={filteredSaldos.length === 0}
          >
            {generated ? 'Rascunhos Criados!' : 'Gerar Rascunhos Outlook'}
          </Button>
        </div>
      </div>

      {/* KPIs */}
      <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
        <div className="bg-somus-bg-secondary rounded-lg border border-somus-border p-5">
          <div className="flex items-center gap-2 mb-2">
            <Users className="h-4 w-4 text-somus-green" />
            <span className="text-sm text-somus-text-secondary">Clientes</span>
          </div>
          <div className="text-2xl font-bold text-somus-text-primary">{filteredSaldos.length}</div>
        </div>
        <div className="bg-somus-bg-secondary rounded-lg border border-somus-border p-5">
          <div className="flex items-center gap-2 mb-2">
            <DollarSign className="h-4 w-4 text-somus-green" />
            <span className="text-sm text-somus-text-secondary">Total Disponivel</span>
          </div>
          <div className="text-2xl font-bold text-somus-text-primary">{formatCurrency(totalDisponivel)}</div>
        </div>
        <div className="bg-somus-bg-secondary rounded-lg border border-somus-border p-5">
          <div className="flex items-center gap-2 mb-2">
            <DollarSign className="h-4 w-4 text-somus-green" />
            <span className="text-sm text-somus-text-secondary">Total Aplicado</span>
          </div>
          <div className="text-2xl font-bold text-somus-text-primary">{formatCurrency(totalAplicado)}</div>
        </div>
      </div>

      {/* Filters */}
      <Card padding="sm">
        <div className="flex flex-wrap items-center gap-4 p-2">
          <div className="relative">
            <Search className="absolute left-3 top-1/2 -translate-y-1/2 h-4 w-4 text-somus-text-tertiary" />
            <input
              type="text"
              placeholder="Buscar assessor ou cliente..."
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
            {ASSESSORES.map((a) => (
              <option key={a} value={a}>{a}</option>
            ))}
          </select>
        </div>
      </Card>

      {/* Table */}
      <Card title={`Saldos (${filteredSaldos.length})`}>
        <div className="mt-4 overflow-x-auto">
          <table className="w-full text-sm">
            <thead>
              <tr className="border-b border-somus-border">
                <th className="text-left py-3 px-4 font-semibold text-somus-text-secondary">Assessor</th>
                <th className="text-left py-3 px-4 font-semibold text-somus-text-secondary">Cliente</th>
                <th className="text-right py-3 px-4 font-semibold text-somus-text-secondary">Saldo Disponivel</th>
                <th className="text-right py-3 px-4 font-semibold text-somus-text-secondary">Saldo Aplicado</th>
                <th className="text-right py-3 px-4 font-semibold text-somus-text-secondary">Venc. RF</th>
                <th className="text-right py-3 px-4 font-semibold text-somus-text-secondary">Venc. Fundos</th>
                <th className="py-3 px-4 w-10"></th>
              </tr>
            </thead>
            <tbody>
              {filteredSaldos.map((s) => (
                <tr key={s.id} className="border-b border-somus-border/30 hover:bg-somus-bg-hover transition-colors">
                  <td className="py-3 px-4 text-somus-text-secondary">{s.assessor}</td>
                  <td className="py-3 px-4 font-medium text-somus-text-primary">{s.cliente}</td>
                  <td className={cn('py-3 px-4 text-right font-medium', s.saldoDisponivel > 50000 ? 'text-somus-green-400' : 'text-somus-text-primary')}>
                    {formatCurrency(s.saldoDisponivel)}
                  </td>
                  <td className="py-3 px-4 text-right text-somus-text-primary">{formatCurrency(s.saldoAplicado)}</td>
                  <td className="py-3 px-4 text-right text-somus-text-secondary">
                    {s.vencimentosRF > 0 ? formatCurrency(s.vencimentosRF) : '-'}
                  </td>
                  <td className="py-3 px-4 text-right text-somus-text-secondary">
                    {s.vencimentosFundos > 0 ? formatCurrency(s.vencimentosFundos) : '-'}
                  </td>
                  <td className="py-3 px-4">
                    <button
                      onClick={() => setPreviewAssessor(s.assessor)}
                      className="text-somus-text-tertiary hover:text-somus-green transition-colors"
                    >
                      <Eye className="h-4 w-4" />
                    </button>
                  </td>
                </tr>
              ))}
              {filteredSaldos.length === 0 && (
                <tr>
                  <td colSpan={7} className="py-12 text-center text-somus-text-tertiary">
                    Nenhum saldo encontrado
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </Card>

      {/* Email Preview */}
      {previewAssessor && previewItems.length > 0 && (
        <Card
          title={`Preview Email - ${previewAssessor}`}
          headerRight={
            <button
              onClick={() => setPreviewAssessor(null)}
              className="text-sm text-somus-text-tertiary hover:text-somus-text-secondary"
            >
              Fechar
            </button>
          }
        >
          <div className="mt-4 border border-somus-border rounded-lg overflow-hidden">
            <div className="bg-somus-green text-white p-5">
              <h3 className="text-lg font-semibold">Saldos e Vencimentos</h3>
              <p className="text-sm opacity-80 mt-1">{new Date().toLocaleDateString('pt-BR')}</p>
            </div>
            <div className="p-5">
              <p className="text-sm text-somus-text-primary mb-4">
                Ola {previewAssessor.split(' ')[0]}, segue o resumo dos saldos e vencimentos:
              </p>
              <table className="w-full text-sm border-collapse">
                <thead>
                  <tr className="bg-somus-border/30">
                    <th className="text-left py-2 px-3 border border-somus-border text-somus-text-secondary">Cliente</th>
                    <th className="text-right py-2 px-3 border border-somus-border text-somus-text-secondary">Saldo Disponivel</th>
                    <th className="text-right py-2 px-3 border border-somus-border text-somus-text-secondary">Venc. RF</th>
                    <th className="text-right py-2 px-3 border border-somus-border text-somus-text-secondary">Venc. Fundos</th>
                  </tr>
                </thead>
                <tbody>
                  {previewItems.map((item) => (
                    <tr key={item.id}>
                      <td className="py-2 px-3 border border-somus-border">{item.cliente}</td>
                      <td className={cn('py-2 px-3 border border-somus-border text-right', item.saldoDisponivel > 50000 && 'text-somus-green-400 font-medium')}>
                        {formatCurrency(item.saldoDisponivel)}
                      </td>
                      <td className="py-2 px-3 border border-somus-border text-right">
                        {item.vencimentosRF > 0 ? formatCurrency(item.vencimentosRF) : '-'}
                      </td>
                      <td className="py-2 px-3 border border-somus-border text-right">
                        {item.vencimentosFundos > 0 ? formatCurrency(item.vencimentosFundos) : '-'}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
              <p className="text-xs text-somus-text-tertiary mt-4">Somus Capital - Mesa de Produtos</p>
            </div>
          </div>
        </Card>
      )}
    </div>
  );
}

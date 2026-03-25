import React, { useState, useMemo, useCallback } from 'react';
import { Card } from '@components/Card';
import { Button } from '@components/Button';
import {
  Upload,
  Mail,
  Search,
  Filter,
  Shield,
  Calendar,
  DollarSign,
  AlertTriangle,
  Check,
  Clock,
} from 'lucide-react';
import { cn, formatCurrency } from '@/lib/utils';
import { readExcelFromFile } from '@/services/excel-reader';
import { createDrafts } from '@/services/outlook';

// ── Types ────────────────────────────────────────────────────────────────────

type RenovacaoStatus = 'proxima' | 'vencida' | 'renovada';

interface RenovacaoItem {
  id: string;
  cliente: string;
  assessor: string;
  seguradora: string;
  produto: string;
  vencimento: string;
  premio: number;
  status: RenovacaoStatus;
}

// ── Mock Data ────────────────────────────────────────────────────────────────

const MOCK_RENOVACOES: RenovacaoItem[] = [
  { id: '1', cliente: 'Joao Mendes', assessor: 'Carlos Silva', seguradora: 'SulAmerica', produto: 'Vida Individual', vencimento: '2026-04-10', premio: 4500, status: 'proxima' },
  { id: '2', cliente: 'Maria Costa', assessor: 'Carlos Silva', seguradora: 'Porto Seguro', produto: 'Residencial', vencimento: '2026-03-15', premio: 2800, status: 'vencida' },
  { id: '3', cliente: 'Pedro Lima', assessor: 'Ana Santos', seguradora: 'Bradesco Seguros', produto: 'Auto', vencimento: '2026-04-22', premio: 6200, status: 'proxima' },
  { id: '4', cliente: 'Ana Souza', assessor: 'Pedro Costa', seguradora: 'Allianz', produto: 'Vida', vencimento: '2026-02-28', premio: 3100, status: 'renovada' },
  { id: '5', cliente: 'Lucas Ferreira', assessor: 'Maria Oliveira', seguradora: 'Zurich', produto: 'Empresarial', vencimento: '2026-04-05', premio: 15000, status: 'proxima' },
  { id: '6', cliente: 'Fernanda Dias', assessor: 'Carlos Silva', seguradora: 'Tokio Marine', produto: 'Auto', vencimento: '2026-03-20', premio: 5400, status: 'vencida' },
  { id: '7', cliente: 'Bruno Neves', assessor: 'Ana Santos', seguradora: 'Liberty', produto: 'Residencial', vencimento: '2026-05-12', premio: 1900, status: 'proxima' },
  { id: '8', cliente: 'Camila Rocha', assessor: 'Pedro Costa', seguradora: 'SulAmerica', produto: 'Saude', vencimento: '2026-01-15', premio: 12000, status: 'renovada' },
  { id: '9', cliente: 'Ricardo Gomes', assessor: 'Joao Souza', seguradora: 'HDI', produto: 'Auto', vencimento: '2026-04-18', premio: 4800, status: 'proxima' },
  { id: '10', cliente: 'Tatiana Mello', assessor: 'Fernanda Lima', seguradora: 'Mapfre', produto: 'Vida', vencimento: '2026-03-08', premio: 2200, status: 'vencida' },
  { id: '11', cliente: 'Roberto Alves', assessor: 'Carlos Silva', seguradora: 'Porto Seguro', produto: 'Empresarial', vencimento: '2026-04-30', premio: 22000, status: 'proxima' },
  { id: '12', cliente: 'Patricia Santos', assessor: 'Ana Santos', seguradora: 'Bradesco Seguros', produto: 'Residencial', vencimento: '2026-02-10', premio: 3500, status: 'renovada' },
];

const ASSESSORES = [...new Set(MOCK_RENOVACOES.map((r) => r.assessor))];
const MESES_FILTRO = [
  { value: 0, label: 'Janeiro' }, { value: 1, label: 'Fevereiro' }, { value: 2, label: 'Marco' },
  { value: 3, label: 'Abril' }, { value: 4, label: 'Maio' }, { value: 5, label: 'Junho' },
  { value: 6, label: 'Julho' }, { value: 7, label: 'Agosto' }, { value: 8, label: 'Setembro' },
  { value: 9, label: 'Outubro' }, { value: 10, label: 'Novembro' }, { value: 11, label: 'Dezembro' },
];

const STATUS_CONFIG: Record<RenovacaoStatus, { bg: string; text: string; label: string; icon: React.ReactNode }> = {
  proxima: {
    bg: 'bg-amber-500/10',
    text: 'text-amber-400',
    label: 'Proxima',
    icon: <Clock className="h-3.5 w-3.5" />,
  },
  vencida: {
    bg: 'bg-red-500/10',
    text: 'text-red-400',
    label: 'Vencida',
    icon: <AlertTriangle className="h-3.5 w-3.5" />,
  },
  renovada: {
    bg: 'bg-somus-green-500/10',
    text: 'text-somus-green-400',
    label: 'Renovada',
    icon: <Check className="h-3.5 w-3.5" />,
  },
};

// ── Component ────────────────────────────────────────────────────────────────

export default function RenovacoesAnuais() {
  const [renovacoes, setRenovacoes] = useState<RenovacaoItem[]>(MOCK_RENOVACOES);
  const [searchTerm, setSearchTerm] = useState('');
  const [assessorFilter, setAssessorFilter] = useState('TODOS');
  const [mesFilter, setMesFilter] = useState<number | 'TODOS'>('TODOS');
  const [statusFilter, setStatusFilter] = useState<RenovacaoStatus | 'TODOS'>('TODOS');
  const [uploading, setUploading] = useState(false);
  const [generating, setGenerating] = useState(false);
  const [generated, setGenerated] = useState(false);

  const filteredRenovacoes = useMemo(() => {
    return renovacoes.filter((r) => {
      if (assessorFilter !== 'TODOS' && r.assessor !== assessorFilter) return false;
      if (statusFilter !== 'TODOS' && r.status !== statusFilter) return false;
      if (mesFilter !== 'TODOS') {
        const vencMes = new Date(r.vencimento + 'T00:00:00').getMonth();
        if (vencMes !== mesFilter) return false;
      }
      if (searchTerm) {
        const term = searchTerm.toLowerCase();
        return (
          r.cliente.toLowerCase().includes(term) ||
          r.seguradora.toLowerCase().includes(term) ||
          r.produto.toLowerCase().includes(term) ||
          r.assessor.toLowerCase().includes(term)
        );
      }
      return true;
    });
  }, [renovacoes, searchTerm, assessorFilter, mesFilter, statusFilter]);

  // KPIs
  const totalRenovacoes = filteredRenovacoes.length;
  const totalPremios = filteredRenovacoes.reduce((s, r) => s + r.premio, 0);
  const vencendoEsteMes = filteredRenovacoes.filter((r) => {
    const vencMes = new Date(r.vencimento + 'T00:00:00').getMonth();
    return vencMes === new Date().getMonth() && r.status !== 'renovada';
  }).length;

  const handleUpload = useCallback(async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setUploading(true);
    try {
      const rows = await readExcelFromFile(file);
      const parsed: RenovacaoItem[] = rows.map((row, i) => {
        const vencStr = String(row['Vencimento'] || row['vencimento'] || '');
        const today = new Date();
        const vencDate = new Date(vencStr + 'T00:00:00');
        let status: RenovacaoStatus = 'proxima';
        if (row['Status'] === 'renovada' || row['status'] === 'renovada') {
          status = 'renovada';
        } else if (vencDate < today) {
          status = 'vencida';
        }

        return {
          id: String(i + 1),
          cliente: String(row['Cliente'] || ''),
          assessor: String(row['Assessor'] || ''),
          seguradora: String(row['Seguradora'] || ''),
          produto: String(row['Produto'] || ''),
          vencimento: vencStr,
          premio: Number(row['Premio'] || row['Prêmio'] || 0),
          status,
        };
      });
      if (parsed.length > 0) setRenovacoes(parsed);
    } catch (err) {
      console.error('Erro ao ler arquivo:', err);
    } finally {
      setUploading(false);
    }
  }, []);

  const handleGerarRascunhos = useCallback(async () => {
    setGenerating(true);
    try {
      // Group by assessor
      const byAssessor: Record<string, RenovacaoItem[]> = {};
      filteredRenovacoes
        .filter((r) => r.status !== 'renovada')
        .forEach((r) => {
          if (!byAssessor[r.assessor]) byAssessor[r.assessor] = [];
          byAssessor[r.assessor].push(r);
        });

      const emails = Object.entries(byAssessor).map(([assessor, items]) => ({
        to: `${assessor.toLowerCase().replace(/\s/g, '.')}@somus.com`,
        subject: `Renovacoes de Seguros Pendentes - ${new Date().toLocaleDateString('pt-BR')}`,
        body: `<div style="font-family: DM Sans, sans-serif; max-width: 700px;">
          <div style="background: #004D33; color: white; padding: 20px; border-radius: 8px 8px 0 0;">
            <h2 style="margin: 0;">Renovacoes de Seguros Pendentes</h2>
            <p style="margin: 5px 0 0; opacity: 0.8; font-size: 13px;">${assessor} | ${new Date().toLocaleDateString('pt-BR')}</p>
          </div>
          <div style="padding: 20px; border: 1px solid #E5E7EB; border-top: none; border-radius: 0 0 8px 8px;">
            <p>Ola ${assessor.split(' ')[0]},</p>
            <p>Seguem as renovacoes de seguros que necessitam de atencao:</p>
            <table style="width: 100%; border-collapse: collapse; font-size: 13px; margin-top: 15px;">
              <tr style="background: #F0F0E0;">
                <th style="padding: 10px; text-align: left; border: 1px solid #E5E7EB;">Cliente</th>
                <th style="padding: 10px; text-align: left; border: 1px solid #E5E7EB;">Seguradora</th>
                <th style="padding: 10px; text-align: left; border: 1px solid #E5E7EB;">Produto</th>
                <th style="padding: 10px; text-align: left; border: 1px solid #E5E7EB;">Vencimento</th>
                <th style="padding: 10px; text-align: right; border: 1px solid #E5E7EB;">Premio</th>
                <th style="padding: 10px; text-align: center; border: 1px solid #E5E7EB;">Status</th>
              </tr>
              ${items.map((r) => {
                const statusColor = r.status === 'vencida' ? '#DC2626' : '#D97706';
                const statusText = r.status === 'vencida' ? 'VENCIDA' : 'PROXIMA';
                return `
                  <tr>
                    <td style="padding: 8px; border: 1px solid #E5E7EB;">${r.cliente}</td>
                    <td style="padding: 8px; border: 1px solid #E5E7EB;">${r.seguradora}</td>
                    <td style="padding: 8px; border: 1px solid #E5E7EB;">${r.produto}</td>
                    <td style="padding: 8px; border: 1px solid #E5E7EB;">${new Date(r.vencimento + 'T00:00:00').toLocaleDateString('pt-BR')}</td>
                    <td style="padding: 8px; text-align: right; border: 1px solid #E5E7EB;">${formatCurrency(r.premio)}</td>
                    <td style="padding: 8px; text-align: center; border: 1px solid #E5E7EB; color: ${statusColor}; font-weight: 600;">${statusText}</td>
                  </tr>
                `;
              }).join('')}
            </table>
            <p style="color: #888; font-size: 12px; margin-top: 20px;">Somus Capital - Seguros</p>
          </div>
        </div>`,
      }));

      if (emails.length > 0) {
        await createDrafts(emails);
      }
      setGenerated(true);
      setTimeout(() => setGenerated(false), 4000);
    } catch (err) {
      console.error('Erro ao gerar rascunhos:', err);
    } finally {
      setGenerating(false);
    }
  }, [filteredRenovacoes]);

  const formatDate = (dateStr: string) => {
    if (!dateStr) return '';
    return new Date(dateStr + 'T00:00:00').toLocaleDateString('pt-BR');
  };

  return (
    <div className="space-y-6">
      {/* Header */}
      <div className="flex items-center justify-between">
        <div>
          <h1 className="text-2xl font-bold text-somus-text-primary">
            Renovacoes Anuais de Seguros
          </h1>
          <p className="text-sm text-somus-text-secondary mt-1">
            Acompanhamento e envio de renovacoes
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
            icon={<Mail className="h-4 w-4" />}
            loading={generating}
            onClick={handleGerarRascunhos}
            disabled={filteredRenovacoes.filter((r) => r.status !== 'renovada').length === 0}
          >
            {generated ? 'Rascunhos Criados!' : 'Gerar Rascunhos'}
          </Button>
        </div>
      </div>

      {/* KPIs */}
      <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
        <div className="bg-somus-bg-secondary rounded-lg border border-somus-border p-5">
          <div className="flex items-center gap-2 mb-2">
            <Shield className="h-5 w-5 text-somus-green" />
            <span className="text-sm text-somus-text-secondary">Total Renovacoes</span>
          </div>
          <div className="text-2xl font-bold text-somus-text-primary">{totalRenovacoes}</div>
        </div>
        <div className="bg-somus-bg-secondary rounded-lg border border-somus-border p-5">
          <div className="flex items-center gap-2 mb-2">
            <DollarSign className="h-5 w-5 text-somus-green" />
            <span className="text-sm text-somus-text-secondary">Total Premios</span>
          </div>
          <div className="text-2xl font-bold text-somus-text-primary">{formatCurrency(totalPremios)}</div>
        </div>
        <div className="bg-somus-bg-secondary rounded-lg border border-somus-border p-5">
          <div className="flex items-center gap-2 mb-2">
            <Calendar className="h-5 w-5 text-amber-400" />
            <span className="text-sm text-somus-text-secondary">Vencendo Este Mes</span>
          </div>
          <div className="text-2xl font-bold text-amber-400">{vencendoEsteMes}</div>
        </div>
      </div>

      {/* Filters */}
      <Card padding="sm">
        <div className="flex flex-wrap items-center gap-4 p-2">
          <div className="flex items-center gap-2">
            <Filter className="h-4 w-4 text-somus-text-tertiary" />
          </div>
          <div className="relative">
            <Search className="absolute left-3 top-1/2 -translate-y-1/2 h-4 w-4 text-somus-text-tertiary" />
            <input
              type="text"
              placeholder="Buscar..."
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
          <select
            value={mesFilter === 'TODOS' ? 'TODOS' : String(mesFilter)}
            onChange={(e) => setMesFilter(e.target.value === 'TODOS' ? 'TODOS' : Number(e.target.value))}
            className="text-sm border border-somus-border rounded-lg px-3 py-1.5 bg-somus-bg-input text-somus-text-primary focus:ring-2 focus:ring-somus-green/40 focus:outline-none"
          >
            <option value="TODOS">Todos os meses</option>
            {MESES_FILTRO.map((m) => (
              <option key={m.value} value={m.value}>{m.label}</option>
            ))}
          </select>
          <select
            value={statusFilter}
            onChange={(e) => setStatusFilter(e.target.value as any)}
            className="text-sm border border-somus-border rounded-lg px-3 py-1.5 bg-somus-bg-input text-somus-text-primary focus:ring-2 focus:ring-somus-green/40 focus:outline-none"
          >
            <option value="TODOS">Todos os status</option>
            <option value="proxima">Proxima</option>
            <option value="vencida">Vencida</option>
            <option value="renovada">Renovada</option>
          </select>
        </div>
      </Card>

      {/* Table */}
      <Card title={`Renovacoes (${filteredRenovacoes.length})`}>
        <div className="mt-4 overflow-x-auto">
          <table className="w-full text-sm">
            <thead>
              <tr className="border-b border-somus-border">
                <th className="text-left py-3 px-4 font-semibold text-somus-text-secondary">Cliente</th>
                <th className="text-left py-3 px-4 font-semibold text-somus-text-secondary">Assessor</th>
                <th className="text-left py-3 px-4 font-semibold text-somus-text-secondary">Seguradora</th>
                <th className="text-left py-3 px-4 font-semibold text-somus-text-secondary">Produto</th>
                <th className="text-left py-3 px-4 font-semibold text-somus-text-secondary">Vencimento</th>
                <th className="text-right py-3 px-4 font-semibold text-somus-text-secondary">Premio</th>
                <th className="text-center py-3 px-4 font-semibold text-somus-text-secondary">Status</th>
              </tr>
            </thead>
            <tbody>
              {filteredRenovacoes.map((r) => {
                const statusConf = STATUS_CONFIG[r.status];
                return (
                  <tr
                    key={r.id}
                    className={cn(
                      'border-b border-somus-border/30 transition-colors',
                      r.status === 'vencida' ? 'bg-red-500/5 hover:bg-red-500/10' : 'hover:bg-somus-bg-hover'
                    )}
                  >
                    <td className="py-3 px-4 font-medium text-somus-text-primary">{r.cliente}</td>
                    <td className="py-3 px-4 text-somus-text-secondary">{r.assessor}</td>
                    <td className="py-3 px-4 text-somus-text-primary">{r.seguradora}</td>
                    <td className="py-3 px-4">
                      <span className="inline-block px-2 py-0.5 text-xs rounded-full bg-somus-border/30 text-somus-text-primary font-medium">
                        {r.produto}
                      </span>
                    </td>
                    <td className="py-3 px-4 text-somus-text-secondary">{formatDate(r.vencimento)}</td>
                    <td className="py-3 px-4 text-right font-medium text-somus-text-primary">
                      {formatCurrency(r.premio)}
                    </td>
                    <td className="py-3 px-4 text-center">
                      <span
                        className={cn(
                          'inline-flex items-center gap-1 px-2.5 py-1 text-xs rounded-full font-semibold',
                          statusConf.bg,
                          statusConf.text
                        )}
                      >
                        {statusConf.icon}
                        {statusConf.label}
                      </span>
                    </td>
                  </tr>
                );
              })}
              {filteredRenovacoes.length === 0 && (
                <tr>
                  <td colSpan={7} className="py-12 text-center text-somus-text-tertiary">
                    Nenhuma renovacao encontrada
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

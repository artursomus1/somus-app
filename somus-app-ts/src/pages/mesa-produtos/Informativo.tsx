import React, { useState, useCallback } from 'react';
import { Card } from '@components/Card';
import { Button } from '@components/Button';
import {
  Send,
  Users,
  FileText,
  Clock,
  ChevronDown,
  Check,
  Mail,
} from 'lucide-react';
import { cn } from '@/lib/utils';
import { createDrafts } from '@/services/outlook';

// ── Types ────────────────────────────────────────────────────────────────────

interface Destinatario {
  nome: string;
  email: string;
  equipe: string;
  selecionado: boolean;
}

interface HistoricoItem {
  id: string;
  titulo: string;
  data: string;
  destinatarios: number;
  template: string;
}

// ── Mock Data ────────────────────────────────────────────────────────────────

const TEMPLATES = [
  { id: 'mercado', nome: 'Informativo de Mercado' },
  { id: 'produto', nome: 'Novo Produto' },
  { id: 'oportunidade', nome: 'Oportunidade de Investimento' },
  { id: 'aviso', nome: 'Aviso Importante' },
  { id: 'custom', nome: 'Personalizado' },
];

const MOCK_DESTINATARIOS: Destinatario[] = [
  { nome: 'Carlos Silva', email: 'carlos@somus.com', equipe: 'SP', selecionado: false },
  { nome: 'Ana Santos', email: 'ana@somus.com', equipe: 'SP', selecionado: false },
  { nome: 'Pedro Costa', email: 'pedro@somus.com', equipe: 'LEBLON', selecionado: false },
  { nome: 'Maria Oliveira', email: 'maria@somus.com', equipe: 'LEBLON', selecionado: false },
  { nome: 'Joao Souza', email: 'joao@somus.com', equipe: 'PRODUTOS', selecionado: false },
  { nome: 'Fernanda Lima', email: 'fernanda@somus.com', equipe: 'CORPORATE', selecionado: false },
  { nome: 'Lucas Almeida', email: 'lucas@somus.com', equipe: 'CORPORATE', selecionado: false },
  { nome: 'Juliana Rocha', email: 'juliana@somus.com', equipe: 'BACKOFFICE', selecionado: false },
];

const MOCK_HISTORICO: HistoricoItem[] = [
  { id: '1', titulo: 'Oportunidade CDB IPCA+7%', data: '2026-03-20', destinatarios: 12, template: 'Oportunidade de Investimento' },
  { id: '2', titulo: 'Novo Fundo Multimercado', data: '2026-03-18', destinatarios: 25, template: 'Novo Produto' },
  { id: '3', titulo: 'Cenario Macro Semanal', data: '2026-03-14', destinatarios: 30, template: 'Informativo de Mercado' },
  { id: '4', titulo: 'Alteracao Horario Mesas', data: '2026-03-10', destinatarios: 35, template: 'Aviso Importante' },
];

const EQUIPES = ['SP', 'LEBLON', 'PRODUTOS', 'CORPORATE', 'BACKOFFICE'];

// ── Component ────────────────────────────────────────────────────────────────

export default function Informativo() {
  const [titulo, setTitulo] = useState('');
  const [conteudo, setConteudo] = useState('');
  const [templateId, setTemplateId] = useState('mercado');
  const [destinatarios, setDestinatarios] = useState<Destinatario[]>(MOCK_DESTINATARIOS);
  const [equipeFilter, setEquipeFilter] = useState('TODAS');
  const [sending, setSending] = useState(false);
  const [sent, setSent] = useState(false);
  const [historico] = useState<HistoricoItem[]>(MOCK_HISTORICO);

  const selecionados = destinatarios.filter((d) => d.selecionado);

  const toggleDestinatario = (email: string) => {
    setDestinatarios((prev) =>
      prev.map((d) =>
        d.email === email ? { ...d, selecionado: !d.selecionado } : d
      )
    );
  };

  const selecionarEquipe = (equipe: string) => {
    setDestinatarios((prev) =>
      prev.map((d) =>
        d.equipe === equipe ? { ...d, selecionado: true } : d
      )
    );
  };

  const selecionarTodos = () => {
    setDestinatarios((prev) => prev.map((d) => ({ ...d, selecionado: true })));
  };

  const limparSelecao = () => {
    setDestinatarios((prev) => prev.map((d) => ({ ...d, selecionado: false })));
  };

  const filteredDestinatarios = equipeFilter === 'TODAS'
    ? destinatarios
    : destinatarios.filter((d) => d.equipe === equipeFilter);

  const handleEnviar = useCallback(async () => {
    if (!titulo || !conteudo || selecionados.length === 0) return;
    setSending(true);
    try {
      await createDrafts(
        selecionados.map((d) => ({
          to: d.email,
          subject: titulo,
          body: `<div style="font-family: DM Sans, sans-serif; color: #111;">
            <h2 style="color: #004D33;">${titulo}</h2>
            <div>${conteudo.replace(/\n/g, '<br/>')}</div>
            <hr style="border-color: #E5E7EB; margin-top: 20px;"/>
            <p style="color: #888; font-size: 12px;">Somus Capital - Mesa de Produtos</p>
          </div>`,
        }))
      );
      setSent(true);
      setTimeout(() => setSent(false), 3000);
    } catch (err) {
      console.error('Erro ao enviar:', err);
    } finally {
      setSending(false);
    }
  }, [titulo, conteudo, selecionados]);

  return (
    <div className="space-y-6">
      {/* Header */}
      <div>
        <h1 className="text-2xl font-bold text-somus-gray-900">
          Informativo
        </h1>
        <p className="text-sm text-somus-gray-500 mt-1">
          Crie e envie informativos para a equipe
        </p>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
        {/* Form */}
        <div className="lg:col-span-2 space-y-6">
          <Card title="Conteudo do Informativo">
            <div className="mt-4 space-y-4">
              {/* Template */}
              <div>
                <label className="block text-sm font-medium text-somus-gray-700 mb-1.5">
                  Template
                </label>
                <select
                  value={templateId}
                  onChange={(e) => setTemplateId(e.target.value)}
                  className="w-full text-sm border border-somus-gray-300 rounded-lg px-3 py-2 focus:ring-2 focus:ring-somus-green/40 focus:outline-none"
                >
                  {TEMPLATES.map((t) => (
                    <option key={t.id} value={t.id}>{t.nome}</option>
                  ))}
                </select>
              </div>

              {/* Titulo */}
              <div>
                <label className="block text-sm font-medium text-somus-gray-700 mb-1.5">
                  Titulo
                </label>
                <input
                  type="text"
                  value={titulo}
                  onChange={(e) => setTitulo(e.target.value)}
                  placeholder="Digite o titulo do informativo..."
                  className="w-full text-sm border border-somus-gray-300 rounded-lg px-3 py-2 focus:ring-2 focus:ring-somus-green/40 focus:outline-none"
                />
              </div>

              {/* Conteudo */}
              <div>
                <label className="block text-sm font-medium text-somus-gray-700 mb-1.5">
                  Conteudo
                </label>
                <textarea
                  value={conteudo}
                  onChange={(e) => setConteudo(e.target.value)}
                  placeholder="Digite o conteudo do informativo..."
                  rows={12}
                  className="w-full text-sm border border-somus-gray-300 rounded-lg px-3 py-2 focus:ring-2 focus:ring-somus-green/40 focus:outline-none resize-none"
                />
              </div>

              {/* Actions */}
              <div className="flex items-center gap-3 pt-2">
                <Button
                  variant="primary"
                  icon={<Send className="h-4 w-4" />}
                  loading={sending}
                  onClick={handleEnviar}
                  disabled={!titulo || !conteudo || selecionados.length === 0}
                >
                  {sent ? 'Rascunhos Criados!' : 'Enviar'}
                </Button>
                {sent && (
                  <span className="text-sm text-emerald-600 flex items-center gap-1">
                    <Check className="h-4 w-4" /> {selecionados.length} rascunhos criados
                  </span>
                )}
              </div>
            </div>
          </Card>
        </div>

        {/* Destinatarios */}
        <div className="space-y-6">
          <Card
            title="Destinatarios"
            subtitle={`${selecionados.length} selecionados`}
          >
            <div className="mt-4 space-y-3">
              {/* Quick actions */}
              <div className="flex flex-wrap gap-2">
                <button
                  onClick={selecionarTodos}
                  className="text-xs px-2 py-1 rounded bg-somus-green/10 text-somus-green hover:bg-somus-green/20 transition-colors"
                >
                  Selecionar todos
                </button>
                <button
                  onClick={limparSelecao}
                  className="text-xs px-2 py-1 rounded bg-somus-gray-100 text-somus-gray-600 hover:bg-somus-gray-200 transition-colors"
                >
                  Limpar
                </button>
              </div>

              {/* Equipe buttons */}
              <div className="flex flex-wrap gap-1.5">
                {EQUIPES.map((eq) => (
                  <button
                    key={eq}
                    onClick={() => selecionarEquipe(eq)}
                    className="text-xs px-2 py-1 rounded border border-somus-gray-200 text-somus-gray-600 hover:bg-somus-green/10 hover:text-somus-green hover:border-somus-green/30 transition-colors"
                  >
                    {eq}
                  </button>
                ))}
              </div>

              {/* Filter */}
              <select
                value={equipeFilter}
                onChange={(e) => setEquipeFilter(e.target.value)}
                className="w-full text-xs border border-somus-gray-300 rounded-lg px-2 py-1.5 focus:ring-2 focus:ring-somus-green/40 focus:outline-none"
              >
                <option value="TODAS">Todas as equipes</option>
                {EQUIPES.map((eq) => (
                  <option key={eq} value={eq}>{eq}</option>
                ))}
              </select>

              {/* List */}
              <div className="max-h-64 overflow-y-auto space-y-1">
                {filteredDestinatarios.map((d) => (
                  <label
                    key={d.email}
                    className={cn(
                      'flex items-center gap-3 py-2 px-3 rounded-lg cursor-pointer transition-colors',
                      d.selecionado
                        ? 'bg-somus-green/5 border border-somus-green/20'
                        : 'hover:bg-somus-gray-50 border border-transparent'
                    )}
                  >
                    <input
                      type="checkbox"
                      checked={d.selecionado}
                      onChange={() => toggleDestinatario(d.email)}
                      className="w-4 h-4 rounded border-somus-gray-300 text-somus-green focus:ring-somus-green/40"
                    />
                    <div className="flex-1 min-w-0">
                      <div className="text-sm font-medium text-somus-gray-900 truncate">
                        {d.nome}
                      </div>
                      <div className="text-xs text-somus-gray-400">{d.equipe}</div>
                    </div>
                  </label>
                ))}
              </div>
            </div>
          </Card>
        </div>
      </div>

      {/* Historico */}
      <Card title="Historico de Envios" subtitle="Ultimos informativos enviados">
        <div className="mt-4 overflow-x-auto">
          <table className="w-full text-sm">
            <thead>
              <tr className="border-b border-somus-gray-200">
                <th className="text-left py-3 px-4 font-semibold text-somus-gray-600">Titulo</th>
                <th className="text-left py-3 px-4 font-semibold text-somus-gray-600">Template</th>
                <th className="text-left py-3 px-4 font-semibold text-somus-gray-600">Data</th>
                <th className="text-right py-3 px-4 font-semibold text-somus-gray-600">Destinatarios</th>
              </tr>
            </thead>
            <tbody>
              {historico.map((item) => (
                <tr
                  key={item.id}
                  className="border-b border-somus-gray-100 hover:bg-somus-gray-50 transition-colors"
                >
                  <td className="py-3 px-4 font-medium text-somus-gray-900">
                    <div className="flex items-center gap-2">
                      <Mail className="h-4 w-4 text-somus-green" />
                      {item.titulo}
                    </div>
                  </td>
                  <td className="py-3 px-4 text-somus-gray-600">{item.template}</td>
                  <td className="py-3 px-4 text-somus-gray-600">
                    {new Date(item.data + 'T00:00:00').toLocaleDateString('pt-BR')}
                  </td>
                  <td className="py-3 px-4 text-right text-somus-gray-600">
                    {item.destinatarios}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </Card>
    </div>
  );
}

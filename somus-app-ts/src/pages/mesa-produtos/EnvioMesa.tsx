import React, { useState, useCallback } from 'react';
import { Card } from '@components/Card';
import { Button } from '@components/Button';
import {
  Send,
  Users,
  Check,
  Mail,
  Plus,
  Trash2,
} from 'lucide-react';
import { cn } from '@/lib/utils';
import { createDrafts } from '@/services/outlook';

// ── Types ────────────────────────────────────────────────────────────────────

interface DataRow {
  id: string;
  campo: string;
  valor: string;
}

interface Destinatario {
  nome: string;
  email: string;
  selecionado: boolean;
}

// ── Mock Data ────────────────────────────────────────────────────────────────

const MOCK_DESTINATARIOS: Destinatario[] = [
  { nome: 'Carlos Silva', email: 'carlos@somus.com', selecionado: false },
  { nome: 'Ana Santos', email: 'ana@somus.com', selecionado: false },
  { nome: 'Pedro Costa', email: 'pedro@somus.com', selecionado: false },
  { nome: 'Maria Oliveira', email: 'maria@somus.com', selecionado: false },
  { nome: 'Joao Souza', email: 'joao@somus.com', selecionado: false },
  { nome: 'Fernanda Lima', email: 'fernanda@somus.com', selecionado: false },
  { nome: 'Lucas Almeida', email: 'lucas@somus.com', selecionado: false },
  { nome: 'Juliana Rocha', email: 'juliana@somus.com', selecionado: false },
];

const EMAIL_TEMPLATES = [
  { id: 'atualizacao', nome: 'Atualizacao de Mesa', assunto: 'Atualizacao Mesa de Produtos' },
  { id: 'abertura', nome: 'Abertura de Mercado', assunto: 'Abertura de Mercado - Destaques' },
  { id: 'fechamento', nome: 'Fechamento de Mercado', assunto: 'Fechamento de Mercado - Resumo' },
  { id: 'oportunidade', nome: 'Oportunidade', assunto: 'Oportunidade de Investimento' },
  { id: 'custom', nome: 'Personalizado', assunto: '' },
];

// ── Component ────────────────────────────────────────────────────────────────

export default function EnvioMesa() {
  const [templateId, setTemplateId] = useState('atualizacao');
  const [assunto, setAssunto] = useState(EMAIL_TEMPLATES[0].assunto);
  const [dataRows, setDataRows] = useState<DataRow[]>([
    { id: '1', campo: 'Tesouro Selic', valor: '14,75% a.a.' },
    { id: '2', campo: 'CDB XP IPCA+', valor: '7,20% a.a.' },
    { id: '3', campo: 'LCI Itau', valor: '97% CDI' },
  ]);
  const [observacoes, setObservacoes] = useState('');
  const [destinatarios, setDestinatarios] = useState<Destinatario[]>(MOCK_DESTINATARIOS);
  const [sending, setSending] = useState(false);
  const [sent, setSent] = useState(false);

  const selecionados = destinatarios.filter((d) => d.selecionado);

  const handleTemplateChange = (id: string) => {
    setTemplateId(id);
    const template = EMAIL_TEMPLATES.find((t) => t.id === id);
    if (template && template.assunto) {
      setAssunto(template.assunto);
    }
  };

  const addRow = () => {
    setDataRows((prev) => [
      ...prev,
      { id: String(Date.now()), campo: '', valor: '' },
    ]);
  };

  const removeRow = (id: string) => {
    setDataRows((prev) => prev.filter((r) => r.id !== id));
  };

  const updateRow = (id: string, field: 'campo' | 'valor', value: string) => {
    setDataRows((prev) =>
      prev.map((r) => (r.id === id ? { ...r, [field]: value } : r))
    );
  };

  const toggleDestinatario = (email: string) => {
    setDestinatarios((prev) =>
      prev.map((d) =>
        d.email === email ? { ...d, selecionado: !d.selecionado } : d
      )
    );
  };

  const selecionarTodos = () => {
    setDestinatarios((prev) => prev.map((d) => ({ ...d, selecionado: true })));
  };

  const limparSelecao = () => {
    setDestinatarios((prev) => prev.map((d) => ({ ...d, selecionado: false })));
  };

  const handleEnviar = useCallback(async () => {
    if (!assunto || selecionados.length === 0) return;
    setSending(true);
    try {
      const tableHtml = dataRows
        .filter((r) => r.campo)
        .map(
          (r) => `<tr><td style="padding: 8px; border: 1px solid #E5E7EB; font-weight: 600;">${r.campo}</td><td style="padding: 8px; border: 1px solid #E5E7EB;">${r.valor}</td></tr>`
        )
        .join('');

      await createDrafts(
        selecionados.map((d) => ({
          to: d.email,
          subject: assunto,
          body: `<div style="font-family: DM Sans, sans-serif; max-width: 600px;">
            <div style="background: #004D33; color: white; padding: 16px 20px; border-radius: 8px 8px 0 0;">
              <h2 style="margin: 0; font-size: 18px;">${assunto}</h2>
              <p style="margin: 4px 0 0; opacity: 0.8; font-size: 12px;">${new Date().toLocaleDateString('pt-BR')} - Mesa de Produtos</p>
            </div>
            <div style="padding: 20px; border: 1px solid #E5E7EB; border-top: none; border-radius: 0 0 8px 8px;">
              ${tableHtml ? `
                <table style="width: 100%; border-collapse: collapse; font-size: 13px;">
                  <tr style="background: #F0F0E0;">
                    <th style="padding: 8px; text-align: left; border: 1px solid #E5E7EB;">Item</th>
                    <th style="padding: 8px; text-align: left; border: 1px solid #E5E7EB;">Valor/Taxa</th>
                  </tr>
                  ${tableHtml}
                </table>
              ` : ''}
              ${observacoes ? `<p style="margin-top: 15px; font-size: 13px; color: #374151;">${observacoes.replace(/\n/g, '<br/>')}</p>` : ''}
              <hr style="border-color: #E5E7EB; margin-top: 20px;"/>
              <p style="color: #888; font-size: 11px;">Somus Capital - Mesa de Produtos</p>
            </div>
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
  }, [assunto, dataRows, observacoes, selecionados]);

  return (
    <div className="space-y-6">
      {/* Header */}
      <div>
        <h1 className="text-2xl font-bold text-somus-text-primary">Envio Mesa</h1>
        <p className="text-sm text-somus-text-secondary mt-1">
          Envie dados e atualizacoes da mesa para a equipe
        </p>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
        {/* Form */}
        <div className="lg:col-span-2 space-y-6">
          <Card title="Dados do Envio">
            <div className="mt-4 space-y-4">
              {/* Template */}
              <div>
                <label className="block text-sm font-medium text-somus-text-primary mb-1.5">Template</label>
                <select
                  value={templateId}
                  onChange={(e) => handleTemplateChange(e.target.value)}
                  className="w-full text-sm border border-somus-border rounded-lg px-3 py-2 bg-somus-bg-input text-somus-text-primary placeholder:text-somus-text-tertiary focus:ring-2 focus:ring-somus-green/40 focus:outline-none"
                >
                  {EMAIL_TEMPLATES.map((t) => (
                    <option key={t.id} value={t.id}>{t.nome}</option>
                  ))}
                </select>
              </div>

              {/* Assunto */}
              <div>
                <label className="block text-sm font-medium text-somus-text-primary mb-1.5">Assunto</label>
                <input
                  type="text"
                  value={assunto}
                  onChange={(e) => setAssunto(e.target.value)}
                  className="w-full text-sm border border-somus-border rounded-lg px-3 py-2 bg-somus-bg-input text-somus-text-primary placeholder:text-somus-text-tertiary focus:ring-2 focus:ring-somus-green/40 focus:outline-none"
                />
              </div>

              {/* Data Rows */}
              <div>
                <div className="flex items-center justify-between mb-2">
                  <label className="text-sm font-medium text-somus-text-primary">Dados</label>
                  <button
                    onClick={addRow}
                    className="text-xs text-somus-green hover:text-somus-green-light flex items-center gap-1"
                  >
                    <Plus className="h-3 w-3" /> Adicionar linha
                  </button>
                </div>
                <div className="space-y-2">
                  {dataRows.map((row) => (
                    <div key={row.id} className="flex items-center gap-2">
                      <input
                        type="text"
                        placeholder="Campo"
                        value={row.campo}
                        onChange={(e) => updateRow(row.id, 'campo', e.target.value)}
                        className="flex-1 text-sm border border-somus-border rounded-lg px-3 py-2 bg-somus-bg-input text-somus-text-primary placeholder:text-somus-text-tertiary focus:ring-2 focus:ring-somus-green/40 focus:outline-none"
                      />
                      <input
                        type="text"
                        placeholder="Valor"
                        value={row.valor}
                        onChange={(e) => updateRow(row.id, 'valor', e.target.value)}
                        className="flex-1 text-sm border border-somus-border rounded-lg px-3 py-2 bg-somus-bg-input text-somus-text-primary placeholder:text-somus-text-tertiary focus:ring-2 focus:ring-somus-green/40 focus:outline-none"
                      />
                      <button
                        onClick={() => removeRow(row.id)}
                        className="text-somus-text-tertiary hover:text-red-400 transition-colors p-1"
                      >
                        <Trash2 className="h-4 w-4" />
                      </button>
                    </div>
                  ))}
                </div>
              </div>

              {/* Observacoes */}
              <div>
                <label className="block text-sm font-medium text-somus-text-primary mb-1.5">Observacoes</label>
                <textarea
                  value={observacoes}
                  onChange={(e) => setObservacoes(e.target.value)}
                  placeholder="Observacoes adicionais..."
                  rows={4}
                  className="w-full text-sm border border-somus-border rounded-lg px-3 py-2 bg-somus-bg-input text-somus-text-primary placeholder:text-somus-text-tertiary focus:ring-2 focus:ring-somus-green/40 focus:outline-none resize-none"
                />
              </div>

              {/* Send button */}
              <div className="flex items-center gap-3 pt-2">
                <Button
                  variant="primary"
                  icon={<Send className="h-4 w-4" />}
                  loading={sending}
                  onClick={handleEnviar}
                  disabled={!assunto || selecionados.length === 0}
                >
                  {sent ? 'Enviado!' : 'Enviar'}
                </Button>
                {sent && (
                  <span className="text-sm text-somus-green-400 flex items-center gap-1">
                    <Check className="h-4 w-4" /> {selecionados.length} rascunhos criados
                  </span>
                )}
              </div>
            </div>
          </Card>
        </div>

        {/* Destinatarios */}
        <Card
          title="Destinatarios"
          subtitle={`${selecionados.length} selecionados`}
        >
          <div className="mt-4 space-y-3">
            <div className="flex gap-2">
              <button
                onClick={selecionarTodos}
                className="text-xs px-2 py-1 rounded bg-somus-green/10 text-somus-green hover:bg-somus-green/20 transition-colors"
              >
                Selecionar todos
              </button>
              <button
                onClick={limparSelecao}
                className="text-xs px-2 py-1 rounded bg-somus-border/30 text-somus-text-secondary hover:bg-somus-border transition-colors"
              >
                Limpar
              </button>
            </div>

            <div className="space-y-1">
              {destinatarios.map((d) => (
                <label
                  key={d.email}
                  className={cn(
                    'flex items-center gap-3 py-2 px-3 rounded-lg cursor-pointer transition-colors',
                    d.selecionado
                      ? 'bg-somus-green/5 border border-somus-green/20'
                      : 'hover:bg-somus-bg-hover border border-transparent'
                  )}
                >
                  <input
                    type="checkbox"
                    checked={d.selecionado}
                    onChange={() => toggleDestinatario(d.email)}
                    className="w-4 h-4 rounded border-somus-border text-somus-green focus:ring-somus-green/40"
                  />
                  <div>
                    <div className="text-sm font-medium text-somus-text-primary">{d.nome}</div>
                    <div className="text-xs text-somus-text-tertiary">{d.email}</div>
                  </div>
                </label>
              ))}
            </div>
          </div>
        </Card>
      </div>
    </div>
  );
}

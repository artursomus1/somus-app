import React, { useState } from 'react';
import { useForm, Controller } from 'react-hook-form';
import { zodResolver } from '@hookform/resolvers/zod';
import { z } from 'zod';
import { ArrowLeft, Download, Eye, Plus, Minus } from 'lucide-react';
import PptxGenJS from 'pptxgenjs';
import { PageLayout } from '@components/PageLayout';
import { Card } from '@components/Card';
import { Button } from '@components/Button';
import { FormField } from '@components/FormField';
import { CurrencyInput } from '@components/CurrencyInput';
import { PercentInput } from '@components/PercentInput';
import { Tabs } from '@components/Tabs';
import { useAppStore } from '@/stores/appStore';
import type { PropostaSubtipo } from '@/types';

// ── PPTX brand colors ──────────────────────────────────────────────────────

const BRAND = {
  darkGreen: '004D33',
  lightGreen: '005C3D',
  white: 'FFFFFF',
  lightBg: 'F0FFF8',
  gray: '6B7280',
  darkText: '111111',
};

// ── Format helpers ──────────────────────────────────────────────────────────

function fmtBrl(v: number): string {
  const abs = Math.abs(v);
  const inteiro = Math.floor(abs);
  const centavos = Math.round((abs - inteiro) * 100);
  const parte = inteiro.toLocaleString('pt-BR');
  const r = `R$ ${parte},${String(centavos).padStart(2, '0')}`;
  return v < 0 ? `-${r}` : r;
}

function fmtPctAm(v: number): string {
  return `${v.toLocaleString('pt-BR', { minimumFractionDigits: 4, maximumFractionDigits: 4 })}% a.m.`;
}

function fmtPctAa(v: number): string {
  return `${v.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}% a.a.`;
}

// ── Cenario schema ──────────────────────────────────────────────────────────

const cenarioSchema = z.object({
  contemplacaoMes: z.number().min(1),
  cetMensal: z.number(),
  cetAnual: z.number(),
  lanceLivrePct: z.number().min(0).max(100),
  lanceLivreValor: z.number().min(0),
  creditoDisponivel: z.number().min(0),
  alavancagem: z.number(),
  valorCarta: z.number().min(1),
  lanceEmbutidoPct: z.number().min(0).max(100),
  lanceEmbutidoValor: z.number().min(0),
  taxaAdm: z.number().min(0).max(100),
  fundoReserva: z.number().min(0).max(100),
  prazoTotal: z.number().min(1),
  correcaoAnual: z.number().min(0).max(100),
  parcelaF1: z.number().min(0),
  parcelaF2: z.number().min(0),
});

type CenarioData = z.infer<typeof cenarioSchema>;

// ── Proposta schema ─────────────────────────────────────────────────────────

const propostaSchema = z.object({
  clienteNome: z.string().min(1, 'Nome obrigatorio'),
  dataMesAno: z.string().min(1, 'Data obrigatoria'),
  empresaNome: z.string().optional(),
  localizacao: z.string().optional(),
  descricaoOperacao: z.string().optional(),
  objetivoTitulo: z.string().optional(),
  horizonteTexto: z.string().optional(),
  objetivoDescricao: z.string().optional(),
  administradora: z.string().min(1, 'Administradora obrigatoria'),
  contato1Nome: z.string().optional(),
  contato1Telefone: z.string().optional(),
  contato2Nome: z.string().optional(),
  contato2Telefone: z.string().optional(),
  cenario: cenarioSchema,
});

type PropostaFormData = z.infer<typeof propostaSchema>;

// ── Comparativa schema ──────────────────────────────────────────────────────

const comparativaSchema = z.object({
  clienteNome: z.string().min(1, 'Nome obrigatorio'),
  dataMesAno: z.string().min(1, 'Data obrigatoria'),
  contato1Nome: z.string().optional(),
  contato1Telefone: z.string().optional(),
  contato2Nome: z.string().optional(),
  contato2Telefone: z.string().optional(),
  cenarios: z.array(cenarioSchema).min(1).max(3),
});

type ComparativaFormData = z.infer<typeof comparativaSchema>;

// ── PPTX Generation ─────────────────────────────────────────────────────────

function generatePropostaPPTX(data: PropostaFormData, subtipo: PropostaSubtipo) {
  const pptx = new PptxGenJS();
  pptx.layout = 'LAYOUT_WIDE';
  pptx.author = 'Somus Capital';
  const c = data.cenario;

  // Slide 1: Cover
  const s1 = pptx.addSlide();
  s1.background = { color: BRAND.darkGreen };
  s1.addText('SOMUS CAPITAL', { x: 0.8, y: 1.5, w: 8, h: 0.6, fontSize: 36, bold: true, color: BRAND.white, fontFace: 'DM Sans' });
  s1.addText('Proposta de Consorcio', { x: 0.8, y: 2.2, w: 8, h: 0.5, fontSize: 20, color: BRAND.lightBg, fontFace: 'DM Sans' });
  s1.addText(data.clienteNome, { x: 0.8, y: 3.5, w: 8, h: 0.5, fontSize: 18, color: BRAND.white, fontFace: 'DM Sans' });
  s1.addText(data.dataMesAno, { x: 0.8, y: 4.1, w: 8, h: 0.4, fontSize: 14, color: BRAND.lightBg, fontFace: 'DM Sans' });
  s1.addText(`Subtipo: ${subtipo.replace(/_/g, ' ')}`, { x: 0.8, y: 4.6, w: 8, h: 0.3, fontSize: 12, color: BRAND.lightBg, fontFace: 'DM Sans' });

  // Slide 2: Contexto
  if (subtipo !== 'PF_Tradicional') {
    const s2 = pptx.addSlide();
    s2.addText('Contexto', { x: 0.8, y: 0.5, w: 8, h: 0.5, fontSize: 24, bold: true, color: BRAND.darkGreen, fontFace: 'DM Sans' });
    s2.addText(data.descricaoOperacao || 'Operacao de consorcio estruturada para aquisicao de bem imovel.', { x: 0.8, y: 1.3, w: 11, h: 2, fontSize: 14, color: BRAND.darkText, fontFace: 'DM Sans', wrap: true });
  }

  // Slide 3: Objetivo
  const s3 = pptx.addSlide();
  s3.addText(data.objetivoTitulo || 'Objetivo', { x: 0.8, y: 0.5, w: 8, h: 0.5, fontSize: 24, bold: true, color: BRAND.darkGreen, fontFace: 'DM Sans' });
  s3.addText(data.objetivoDescricao || 'Estruturar operacao de consorcio com as melhores condicoes de mercado.', { x: 0.8, y: 1.3, w: 11, h: 2, fontSize: 14, color: BRAND.darkText, fontFace: 'DM Sans', wrap: true });
  s3.addText(`Horizonte: ${data.horizonteTexto || `${c.prazoTotal} meses`}`, { x: 0.8, y: 3.5, w: 8, h: 0.4, fontSize: 14, bold: true, color: BRAND.lightGreen, fontFace: 'DM Sans' });

  // Slide 4: Estrutura
  const s4 = pptx.addSlide();
  s4.addText('Estrutura da Operacao', { x: 0.8, y: 0.5, w: 8, h: 0.5, fontSize: 24, bold: true, color: BRAND.darkGreen, fontFace: 'DM Sans' });
  const hOpts = { bold: true, fill: { color: BRAND.darkGreen }, color: BRAND.white, fontSize: 12, fontFace: 'DM Sans' } as const;
  const rOpts = { fontSize: 11, fontFace: 'DM Sans' } as const;
  const bOpts = { fontSize: 11, fontFace: 'DM Sans', bold: true } as const;
  const rows: PptxGenJS.TableRow[] = [
    [{ text: 'Parametro', options: hOpts }, { text: 'Valor', options: hOpts }],
    [{ text: 'Administradora', options: rOpts }, { text: data.administradora, options: bOpts }],
    [{ text: 'Valor da Carta', options: rOpts }, { text: fmtBrl(c.valorCarta), options: bOpts }],
    [{ text: 'Prazo Total', options: rOpts }, { text: `${c.prazoTotal} meses`, options: bOpts }],
    [{ text: 'Taxa de Administracao', options: rOpts }, { text: `${c.taxaAdm.toFixed(2)}%`, options: bOpts }],
    [{ text: 'Fundo de Reserva', options: rOpts }, { text: `${c.fundoReserva.toFixed(2)}%`, options: bOpts }],
    [{ text: 'Correcao Anual', options: rOpts }, { text: `${c.correcaoAnual.toFixed(2)}%`, options: bOpts }],
    [{ text: 'Contemplacao', options: rOpts }, { text: `${c.contemplacaoMes}o mes`, options: bOpts }],
  ];
  s4.addTable(rows, { x: 0.8, y: 1.3, w: 11, colW: [5.5, 5.5], border: { pt: 0.5, color: 'D1D5DB' }, rowH: 0.4 });

  // Slide 5: Lances
  const s5 = pptx.addSlide();
  s5.addText('Lances', { x: 0.8, y: 0.5, w: 8, h: 0.5, fontSize: 24, bold: true, color: BRAND.darkGreen, fontFace: 'DM Sans' });
  const lRows: PptxGenJS.TableRow[] = [
    [{ text: 'Tipo', options: hOpts }, { text: '%', options: hOpts }, { text: 'Valor', options: hOpts }],
    [{ text: 'Lance Livre', options: rOpts }, { text: `${c.lanceLivrePct.toFixed(2)}%`, options: rOpts }, { text: fmtBrl(c.lanceLivreValor), options: bOpts }],
    [{ text: 'Lance Embutido', options: rOpts }, { text: `${c.lanceEmbutidoPct.toFixed(2)}%`, options: rOpts }, { text: fmtBrl(c.lanceEmbutidoValor), options: bOpts }],
    [{ text: 'Credito Disponivel', options: { ...rOpts, bold: true } }, { text: '', options: rOpts }, { text: fmtBrl(c.creditoDisponivel), options: { ...bOpts, color: BRAND.darkGreen } }],
  ];
  s5.addTable(lRows, { x: 0.8, y: 1.3, w: 11, colW: [4, 3, 4], border: { pt: 0.5, color: 'D1D5DB' }, rowH: 0.4 });

  // Slide 6: Parcelas
  const s6 = pptx.addSlide();
  s6.addText('Parcelas', { x: 0.8, y: 0.5, w: 8, h: 0.5, fontSize: 24, bold: true, color: BRAND.darkGreen, fontFace: 'DM Sans' });
  s6.addText(`Fase 1 (meses 1 a ${c.contemplacaoMes}):`, { x: 0.8, y: 1.3, w: 8, h: 0.4, fontSize: 14, bold: true, color: BRAND.lightGreen, fontFace: 'DM Sans' });
  s6.addText(fmtBrl(c.parcelaF1), { x: 0.8, y: 1.8, w: 8, h: 0.5, fontSize: 22, bold: true, color: BRAND.darkGreen, fontFace: 'DM Sans' });
  s6.addText(`Fase 2 (meses ${c.contemplacaoMes + 1} a ${c.prazoTotal}):`, { x: 0.8, y: 2.8, w: 8, h: 0.4, fontSize: 14, bold: true, color: BRAND.lightGreen, fontFace: 'DM Sans' });
  s6.addText(fmtBrl(c.parcelaF2), { x: 0.8, y: 3.3, w: 8, h: 0.5, fontSize: 22, bold: true, color: BRAND.darkGreen, fontFace: 'DM Sans' });

  // Slide 7: CET
  const s7 = pptx.addSlide();
  s7.addText('Custo Efetivo Total', { x: 0.8, y: 0.5, w: 8, h: 0.5, fontSize: 24, bold: true, color: BRAND.darkGreen, fontFace: 'DM Sans' });
  s7.addText(`CET Mensal: ${fmtPctAm(c.cetMensal)}`, { x: 0.8, y: 1.5, w: 8, h: 0.5, fontSize: 18, bold: true, color: BRAND.darkText, fontFace: 'DM Sans' });
  s7.addText(`CET Anual: ${fmtPctAa(c.cetAnual)}`, { x: 0.8, y: 2.2, w: 8, h: 0.5, fontSize: 18, bold: true, color: BRAND.darkText, fontFace: 'DM Sans' });
  s7.addText(`Alavancagem: ${c.alavancagem.toFixed(2)}x`, { x: 0.8, y: 3.2, w: 8, h: 0.4, fontSize: 16, color: BRAND.lightGreen, fontFace: 'DM Sans' });

  // Slide 8: Cronograma
  const s8 = pptx.addSlide();
  s8.addText('Cronograma', { x: 0.8, y: 0.5, w: 8, h: 0.5, fontSize: 24, bold: true, color: BRAND.darkGreen, fontFace: 'DM Sans' });
  [
    { s: '1', t: `Adesao ao consorcio - ${data.administradora}` },
    { s: '2', t: `Contemplacao no ${c.contemplacaoMes}o mes via lance` },
    { s: '3', t: 'Utilizacao do credito para aquisicao do bem' },
    { s: '4', t: `Pagamento das parcelas ate o ${c.prazoTotal}o mes` },
  ].forEach((item, idx) => {
    const y = 1.3 + idx * 0.7;
    s8.addShape('oval' as any, { x: 0.8, y, w: 0.4, h: 0.4, fill: { color: BRAND.darkGreen } });
    s8.addText(item.s, { x: 0.8, y, w: 0.4, h: 0.4, fontSize: 12, bold: true, color: BRAND.white, fontFace: 'DM Sans', align: 'center', valign: 'middle' });
    s8.addText(item.t, { x: 1.5, y, w: 10, h: 0.4, fontSize: 13, color: BRAND.darkText, fontFace: 'DM Sans', valign: 'middle' });
  });

  // Slide 9: Contatos
  const s9 = pptx.addSlide();
  s9.background = { color: BRAND.darkGreen };
  s9.addText('Contatos', { x: 0.8, y: 1, w: 8, h: 0.5, fontSize: 24, bold: true, color: BRAND.white, fontFace: 'DM Sans' });
  if (data.contato1Nome) {
    s9.addText(data.contato1Nome, { x: 0.8, y: 2, w: 6, h: 0.4, fontSize: 16, bold: true, color: BRAND.white, fontFace: 'DM Sans' });
    s9.addText(data.contato1Telefone || '', { x: 0.8, y: 2.5, w: 6, h: 0.3, fontSize: 14, color: BRAND.lightBg, fontFace: 'DM Sans' });
  }
  if (data.contato2Nome) {
    s9.addText(data.contato2Nome, { x: 0.8, y: 3.3, w: 6, h: 0.4, fontSize: 16, bold: true, color: BRAND.white, fontFace: 'DM Sans' });
    s9.addText(data.contato2Telefone || '', { x: 0.8, y: 3.8, w: 6, h: 0.3, fontSize: 14, color: BRAND.lightBg, fontFace: 'DM Sans' });
  }

  // Slide 10: Disclaimer
  const s10 = pptx.addSlide();
  s10.addText('Aviso Legal', { x: 0.8, y: 0.5, w: 8, h: 0.5, fontSize: 24, bold: true, color: BRAND.darkGreen, fontFace: 'DM Sans' });
  s10.addText(
    'Esta apresentacao tem carater exclusivamente informativo e nao constitui oferta ou recomendacao de investimento. Os valores apresentados sao estimativas e podem variar conforme condicoes de mercado. Somus Capital.',
    { x: 0.8, y: 1.3, w: 11, h: 3, fontSize: 12, color: BRAND.gray, fontFace: 'DM Sans', wrap: true },
  );

  return pptx;
}

function generateComparativaPPTX(data: ComparativaFormData) {
  const pptx = new PptxGenJS();
  pptx.layout = 'LAYOUT_WIDE';
  pptx.author = 'Somus Capital';

  const s1 = pptx.addSlide();
  s1.background = { color: BRAND.darkGreen };
  s1.addText('SOMUS CAPITAL', { x: 0.8, y: 1.5, w: 8, h: 0.6, fontSize: 36, bold: true, color: BRAND.white, fontFace: 'DM Sans' });
  s1.addText('Comparativo de Cenarios', { x: 0.8, y: 2.2, w: 8, h: 0.5, fontSize: 20, color: BRAND.lightBg, fontFace: 'DM Sans' });
  s1.addText(data.clienteNome, { x: 0.8, y: 3.5, w: 8, h: 0.5, fontSize: 18, color: BRAND.white, fontFace: 'DM Sans' });
  s1.addText(data.dataMesAno, { x: 0.8, y: 4.1, w: 8, h: 0.4, fontSize: 14, color: BRAND.lightBg, fontFace: 'DM Sans' });

  const s2 = pptx.addSlide();
  s2.addText('Comparativo', { x: 0.8, y: 0.3, w: 8, h: 0.5, fontSize: 24, bold: true, color: BRAND.darkGreen, fontFace: 'DM Sans' });

  const hOpts = { bold: true, fill: { color: BRAND.darkGreen }, color: BRAND.white, fontSize: 10, fontFace: 'DM Sans' } as const;
  const headerRow: PptxGenJS.TableRow = [
    { text: 'Parametro', options: hOpts },
    ...data.cenarios.map((_, i) => ({ text: `Cenario ${i + 1}`, options: hOpts } as PptxGenJS.TableCell)),
  ];

  const rOpts = { fontSize: 9, fontFace: 'DM Sans' } as const;
  const bOpts = { fontSize: 9, fontFace: 'DM Sans', bold: true } as const;
  const row = (label: string, fn: (c: CenarioData) => string): PptxGenJS.TableRow => [
    { text: label, options: rOpts },
    ...data.cenarios.map((c) => ({ text: fn(c), options: bOpts } as PptxGenJS.TableCell)),
  ];

  const tableRows: PptxGenJS.TableRow[] = [
    headerRow,
    row('Valor da Carta', (c) => fmtBrl(c.valorCarta)),
    row('Prazo Total', (c) => `${c.prazoTotal} meses`),
    row('Contemplacao', (c) => `${c.contemplacaoMes}o mes`),
    row('Lance Livre', (c) => `${c.lanceLivrePct.toFixed(2)}% (${fmtBrl(c.lanceLivreValor)})`),
    row('Lance Embutido', (c) => `${c.lanceEmbutidoPct.toFixed(2)}% (${fmtBrl(c.lanceEmbutidoValor)})`),
    row('Credito Disponivel', (c) => fmtBrl(c.creditoDisponivel)),
    row('Taxa Adm', (c) => `${c.taxaAdm.toFixed(2)}%`),
    row('Fdo Reserva', (c) => `${c.fundoReserva.toFixed(2)}%`),
    row('Correcao', (c) => `${c.correcaoAnual.toFixed(2)}%`),
    row('CET Mensal', (c) => fmtPctAm(c.cetMensal)),
    row('CET Anual', (c) => fmtPctAa(c.cetAnual)),
    row('Alavancagem', (c) => `${c.alavancagem.toFixed(2)}x`),
    row('Parcela F1', (c) => fmtBrl(c.parcelaF1)),
    row('Parcela F2', (c) => fmtBrl(c.parcelaF2)),
  ];

  const colW = [3, ...data.cenarios.map(() => (11 - 3) / data.cenarios.length)];
  s2.addTable(tableRows, { x: 0.4, y: 1, w: 12.4, colW, border: { pt: 0.5, color: 'D1D5DB' }, rowH: 0.32 });

  const s3 = pptx.addSlide();
  s3.background = { color: BRAND.darkGreen };
  s3.addText('Contatos', { x: 0.8, y: 1, w: 8, h: 0.5, fontSize: 24, bold: true, color: BRAND.white, fontFace: 'DM Sans' });
  if (data.contato1Nome) {
    s3.addText(data.contato1Nome, { x: 0.8, y: 2, w: 6, h: 0.4, fontSize: 16, bold: true, color: BRAND.white, fontFace: 'DM Sans' });
    s3.addText(data.contato1Telefone || '', { x: 0.8, y: 2.5, w: 6, h: 0.3, fontSize: 14, color: BRAND.lightBg, fontFace: 'DM Sans' });
  }
  if (data.contato2Nome) {
    s3.addText(data.contato2Nome, { x: 0.8, y: 3.3, w: 6, h: 0.4, fontSize: 16, bold: true, color: BRAND.white, fontFace: 'DM Sans' });
    s3.addText(data.contato2Telefone || '', { x: 0.8, y: 3.8, w: 6, h: 0.3, fontSize: 14, color: BRAND.lightBg, fontFace: 'DM Sans' });
  }

  return pptx;
}

// ── Defaults ────────────────────────────────────────────────────────────────

const defaultCenario: CenarioData = {
  contemplacaoMes: 3, cetMensal: 0.62, cetAnual: 7.71,
  lanceLivrePct: 20, lanceLivreValor: 100000, creditoDisponivel: 450000,
  alavancagem: 4.5, valorCarta: 500000, lanceEmbutidoPct: 10,
  lanceEmbutidoValor: 50000, taxaAdm: 20, fundoReserva: 3,
  prazoTotal: 200, correcaoAnual: 7, parcelaF1: 2925, parcelaF2: 2787.5,
};

const defaultProposta: PropostaFormData = {
  clienteNome: '', dataMesAno: '', empresaNome: '', localizacao: '',
  descricaoOperacao: '', objetivoTitulo: 'Objetivo', horizonteTexto: '',
  objetivoDescricao: '', administradora: '', contato1Nome: '',
  contato1Telefone: '', contato2Nome: '', contato2Telefone: '',
  cenario: defaultCenario,
};

const defaultComparativa: ComparativaFormData = {
  clienteNome: '', dataMesAno: '', contato1Nome: '', contato1Telefone: '',
  contato2Nome: '', contato2Telefone: '',
  cenarios: [{ ...defaultCenario }, { ...defaultCenario }, { ...defaultCenario }],
};

const SUBTIPOS: { value: PropostaSubtipo; label: string }[] = [
  { value: 'CETHD', label: 'CET HD' },
  { value: 'CETHD_Lance_Fixo', label: 'CET HD Lance Fixo' },
  { value: 'Lance_Fixo', label: 'Lance Fixo' },
  { value: 'PF_Tradicional', label: 'PF Tradicional' },
  { value: 'PJ_Tradicional', label: 'PJ Tradicional' },
];

const TABS = [
  { id: 'proposta', label: 'Proposta' },
  { id: 'comparativa', label: 'Comparativa' },
];

// ── Main Component ──────────────────────────────────────────────────────────

export default function GeradorPropostas() {
  const setPage = useAppStore((s) => s.setPage);
  const [activeTab, setActiveTab] = useState('proposta');
  const [subtipo, setSubtipo] = useState<PropostaSubtipo>('CETHD');
  const [generating, setGenerating] = useState(false);
  const [numCenarios, setNumCenarios] = useState(3);

  const propostaForm = useForm<PropostaFormData>({ resolver: zodResolver(propostaSchema), defaultValues: defaultProposta });
  const comparativaForm = useForm<ComparativaFormData>({ resolver: zodResolver(comparativaSchema), defaultValues: defaultComparativa });

  async function handleGerarProposta(data: PropostaFormData) {
    setGenerating(true);
    try {
      const pptx = generatePropostaPPTX(data, subtipo);
      await pptx.writeFile({ fileName: `Proposta_${data.clienteNome.replace(/\s+/g, '_')}_${subtipo}.pptx` });
    } finally {
      setGenerating(false);
    }
  }

  async function handleGerarComparativa(data: ComparativaFormData) {
    setGenerating(true);
    try {
      const limitedData = { ...data, cenarios: data.cenarios.slice(0, numCenarios) };
      const pptx = generateComparativaPPTX(limitedData);
      await pptx.writeFile({ fileName: `Comparativa_${data.clienteNome.replace(/\s+/g, '_')}.pptx` });
    } finally {
      setGenerating(false);
    }
  }

  const inputClass = 'w-full px-3 py-2 text-sm border border-somus-border rounded-lg focus:ring-2 focus:ring-somus-green/40 focus:border-somus-green outline-none bg-somus-bg-input text-somus-text-primary';

  // Fields for cenario card in comparativa
  const cenarioFields = [
    { label: 'Valor Carta', key: 'valorCarta', step: '1' },
    { label: 'Prazo (meses)', key: 'prazoTotal', step: '1' },
    { label: 'Contemplacao (mes)', key: 'contemplacaoMes', step: '1' },
    { label: 'Taxa Adm (%)', key: 'taxaAdm', step: '0.01' },
    { label: 'Fdo Reserva (%)', key: 'fundoReserva', step: '0.01' },
    { label: 'Correcao (%)', key: 'correcaoAnual', step: '0.01' },
    { label: 'Lance Livre (%)', key: 'lanceLivrePct', step: '0.01' },
    { label: 'Lance Livre (R$)', key: 'lanceLivreValor', step: '1' },
    { label: 'Lance Emb. (%)', key: 'lanceEmbutidoPct', step: '0.01' },
    { label: 'Lance Emb. (R$)', key: 'lanceEmbutidoValor', step: '1' },
    { label: 'Credito Disp. (R$)', key: 'creditoDisponivel', step: '1' },
    { label: 'Alavancagem (x)', key: 'alavancagem', step: '0.01' },
    { label: 'CET Mensal (%)', key: 'cetMensal', step: '0.0001' },
    { label: 'CET Anual (%)', key: 'cetAnual', step: '0.01' },
    { label: 'Parcela F1 (R$)', key: 'parcelaF1', step: '0.01' },
    { label: 'Parcela F2 (R$)', key: 'parcelaF2', step: '0.01' },
  ];

  return (
    <PageLayout title="Gerador de Propostas" subtitle="Gere apresentacoes PPTX profissionais">
      <div className="max-w-6xl mx-auto">
        <div className="mb-4">
          <button onClick={() => setPage('dashboard')} className="inline-flex items-center gap-1.5 text-sm text-somus-text-secondary hover:text-somus-text-primary transition-colors">
            <ArrowLeft className="h-4 w-4" /> Voltar ao Dashboard
          </button>
        </div>

        {/* Subtipo Selector */}
        <div className="mb-5">
          <label className="text-sm font-medium text-somus-text-primary mb-2 block">Subtipo</label>
          <div className="flex flex-wrap gap-2">
            {SUBTIPOS.map((s) => (
              <button
                key={s.value}
                type="button"
                onClick={() => setSubtipo(s.value)}
                className={`px-4 py-2 rounded-lg text-sm font-medium transition-colors border ${
                  subtipo === s.value
                    ? 'bg-somus-green text-white border-somus-green'
                    : 'bg-somus-bg-secondary text-somus-text-secondary border-somus-border hover:bg-somus-bg-hover'
                }`}
              >
                {s.label}
              </button>
            ))}
          </div>
        </div>

        {/* Tabs */}
        <Tabs tabs={TABS} activeTab={activeTab} onChange={setActiveTab}>
          {(tabId) =>
            tabId === 'proposta' ? (
              <form onSubmit={propostaForm.handleSubmit(handleGerarProposta)}>
                <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
                  {/* Client & Context */}
                  <div className="space-y-5">
                    <Card title="Dados do Cliente" padding="md">
                      <div className="space-y-4 mt-4">
                        <FormField label="Nome do Cliente" required error={propostaForm.formState.errors.clienteNome?.message}>
                          <input {...propostaForm.register('clienteNome')} className={inputClass} placeholder="Nome completo" />
                        </FormField>
                        <div className="grid grid-cols-2 gap-3">
                          <FormField label="Data (Mes/Ano)" required>
                            <input {...propostaForm.register('dataMesAno')} className={inputClass} placeholder="Marco/2026" />
                          </FormField>
                          <FormField label="Administradora" required>
                            <input {...propostaForm.register('administradora')} className={inputClass} placeholder="Embracon" />
                          </FormField>
                        </div>
                        <FormField label="Empresa">
                          <input {...propostaForm.register('empresaNome')} className={inputClass} />
                        </FormField>
                        <FormField label="Localizacao">
                          <input {...propostaForm.register('localizacao')} className={inputClass} />
                        </FormField>
                      </div>
                    </Card>
                    <Card title="Contexto e Objetivo" padding="md">
                      <div className="space-y-4 mt-4">
                        <FormField label="Descricao da Operacao">
                          <textarea {...propostaForm.register('descricaoOperacao')} rows={3} className={`${inputClass} resize-none`} />
                        </FormField>
                        <FormField label="Titulo do Objetivo">
                          <input {...propostaForm.register('objetivoTitulo')} className={inputClass} />
                        </FormField>
                        <FormField label="Descricao do Objetivo">
                          <textarea {...propostaForm.register('objetivoDescricao')} rows={3} className={`${inputClass} resize-none`} />
                        </FormField>
                        <FormField label="Horizonte">
                          <input {...propostaForm.register('horizonteTexto')} className={inputClass} placeholder="200 meses" />
                        </FormField>
                      </div>
                    </Card>
                    <Card title="Contatos" padding="md">
                      <div className="space-y-4 mt-4">
                        <div className="grid grid-cols-2 gap-3">
                          <FormField label="Contato 1 - Nome"><input {...propostaForm.register('contato1Nome')} className={inputClass} /></FormField>
                          <FormField label="Telefone"><input {...propostaForm.register('contato1Telefone')} className={inputClass} /></FormField>
                        </div>
                        <div className="grid grid-cols-2 gap-3">
                          <FormField label="Contato 2 - Nome"><input {...propostaForm.register('contato2Nome')} className={inputClass} /></FormField>
                          <FormField label="Telefone"><input {...propostaForm.register('contato2Telefone')} className={inputClass} /></FormField>
                        </div>
                      </div>
                    </Card>
                  </div>

                  {/* Cenario */}
                  <div className="space-y-5">
                    <Card title="Cenario Financeiro" padding="md">
                      <div className="space-y-4 mt-4">
                        {cenarioFields.map((f) => (
                          <FormField key={f.key} label={f.label}>
                            <Controller
                              name={`cenario.${f.key}` as any}
                              control={propostaForm.control}
                              render={({ field }) => (
                                <input type="number" step={f.step} value={field.value as number} onChange={(e) => field.onChange(Number(e.target.value))} className={inputClass} />
                              )}
                            />
                          </FormField>
                        ))}
                      </div>
                    </Card>
                    <div className="flex gap-3">
                      <Button type="submit" variant="primary" loading={generating} icon={<Download className="h-4 w-4" />} className="flex-1">Gerar PPTX</Button>
                      <Button type="button" variant="secondary" icon={<Eye className="h-4 w-4" />}>Pre-visualizar</Button>
                    </div>
                  </div>
                </div>
              </form>
            ) : (
              <form onSubmit={comparativaForm.handleSubmit(handleGerarComparativa)}>
                <div className="space-y-6">
                  <Card title="Dados do Cliente" padding="md">
                    <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4 mt-4">
                      <FormField label="Nome do Cliente" required error={comparativaForm.formState.errors.clienteNome?.message}>
                        <input {...comparativaForm.register('clienteNome')} className={inputClass} />
                      </FormField>
                      <FormField label="Data (Mes/Ano)" required>
                        <input {...comparativaForm.register('dataMesAno')} className={inputClass} placeholder="Marco/2026" />
                      </FormField>
                      <FormField label="Contato 1"><input {...comparativaForm.register('contato1Nome')} className={inputClass} /></FormField>
                      <FormField label="Telefone 1"><input {...comparativaForm.register('contato1Telefone')} className={inputClass} /></FormField>
                    </div>
                  </Card>

                  {/* Cenarios count */}
                  <div className="flex items-center gap-3">
                    <span className="text-sm font-medium text-somus-text-primary">Cenarios:</span>
                    <div className="flex items-center gap-2">
                      <button type="button" onClick={() => setNumCenarios(Math.max(1, numCenarios - 1))} className="p-1.5 rounded-md bg-somus-border/30 hover:bg-somus-border"><Minus className="h-4 w-4" /></button>
                      <span className="text-sm font-bold w-6 text-center">{numCenarios}</span>
                      <button type="button" onClick={() => setNumCenarios(Math.min(3, numCenarios + 1))} className="p-1.5 rounded-md bg-somus-border/30 hover:bg-somus-border"><Plus className="h-4 w-4" /></button>
                    </div>
                  </div>

                  <div className={`grid grid-cols-1 ${numCenarios >= 2 ? 'lg:grid-cols-2' : ''} ${numCenarios >= 3 ? 'xl:grid-cols-3' : ''} gap-5`}>
                    {Array.from({ length: numCenarios }).map((_, idx) => (
                      <Card key={idx} title={`Cenario ${idx + 1}`} padding="md" className="border-t-4 border-t-somus-green">
                        <div className="space-y-3 mt-3">
                          {cenarioFields.map((f) => (
                            <FormField key={`${idx}-${f.key}`} label={f.label}>
                              <Controller
                                name={`cenarios.${idx}.${f.key}` as any}
                                control={comparativaForm.control}
                                render={({ field }) => (
                                  <input type="number" step={f.step} value={field.value as number} onChange={(e) => field.onChange(Number(e.target.value))} className={inputClass} />
                                )}
                              />
                            </FormField>
                          ))}
                        </div>
                      </Card>
                    ))}
                  </div>

                  <div className="flex gap-3">
                    <Button type="submit" variant="primary" loading={generating} icon={<Download className="h-4 w-4" />} className="flex-1" size="lg">Gerar PPTX Comparativa</Button>
                    <Button type="button" variant="secondary" icon={<Eye className="h-4 w-4" />} size="lg">Pre-visualizar</Button>
                  </div>
                </div>
              </form>
            )
          }
        </Tabs>
      </div>
    </PageLayout>
  );
}

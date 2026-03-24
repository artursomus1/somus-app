import PptxGenJS from 'pptxgenjs';
import type { PropostaData, ComparativaData, PropostaSubtipo } from '@/types';

const COLORS = {
  green: '004D33',
  dark: '111111',
  gray: '888888',
  headerBg: 'F0F0E0',
  white: 'FFFFFF',
  lightGray: 'F5F5F5',
};

const FONT = 'DM Sans';

function formatCurrency(value: number): string {
  return `R$ ${value.toLocaleString('pt-BR', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  })}`;
}

function formatPercent(value: number): string {
  return `${value.toLocaleString('pt-BR', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  })}%`;
}

function createBasePresentation(): PptxGenJS {
  const pptx = new PptxGenJS();
  pptx.layout = 'LAYOUT_WIDE'; // 16:9
  pptx.author = 'Somus Capital';
  pptx.company = 'Somus Capital';
  return pptx;
}

function addCoverSlide(
  pptx: PptxGenJS,
  title: string,
  subtitle: string,
  clienteNome: string,
  dataMesAno: string
) {
  const slide = pptx.addSlide();

  // Background
  slide.background = { color: COLORS.green };

  // Title
  slide.addText(title, {
    x: 1,
    y: 1.5,
    w: 8,
    h: 1.2,
    fontSize: 36,
    fontFace: FONT,
    color: COLORS.white,
    bold: true,
  });

  // Subtitle
  slide.addText(subtitle, {
    x: 1,
    y: 2.8,
    w: 8,
    h: 0.6,
    fontSize: 18,
    fontFace: FONT,
    color: COLORS.white,
    italic: true,
  });

  // Client name
  slide.addText(clienteNome, {
    x: 1,
    y: 4,
    w: 8,
    h: 0.5,
    fontSize: 16,
    fontFace: FONT,
    color: COLORS.white,
  });

  // Date
  slide.addText(dataMesAno, {
    x: 1,
    y: 4.6,
    w: 8,
    h: 0.4,
    fontSize: 14,
    fontFace: FONT,
    color: COLORS.white,
  });
}

function addContactSlide(
  pptx: PptxGenJS,
  contato1Nome: string,
  contato1Telefone: string,
  contato2Nome: string,
  contato2Telefone: string
) {
  const slide = pptx.addSlide();
  slide.background = { color: COLORS.green };

  slide.addText('Contato', {
    x: 1,
    y: 1,
    w: 8,
    h: 0.8,
    fontSize: 28,
    fontFace: FONT,
    color: COLORS.white,
    bold: true,
  });

  slide.addText(
    [
      { text: contato1Nome, options: { fontSize: 18, bold: true, breakLine: true } },
      { text: contato1Telefone, options: { fontSize: 16, breakLine: true } },
      { text: '', options: { breakLine: true } },
      { text: contato2Nome, options: { fontSize: 18, bold: true, breakLine: true } },
      { text: contato2Telefone, options: { fontSize: 16 } },
    ],
    {
      x: 1,
      y: 2.5,
      w: 8,
      h: 3,
      fontFace: FONT,
      color: COLORS.white,
    }
  );
}

/**
 * Gera apresentacao PPTX de proposta comercial
 */
export async function gerarPropostaPPTX(
  data: PropostaData,
  subtipo: PropostaSubtipo
): Promise<Blob> {
  const pptx = createBasePresentation();
  pptx.title = `Proposta ${subtipo} - ${data.clienteNome}`;

  // Slide 1 - Cover
  addCoverSlide(
    pptx,
    'Proposta Comercial',
    `${subtipo.replace(/_/g, ' ')}`,
    data.clienteNome,
    data.dataMesAno
  );

  // Slide 2 - Operacao
  const slide2 = pptx.addSlide();
  slide2.addText('Detalhes da Operacao', {
    x: 0.5,
    y: 0.3,
    w: 9,
    h: 0.6,
    fontSize: 24,
    fontFace: FONT,
    color: COLORS.green,
    bold: true,
  });

  const detailRows: [string, string][] = [
    ['Empresa', data.empresaNome],
    ['Localizacao', data.localizacao],
    ['Administradora', data.administradora],
    ['Objetivo', data.objetivoTitulo],
    ['Horizonte', data.horizonteTexto],
    ['Valor da Carta', formatCurrency(data.cenario.valorCarta)],
    ['Prazo Total', `${data.cenario.prazoTotal} meses`],
    ['Taxa Adm.', formatPercent(data.cenario.taxaAdm)],
    ['Fundo Reserva', formatPercent(data.cenario.fundoReserva)],
    ['CET Anual', formatPercent(data.cenario.cetAnual)],
  ];

  const tableRows = detailRows.map(([label, value]) => [
    {
      text: label,
      options: {
        fontFace: FONT,
        fontSize: 11,
        bold: true,
        color: COLORS.dark,
        fill: { color: COLORS.headerBg },
      },
    },
    {
      text: value,
      options: {
        fontFace: FONT,
        fontSize: 11,
        color: COLORS.dark,
      },
    },
  ]);

  slide2.addTable(tableRows as any, {
    x: 0.5,
    y: 1.2,
    w: 9,
    colW: [3, 6],
    border: { type: 'solid', pt: 0.5, color: 'E5E7EB' },
    rowH: 0.4,
  });

  // Slide 3 - Parcelas
  const slide3 = pptx.addSlide();
  slide3.addText('Cronograma de Parcelas', {
    x: 0.5,
    y: 0.3,
    w: 9,
    h: 0.6,
    fontSize: 24,
    fontFace: FONT,
    color: COLORS.green,
    bold: true,
  });

  const parcelaHeader = [
    { text: 'Parcela', options: { fontFace: FONT, fontSize: 10, bold: true, color: COLORS.green, fill: { color: COLORS.headerBg } } },
    { text: 'Valor', options: { fontFace: FONT, fontSize: 10, bold: true, color: COLORS.green, fill: { color: COLORS.headerBg }, align: 'right' as const } },
  ];

  const parcelaRows = (data.cenario.parcelas || []).slice(0, 15).map((val, idx) => [
    { text: `${idx + 1}`, options: { fontFace: FONT, fontSize: 10, color: COLORS.dark } },
    { text: formatCurrency(val), options: { fontFace: FONT, fontSize: 10, color: COLORS.dark, align: 'right' as const } },
  ]);

  slide3.addTable([parcelaHeader, ...parcelaRows] as any, {
    x: 0.5,
    y: 1.2,
    w: 9,
    colW: [2, 7],
    border: { type: 'solid', pt: 0.5, color: 'E5E7EB' },
    rowH: 0.35,
  });

  // Slide 4 - Observacoes
  const slide4 = pptx.addSlide();
  slide4.addText('Observacoes', {
    x: 0.5,
    y: 0.3,
    w: 9,
    h: 0.6,
    fontSize: 24,
    fontFace: FONT,
    color: COLORS.green,
    bold: true,
  });
  slide4.addText(data.descricaoOperacao || 'Sem observacoes adicionais.', {
    x: 0.5,
    y: 1.2,
    w: 9,
    h: 4,
    fontSize: 12,
    fontFace: FONT,
    color: COLORS.dark,
    valign: 'top',
  });

  // Slide 5 - Contact
  addContactSlide(
    pptx,
    data.contato1Nome,
    data.contato1Telefone,
    data.contato2Nome,
    data.contato2Telefone
  );

  const blob = await pptx.write({ outputType: 'blob' });
  return blob as Blob;
}

/**
 * Gera apresentacao PPTX comparativa entre cenarios
 */
export async function gerarComparativaPPTX(
  data: ComparativaData,
  subtipo: PropostaSubtipo
): Promise<Blob> {
  const pptx = createBasePresentation();
  pptx.title = `Comparativa ${subtipo} - ${data.clienteNome}`;

  // Cover
  addCoverSlide(
    pptx,
    'Analise Comparativa',
    `${subtipo.replace(/_/g, ' ')}`,
    data.clienteNome,
    data.dataMesAno
  );

  // Cenarios comparison slide
  const slide2 = pptx.addSlide();
  slide2.addText('Comparativo de Cenarios', {
    x: 0.5,
    y: 0.3,
    w: 9,
    h: 0.6,
    fontSize: 24,
    fontFace: FONT,
    color: COLORS.green,
    bold: true,
  });

  const headerRow = [
    { text: 'Parametro', options: { fontFace: FONT, fontSize: 10, bold: true, color: COLORS.green, fill: { color: COLORS.headerBg } } },
    ...data.cenarios.map((_, i) => ({
      text: `Cenario ${i + 1}`,
      options: { fontFace: FONT, fontSize: 10, bold: true, color: COLORS.green, fill: { color: COLORS.headerBg } },
    })),
  ];

  const compRows = [
    ['Valor Carta', ...data.cenarios.map((c) => formatCurrency(c.valorCarta))],
    ['Prazo', ...data.cenarios.map((c) => `${c.prazoTotal} meses`)],
    ['Contemplacao', ...data.cenarios.map((c) => `Mes ${c.contemplacaoMes}`)],
    ['CET Mensal', ...data.cenarios.map((c) => formatPercent(c.cetMensal))],
    ['CET Anual', ...data.cenarios.map((c) => formatPercent(c.cetAnual))],
    ['Lance Livre', ...data.cenarios.map((c) => formatPercent(c.lanceLivrePct))],
    ['Lance Embutido', ...data.cenarios.map((c) => formatPercent(c.lanceEmbutidoPct))],
    ['Credito Disponivel', ...data.cenarios.map((c) => formatCurrency(c.creditoDisponivel))],
    ['Alavancagem', ...data.cenarios.map((c) => `${c.alavancagem.toFixed(2)}x`)],
  ].map((row) =>
    row.map((cell, idx) => ({
      text: cell,
      options: {
        fontFace: FONT,
        fontSize: 10,
        color: COLORS.dark,
        bold: idx === 0,
        fill: idx === 0 ? { color: COLORS.lightGray } : undefined,
      },
    }))
  );

  slide2.addTable([headerRow, ...compRows] as any, {
    x: 0.5,
    y: 1.2,
    w: 9,
    border: { type: 'solid', pt: 0.5, color: 'E5E7EB' },
    rowH: 0.38,
  });

  // Contact
  addContactSlide(
    pptx,
    data.contato1Nome,
    data.contato1Telefone,
    data.contato2Nome,
    data.contato2Telefone
  );

  const blob = await pptx.write({ outputType: 'blob' });
  return blob as Blob;
}

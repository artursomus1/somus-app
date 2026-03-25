/**
 * report-pdf.ts - PDF report generator using pdfmake
 * Somus Capital - Mesa de Produtos
 *
 * Generates professional PDF documents for consortium analysis reports.
 * Uses a LIGHT theme suitable for printing on paper.
 */

import pdfMake from 'pdfmake/build/pdfmake';
import pdfFonts from 'pdfmake/build/vfs_fonts';
import type { TDocumentDefinitions, Content, ContentTable, ContentColumns } from 'pdfmake/interfaces';

import {
  ReportData,
  fmtBRL,
  fmtPct,
  fmtNum,
  fmtMes,
  safeGet,
  getTotais,
  getMetricas,
  getFluxoMensal,
} from './report-types';

// @ts-ignore
pdfMake.vfs = pdfFonts.pdfMake ? pdfFonts.pdfMake.vfs : pdfFonts.vfs;

// =============================================================================
// BRAND COLORS (light theme for print)
// =============================================================================

const PDF_GREEN = '#004D33';
const PDF_LIGHT_GREEN = '#E8F5EE';
const PDF_GRAY = '#F3F4F6';
const PDF_DARK = '#1F2937';
const PDF_TEXT = '#374151';
const PDF_GOLD = '#B45309';
const PDF_PURPLE = '#7030A0';
const PDF_RED = '#DC2626';
const PDF_HEADER_BG = '#004D33';
const PDF_HEADER_FG = '#FFFFFF';

// =============================================================================
// SHARED HELPERS
// =============================================================================

function somusHeader(title: string, subtitle?: string): Content {
  const cols: any[] = [
    {
      text: 'SOMUS CAPITAL',
      fontSize: 16,
      bold: true,
      color: PDF_HEADER_FG,
    },
  ];

  const rightCol: any = {
    text: [
      { text: title, fontSize: 12, bold: true },
      ...(subtitle ? [{ text: '\n' + subtitle, fontSize: 9 }] : []),
    ],
    color: PDF_HEADER_FG,
    alignment: 'right',
  };
  cols.push(rightCol);

  return {
    table: {
      widths: ['*', '*'],
      body: [[...cols]],
    },
    layout: {
      fillColor: () => PDF_HEADER_BG,
      hLineWidth: () => 0,
      vLineWidth: () => 0,
      paddingLeft: () => 16,
      paddingRight: () => 16,
      paddingTop: () => 12,
      paddingBottom: () => 12,
    },
    margin: [0, 0, 0, 16] as [number, number, number, number],
  } as Content;
}

function sectionTitle(text: string): Content {
  return {
    text,
    fontSize: 13,
    bold: true,
    color: PDF_GREEN,
    margin: [0, 12, 0, 8] as [number, number, number, number],
    decoration: 'underline' as any,
    decorationColor: PDF_LIGHT_GREEN,
  };
}

function kpiRow(items: { label: string; value: string; color?: string }[]): Content {
  const cols = items.map((item) => ({
    stack: [
      { text: item.label, fontSize: 8, color: '#6B7280', margin: [0, 0, 0, 2] as [number, number, number, number] },
      { text: item.value, fontSize: 12, bold: true, color: item.color || PDF_DARK },
    ],
    width: '*' as const,
    margin: [0, 0, 8, 0] as [number, number, number, number],
  }));

  return {
    columns: cols,
    margin: [0, 0, 0, 12] as [number, number, number, number],
    columnGap: 8,
  };
}

function dataTable(
  headers: string[],
  rows: any[][],
  options?: {
    widths?: (string | number)[];
    headerColor?: string;
    stripeColor?: string;
    fontSize?: number;
    rightAlignFrom?: number;
  }
): Content {
  const opts = {
    headerColor: PDF_GREEN,
    stripeColor: PDF_GRAY,
    fontSize: 8,
    rightAlignFrom: 1,
    ...options,
  };

  const headerRow = headers.map((h, i) => ({
    text: h,
    bold: true,
    fontSize: opts.fontSize,
    color: '#FFFFFF',
    alignment: (i >= (opts.rightAlignFrom ?? 999) ? 'right' : 'left') as any,
  }));

  const bodyRows = rows.map((row) =>
    row.map((cell, i) => ({
      text: String(cell ?? ''),
      fontSize: opts.fontSize,
      color: PDF_TEXT,
      alignment: (i >= (opts.rightAlignFrom ?? 999) ? 'right' : 'left') as any,
    }))
  );

  const widths = opts.widths || headers.map(() => '*');

  return {
    table: {
      headerRows: 1,
      widths,
      body: [headerRow, ...bodyRows],
    },
    layout: {
      hLineWidth: (i: number) => (i <= 1 ? 1 : 0.5),
      vLineWidth: () => 0,
      hLineColor: (i: number) => (i <= 1 ? opts.headerColor : '#E5E7EB'),
      fillColor: (rowIndex: number) => {
        if (rowIndex === 0) return opts.headerColor;
        return rowIndex % 2 === 0 ? opts.stripeColor : null;
      },
      paddingLeft: () => 6,
      paddingRight: () => 6,
      paddingTop: () => 4,
      paddingBottom: () => 4,
    },
    margin: [0, 0, 0, 12] as [number, number, number, number],
  } as Content;
}

function labelValueTable(items: { label: string; value: string }[]): Content {
  const body = items.map((item) => [
    { text: item.label, fontSize: 9, bold: true, color: PDF_DARK },
    { text: item.value, fontSize: 9, color: PDF_TEXT, alignment: 'right' as any },
  ]);

  return {
    table: {
      widths: ['*', 'auto'],
      body,
    },
    layout: {
      hLineWidth: () => 0.5,
      vLineWidth: () => 0,
      hLineColor: () => '#E5E7EB',
      fillColor: (i: number) => (i % 2 === 0 ? PDF_GRAY : null),
      paddingLeft: () => 8,
      paddingRight: () => 8,
      paddingTop: () => 4,
      paddingBottom: () => 4,
    },
    margin: [0, 0, 0, 12] as [number, number, number, number],
  } as Content;
}

function valorBadge(criaValor: boolean): Content {
  return {
    table: {
      widths: ['*'],
      body: [[
        {
          text: criaValor ? 'CRIA VALOR' : 'DESTROI VALOR',
          fontSize: 14,
          bold: true,
          color: '#FFFFFF',
          alignment: 'center',
        },
      ]],
    },
    layout: {
      fillColor: () => criaValor ? PDF_GREEN : PDF_RED,
      hLineWidth: () => 0,
      vLineWidth: () => 0,
      paddingTop: () => 10,
      paddingBottom: () => 10,
      paddingLeft: () => 0,
      paddingRight: () => 0,
    },
    margin: [0, 0, 0, 16] as [number, number, number, number],
  } as Content;
}

function buildDoc(content: Content[], title: string): TDocumentDefinitions {
  return {
    pageSize: 'A4',
    pageMargins: [36, 36, 36, 50],
    footer: (currentPage: number, pageCount: number) => ({
      columns: [
        {
          text: `Somus Capital - ${title}`,
          fontSize: 7,
          color: '#9CA3AF',
          margin: [36, 0, 0, 0],
        },
        {
          text: `Gerado em ${new Date().toLocaleDateString('pt-BR')} | Pagina ${currentPage} de ${pageCount}`,
          fontSize: 7,
          color: '#9CA3AF',
          alignment: 'right',
          margin: [0, 0, 36, 0],
        },
      ],
    }),
    content,
    defaultStyle: {
      font: 'Roboto',
      fontSize: 9,
      color: PDF_TEXT,
    },
  };
}

// =============================================================================
// 1. RESUMO EXECUTIVO
// =============================================================================

export function gerarPDFResumoExecutivo(data: ReportData): void {
  const p = data.params;
  const t = getTotais(data.fluxo);
  const m = getMetricas(data.fluxo);
  const v = data.vpl || {};

  const valorCredito = safeGet(p, 'valor_credito', safeGet(p, 'valor_carta', 0));
  const cartaLiquida = safeGet(t, 'carta_liquida', 0);
  const cetAnual = safeGet(m, 'cet_anual', safeGet(v, 'cet_anual', 0));
  const tirAnual = safeGet(m, 'tir_anual', safeGet(v, 'tir_anual', 0));
  const deltaVpl = safeGet(v, 'delta_vpl', 0);
  const breakEven = safeGet(v, 'break_even_lance', 0);
  const totalPago = safeGet(t, 'total_pago', 0);
  const custoTotalPct = safeGet(m, 'custo_total_pct', 0);
  const criaValor = safeGet(v, 'cria_valor', false);

  const content: Content[] = [
    somusHeader('Resumo Executivo', `${data.clienteNome} | ${data.assessor}`),
    valorBadge(criaValor),

    // KPI Row 1
    kpiRow([
      { label: 'Valor Credito', value: fmtBRL(valorCredito) },
      { label: 'Carta Liquida', value: fmtBRL(cartaLiquida) },
      { label: 'CET Anual', value: fmtPct(cetAnual * 100) },
      { label: 'TIR Anual', value: fmtPct(tirAnual * 100) },
    ]),
    // KPI Row 2
    kpiRow([
      { label: 'Delta VPL', value: fmtBRL(deltaVpl), color: deltaVpl >= 0 ? PDF_GREEN : PDF_RED },
      { label: 'Break-even Lance', value: fmtPct(breakEven * 100) },
      { label: 'Total Desembolsado', value: fmtBRL(totalPago) },
      { label: 'Custo Total %', value: fmtPct(custoTotalPct * 100) },
    ]),

    // Dados da Operacao
    sectionTitle('Dados da Operacao'),
    labelValueTable([
      { label: 'Cliente', value: data.clienteNome },
      { label: 'Assessor', value: data.assessor },
      { label: 'Administradora', value: data.administradora },
      { label: 'Data', value: data.dataGeracao },
      { label: 'Valor Credito', value: fmtBRL(valorCredito) },
      { label: 'Prazo (meses)', value: fmtMes(safeGet(p, 'prazo_meses', 0)) },
      { label: 'Contemplacao (mes)', value: fmtMes(safeGet(p, 'momento_contemplacao', safeGet(p, 'prazo_contemp', 0))) },
      { label: 'Taxa Administracao', value: fmtPct(safeGet(p, 'taxa_adm_pct', safeGet(p, 'taxa_adm', 0))) },
      { label: 'Fundo Reserva', value: fmtPct(safeGet(p, 'fundo_reserva_pct', safeGet(p, 'fundo_reserva', 0))) },
      { label: 'Lance Embutido', value: fmtPct(safeGet(p, 'lance_embutido_pct', 0)) },
      { label: 'Lance Livre', value: fmtPct(safeGet(p, 'lance_livre_pct', 0)) },
      { label: 'Reajuste Pre', value: fmtPct(safeGet(p, 'reajuste_pre_pct', 0)) + ' (' + safeGet(p, 'reajuste_pre_freq', 'Anual') + ')' },
      { label: 'Reajuste Pos', value: fmtPct(safeGet(p, 'reajuste_pos_pct', 0)) + ' (' + safeGet(p, 'reajuste_pos_freq', 'Anual') + ')' },
      { label: 'Seguro Vida', value: fmtPct(safeGet(p, 'seguro_vida_pct', safeGet(p, 'seguro', 0))) },
      { label: 'TMA', value: fmtPct(safeGet(p, 'tma', 0) * 100) },
      { label: 'ALM Anual', value: fmtPct(safeGet(p, 'alm_anual', 0)) },
      { label: 'Hurdle Anual', value: fmtPct(safeGet(p, 'hurdle_anual', 0)) },
    ]),

    // Composicao de Custos
    sectionTitle('Composicao de Custos'),
    dataTable(
      ['Componente', 'Valor Total', '% do Credito'],
      [
        ['Fundo Comum', fmtBRL(safeGet(t, 'total_fundo_comum', 0)), fmtPct(valorCredito > 0 ? (safeGet(t, 'total_fundo_comum', 0) / valorCredito) * 100 : 0)],
        ['Taxa Administracao', fmtBRL(safeGet(t, 'total_taxa_adm', 0)), fmtPct(safeGet(p, 'taxa_adm_pct', safeGet(p, 'taxa_adm', 0)))],
        ['Fundo Reserva', fmtBRL(safeGet(t, 'total_fundo_reserva', 0)), fmtPct(safeGet(p, 'fundo_reserva_pct', safeGet(p, 'fundo_reserva', 0)))],
        ['Seguro', fmtBRL(safeGet(t, 'total_seguro', 0)), fmtPct(valorCredito > 0 ? (safeGet(t, 'total_seguro', 0) / valorCredito) * 100 : 0)],
        ['Custos Acessorios', fmtBRL(safeGet(t, 'total_custos_acessorios', 0)), fmtPct(valorCredito > 0 ? (safeGet(t, 'total_custos_acessorios', 0) / valorCredito) * 100 : 0)],
        [
          { text: 'TOTAL', bold: true } as any,
          { text: fmtBRL(totalPago), bold: true } as any,
          { text: fmtPct(custoTotalPct * 100), bold: true } as any,
        ],
      ],
      { widths: ['*', 'auto', 'auto'], rightAlignFrom: 1 }
    ),

    // Parcelas
    sectionTitle('Resumo de Parcelas'),
    kpiRow([
      { label: 'Parcela Media', value: fmtBRL(safeGet(m, 'parcela_media', 0)) },
      { label: 'Parcela Maxima', value: fmtBRL(safeGet(m, 'parcela_maxima', 0)) },
      { label: 'Parcela Minima', value: fmtBRL(safeGet(m, 'parcela_minima', 0)) },
    ]),
  ];

  const doc = buildDoc(content, 'Resumo Executivo');
  pdfMake.createPdf(doc).download(`Resumo_Executivo_${data.clienteNome.replace(/\s+/g, '_')}.pdf`);
}

// =============================================================================
// 2. FLUXO FINANCEIRO
// =============================================================================

export function gerarPDFFluxoFinanceiro(data: ReportData): void {
  const p = data.params;
  const t = getTotais(data.fluxo);
  const fluxoMensal = getFluxoMensal(data.fluxo);

  const valorCredito = safeGet(p, 'valor_credito', safeGet(p, 'valor_carta', 0));

  const content: Content[] = [
    somusHeader('Fluxo Financeiro', `${data.clienteNome} | ${data.administradora}`),

    // Params summary
    kpiRow([
      { label: 'Valor Credito', value: fmtBRL(valorCredito) },
      { label: 'Prazo', value: fmtMes(safeGet(p, 'prazo_meses', 0)) + ' meses' },
      { label: 'Contemplacao', value: 'Mes ' + fmtMes(safeGet(p, 'momento_contemplacao', safeGet(p, 'prazo_contemp', 0))) },
      { label: 'Total Pago', value: fmtBRL(safeGet(t, 'total_pago', 0)) },
    ]),

    sectionTitle('Fluxo Mensal Completo'),
  ];

  // Build table rows from fluxo mensal
  const headers = ['Mes', 'Amortizacao', 'Saldo Princ.', 'Tx Adm', 'Fdo Reserva', 'Parcela Base', 'Reajuste', 'Parc. Reaj.', 'Seguro', 'Outros', 'Total', 'Fluxo Cx'];
  const rows: any[][] = [];

  for (const f of fluxoMensal) {
    rows.push([
      fmtMes(f.mes),
      fmtBRL(f.amortizacao ?? 0),
      fmtBRL(f.saldo_principal ?? f.saldo_devedor ?? 0),
      fmtBRL(f.valor_parcela_ta ?? f.taxa_adm_antecipada ?? 0),
      fmtBRL(f.valor_fundo_reserva ?? 0),
      fmtBRL(f.valor_parcela ?? 0),
      fmtPct((f.pct_reajuste ?? 0) * 100),
      fmtBRL(f.parcela_apos_reajuste ?? 0),
      fmtBRL(f.seguro_vida ?? 0),
      fmtBRL(f.outros_custos ?? 0),
      fmtBRL(f.parcela_com_seguro ?? f.parcela_apos_reajuste ?? 0),
      fmtBRL(f.fluxo_caixa ?? 0),
    ]);
  }

  // Totals row
  rows.push([
    { text: 'TOTAL', bold: true },
    '',
    '',
    '',
    '',
    '',
    '',
    '',
    { text: fmtBRL(safeGet(t, 'total_seguro', 0)), bold: true },
    { text: fmtBRL(safeGet(t, 'total_custos_acessorios', 0)), bold: true },
    { text: fmtBRL(safeGet(t, 'total_pago', 0)), bold: true },
    '',
  ]);

  content.push(
    dataTable(headers, rows, {
      widths: ['auto', 'auto', 'auto', 'auto', 'auto', 'auto', 'auto', 'auto', 'auto', 'auto', 'auto', 'auto'],
      fontSize: 6,
      rightAlignFrom: 1,
    })
  );

  const doc = buildDoc(content, 'Fluxo Financeiro');
  doc.pageOrientation = 'landscape';
  pdfMake.createPdf(doc).download(`Fluxo_Financeiro_${data.clienteNome.replace(/\s+/g, '_')}.pdf`);
}

// =============================================================================
// 3. ANALISE VPL
// =============================================================================

export function gerarPDFAnaliseVPL(data: ReportData): void {
  const v = data.vpl || {};
  const p = data.params;
  const valorCredito = safeGet(p, 'valor_credito', safeGet(p, 'valor_carta', 0));
  const criaValor = safeGet(v, 'cria_valor', false);

  const content: Content[] = [
    somusHeader('Analise de VPL', `${data.clienteNome} | NASA Engine HD`),
    valorBadge(criaValor),

    // Decomposicao
    sectionTitle('Decomposicao do VPL'),
    dataTable(
      ['Componente', 'Valor', '% do Credito'],
      [
        ['B0 (VP Beneficio)', fmtBRL(safeGet(v, 'b0', 0)), fmtPct(valorCredito > 0 ? (safeGet(v, 'b0', 0) / valorCredito) * 100 : 0)],
        ['H0 (VP Custos pre-T)', fmtBRL(safeGet(v, 'h0', 0)), fmtPct(valorCredito > 0 ? (safeGet(v, 'h0', 0) / valorCredito) * 100 : 0)],
        ['D0 (VP Custos pos-T)', fmtBRL(safeGet(v, 'd0', 0)), fmtPct(valorCredito > 0 ? (safeGet(v, 'd0', 0) / valorCredito) * 100 : 0)],
        ['PV pos-T (na contemp)', fmtBRL(safeGet(v, 'pv_pos_t', 0)), fmtPct(valorCredito > 0 ? (safeGet(v, 'pv_pos_t', 0) / valorCredito) * 100 : 0)],
        [
          { text: 'Delta VPL', bold: true } as any,
          { text: fmtBRL(safeGet(v, 'delta_vpl', 0)), bold: true, color: safeGet(v, 'delta_vpl', 0) >= 0 ? PDF_GREEN : PDF_RED } as any,
          { text: fmtPct(valorCredito > 0 ? (safeGet(v, 'delta_vpl', 0) / valorCredito) * 100 : 0), bold: true } as any,
        ],
      ],
      { widths: ['*', 'auto', 'auto'], rightAlignFrom: 1 }
    ),

    // Parametros
    sectionTitle('Parametros da Analise'),
    kpiRow([
      { label: 'ALM Anual', value: fmtPct(safeGet(p, 'alm_anual', 0)) },
      { label: 'Hurdle Anual', value: fmtPct(safeGet(p, 'hurdle_anual', 0)) },
      { label: 'Break-even Lance', value: fmtPct(safeGet(v, 'break_even_lance', 0) * 100) },
      { label: 'TMA', value: fmtPct(safeGet(p, 'tma', 0) * 100) },
    ]),
  ];

  // VP detail table
  const pvPreDetail: any[] = safeGet(v, 'pv_pre_t_detail', []);
  const pvPosDetail: any[] = safeGet(v, 'pv_pos_t_detail', []);

  if (pvPreDetail.length > 0 || pvPosDetail.length > 0) {
    content.push(sectionTitle('VP por Periodo'));

    const maxLen = Math.max(pvPreDetail.length, pvPosDetail.length);
    const vpRows: any[][] = [];
    for (let i = 0; i < maxLen; i++) {
      const pre = pvPreDetail[i] || {};
      const pos = pvPosDetail[i] || {};
      vpRows.push([
        fmtMes(pre.mes ?? pos.mes ?? i),
        fmtBRL(pre.valor ?? 0),
        fmtBRL(pre.pv ?? 0),
        fmtBRL(pos.valor ?? 0),
        fmtBRL(pos.pv ?? 0),
      ]);
    }

    content.push(
      dataTable(
        ['Mes', 'Pgto Pre-T', 'VP Pre-T', 'Pgto Pos-T', 'VP Pos-T'],
        vpRows,
        { widths: ['auto', '*', '*', '*', '*'], fontSize: 7, rightAlignFrom: 1 }
      )
    );
  }

  const doc = buildDoc(content, 'Analise VPL');
  pdfMake.createPdf(doc).download(`Analise_VPL_${data.clienteNome.replace(/\s+/g, '_')}.pdf`);
}

// =============================================================================
// 4. COMPARATIVO (Consorcio vs Financiamento)
// =============================================================================

export function gerarPDFComparativo(data: ReportData): void {
  const comp = data.comparativo || {};
  const p = data.params;

  const content: Content[] = [
    somusHeader('Comparativo', `Consorcio vs Financiamento | ${data.clienteNome}`),

    // Side-by-side KPIs
    sectionTitle('Indicadores Principais'),
    kpiRow([
      { label: 'Total Pago Consorcio', value: fmtBRL(safeGet(comp, 'total_pago_consorcio', 0)), color: PDF_GREEN },
      { label: 'Total Pago Financiamento', value: fmtBRL(safeGet(comp, 'total_pago_financiamento', 0)), color: PDF_PURPLE },
      { label: 'Economia Nominal', value: fmtBRL(safeGet(comp, 'economia_nominal', 0)), color: safeGet(comp, 'economia_nominal', 0) > 0 ? PDF_GREEN : PDF_RED },
    ]),
    kpiRow([
      { label: 'VPL Consorcio', value: fmtBRL(safeGet(comp, 'vpl_consorcio', 0)), color: PDF_GREEN },
      { label: 'VPL Financiamento', value: fmtBRL(safeGet(comp, 'vpl_financiamento', 0)), color: PDF_PURPLE },
      { label: 'Economia VPL', value: fmtBRL(safeGet(comp, 'economia_vpl', 0)), color: safeGet(comp, 'economia_vpl', 0) > 0 ? PDF_GREEN : PDF_RED },
    ]),

    // Comparison table
    sectionTitle('Tabela Comparativa'),
    dataTable(
      ['Indicador', 'Consorcio', 'Financiamento', 'Diferenca'],
      [
        [
          'Total Pago',
          fmtBRL(safeGet(comp, 'total_pago_consorcio', 0)),
          fmtBRL(safeGet(comp, 'total_pago_financiamento', 0)),
          fmtBRL(safeGet(comp, 'economia_nominal', 0)),
        ],
        [
          'VPL (ALM)',
          fmtBRL(safeGet(comp, 'vpl_consorcio', 0)),
          fmtBRL(safeGet(comp, 'vpl_financiamento', 0)),
          fmtBRL(safeGet(comp, 'economia_vpl', 0)),
        ],
        [
          'VPL (TMA)',
          fmtBRL(safeGet(comp, 'vpl_consorcio_tma', 0)),
          fmtBRL(safeGet(comp, 'vpl_financiamento_tma', 0)),
          fmtBRL(safeGet(comp, 'economia_vpl_tma', 0)),
        ],
        [
          'TIR Mensal',
          fmtPct(safeGet(comp, 'tir_consorcio_mensal', 0) * 100),
          fmtPct(safeGet(comp, 'tir_financ_mensal', 0) * 100),
          '-',
        ],
        [
          'TIR Anual',
          fmtPct(safeGet(comp, 'tir_consorcio_anual', 0) * 100),
          fmtPct(safeGet(comp, 'tir_financ_anual', 0) * 100),
          '-',
        ],
        [
          'Razao Custo/Credito',
          fmtNum(safeGet(comp, 'razao_vpl_consorcio', 0), 4),
          fmtNum(safeGet(comp, 'razao_vpl_financ', 0), 4),
          '-',
        ],
      ],
      { widths: ['*', 'auto', 'auto', 'auto'], rightAlignFrom: 1 }
    ),
  ];

  // Nominal flows side by side
  const pvCons: number[] = safeGet(comp, 'pv_consorcio', []);
  const pvFin: number[] = safeGet(comp, 'pv_financiamento', []);
  if (pvCons.length > 0 || pvFin.length > 0) {
    content.push(sectionTitle('Valor Presente dos Fluxos'));
    const maxLen = Math.max(pvCons.length, pvFin.length);
    const pvRows: any[][] = [];
    for (let i = 0; i < Math.min(maxLen, 60); i++) {
      pvRows.push([
        fmtMes(i),
        fmtBRL(pvCons[i] ?? 0),
        fmtBRL(pvFin[i] ?? 0),
        fmtBRL((pvCons[i] ?? 0) - (pvFin[i] ?? 0)),
      ]);
    }
    content.push(
      dataTable(
        ['Mes', 'VP Consorcio', 'VP Financiamento', 'Diferenca'],
        pvRows,
        { widths: ['auto', '*', '*', '*'], fontSize: 7, rightAlignFrom: 1 }
      )
    );
  }

  const doc = buildDoc(content, 'Comparativo');
  pdfMake.createPdf(doc).download(`Comparativo_${data.clienteNome.replace(/\s+/g, '_')}.pdf`);
}

// =============================================================================
// 5. VENDA DA OPERACAO
// =============================================================================

export function gerarPDFVendaOperacao(data: ReportData): void {
  const vd = data.venda || {};

  const ganhoNominal = safeGet(vd, 'ganho_nominal', 0);
  const lucro = ganhoNominal >= 0;

  const content: Content[] = [
    somusHeader('Venda da Operacao', data.clienteNome),

    // Badge
    {
      table: {
        widths: ['*'],
        body: [[{
          text: lucro ? `LUCRO: ${fmtBRL(ganhoNominal)}` : `PREJUIZO: ${fmtBRL(ganhoNominal)}`,
          fontSize: 14,
          bold: true,
          color: '#FFFFFF',
          alignment: 'center',
        }]],
      },
      layout: {
        fillColor: () => lucro ? PDF_GREEN : PDF_RED,
        hLineWidth: () => 0,
        vLineWidth: () => 0,
        paddingTop: () => 10,
        paddingBottom: () => 10,
        paddingLeft: () => 0,
        paddingRight: () => 0,
      },
      margin: [0, 0, 0, 16] as [number, number, number, number],
    } as Content,

    // KPIs
    kpiRow([
      { label: 'VPL Vendedor', value: fmtBRL(safeGet(vd, 'vpl_vendedor', 0)) },
      { label: 'Ganho %', value: fmtPct(safeGet(vd, 'ganho_pct', 0)) },
      { label: 'Prazo Medio', value: fmtNum(safeGet(vd, 'prazo_medio', 0), 1) + ' meses' },
    ]),
    kpiRow([
      { label: 'Ganho Mensal', value: fmtBRL(safeGet(vd, 'ganho_mensal', 0)) },
      { label: 'Margem Mensal', value: fmtPct(safeGet(vd, 'margem_mensal_pct', 0)) },
      { label: 'Total Investido', value: fmtBRL(safeGet(vd, 'total_investido', 0)) },
    ]),
    kpiRow([
      { label: 'Valor Venda', value: fmtBRL(safeGet(vd, 'valor_venda', 0)), color: PDF_GOLD },
      { label: 'TIR Vendedor a.a.', value: fmtPct(safeGet(vd, 'tir_vendedor_anual', 0) * 100) },
      { label: 'TIR Comprador a.a.', value: fmtPct(safeGet(vd, 'tir_comprador_anual', 0) * 100), color: PDF_PURPLE },
    ]),

    // Seller flow
    sectionTitle('Fluxo do Vendedor'),
  ];

  const cfVendedor: number[] = safeGet(vd, 'cashflow_vendedor', []);
  if (cfVendedor.length > 0) {
    const vendRows = cfVendedor.map((v, i) => [fmtMes(i), fmtBRL(v)]);
    content.push(
      dataTable(['Mes', 'Fluxo de Caixa'], vendRows, { widths: ['auto', '*'], fontSize: 8, rightAlignFrom: 1 })
    );
  }

  // Buyer flow
  content.push(sectionTitle('Fluxo do Comprador'));
  const cfComprador: number[] = safeGet(vd, 'cashflow_comprador', []);
  if (cfComprador.length > 0) {
    const compRows = cfComprador.map((v, i) => [fmtMes(i), fmtBRL(v)]);
    content.push(
      dataTable(['Mes', 'Fluxo de Caixa'], compRows, { widths: ['auto', '*'], fontSize: 8, rightAlignFrom: 1 })
    );
  }

  content.push(
    kpiRow([
      { label: 'VPL Comprador', value: fmtBRL(safeGet(vd, 'vpl_comprador', 0)), color: PDF_PURPLE },
    ])
  );

  const doc = buildDoc(content, 'Venda da Operacao');
  pdfMake.createPdf(doc).download(`Venda_Operacao_${data.clienteNome.replace(/\s+/g, '_')}.pdf`);
}

// =============================================================================
// 6. CREDITO PARA LANCE
// =============================================================================

export function gerarPDFCreditoLance(data: ReportData): void {
  const cl = data.creditoLance || {};
  const cc = data.custoCombinado;

  const content: Content[] = [
    somusHeader('Credito para Lance', data.clienteNome),

    // Financing params
    sectionTitle('Parametros do Financiamento'),
    kpiRow([
      { label: 'Valor Financiado', value: fmtBRL(safeGet(cl, 'valor', 0)) },
      { label: 'Total Pago', value: fmtBRL(safeGet(cl, 'total_pago', 0)) },
      { label: 'Total Juros', value: fmtBRL(safeGet(cl, 'total_juros', 0)), color: PDF_RED },
    ]),
    kpiRow([
      { label: 'IOF', value: fmtBRL(safeGet(cl, 'iof', 0)) },
      { label: 'CET Anual', value: fmtPct(safeGet(cl, 'cet_anual', 0) * 100) },
      { label: 'TIR Anual', value: fmtPct(safeGet(cl, 'tir_anual', 0) * 100) },
    ]),
  ];

  if (safeGet(cl, 'custos_iniciais', 0) > 0) {
    content.push(
      kpiRow([
        { label: 'Custos Iniciais (TAC + Garantia + Comissao)', value: fmtBRL(safeGet(cl, 'custos_iniciais', 0)) },
      ])
    );
  }

  if (safeGet(cl, 'valor_antecipacao', 0) > 0) {
    content.push(
      kpiRow([
        { label: 'Antecipacao no Mes', value: fmtMes(safeGet(cl, 'mes_antecipacao', 0)) },
        { label: 'Valor Antecipacao', value: fmtBRL(safeGet(cl, 'valor_antecipacao', 0)) },
      ])
    );
  }

  // Amortization table
  const parcelas: any[] = safeGet(cl, 'parcelas', []);
  if (parcelas.length > 0) {
    content.push(sectionTitle('Tabela de Amortizacao'));
    const amortRows = parcelas.map((row: any) => [
      fmtMes(row.mes),
      fmtBRL(row.parcela ?? 0),
      fmtBRL(row.juros ?? 0),
      fmtBRL(row.amortizacao ?? 0),
      fmtBRL(row.saldo ?? 0),
    ]);
    content.push(
      dataTable(
        ['Mes', 'Parcela', 'Juros', 'Amortizacao', 'Saldo'],
        amortRows,
        { widths: ['auto', '*', '*', '*', '*'], fontSize: 7, rightAlignFrom: 1 }
      )
    );
  }

  // Combined cost section
  if (cc) {
    content.push(sectionTitle('Custo Combinado (Consorcio + Financiamento Lance)'));
    content.push(
      kpiRow([
        { label: 'Total Consorcio', value: fmtBRL(safeGet(cc, 'total_pago_consorcio', 0)), color: PDF_GREEN },
        { label: 'Total Lance (Fin.)', value: fmtBRL(safeGet(cc, 'total_pago_lance', 0)), color: PDF_PURPLE },
        { label: 'Total Combinado', value: fmtBRL(safeGet(cc, 'total_pago_combinado', 0)), color: PDF_GOLD },
      ])
    );
    content.push(
      kpiRow([
        { label: 'TIR Combinada Mensal', value: fmtPct(safeGet(cc, 'tir_mensal_combinado', 0) * 100) },
        { label: 'TIR Combinada Anual', value: fmtPct(safeGet(cc, 'tir_anual_combinado', 0) * 100) },
        { label: 'CET Combinado Anual', value: fmtPct(safeGet(cc, 'cet_anual_combinado', 0) * 100) },
      ])
    );
  }

  const doc = buildDoc(content, 'Credito para Lance');
  pdfMake.createPdf(doc).download(`Credito_Lance_${data.clienteNome.replace(/\s+/g, '_')}.pdf`);
}

// =============================================================================
// GERAR TODOS
// =============================================================================

export function gerarTodosPDFs(data: ReportData): void {
  gerarPDFResumoExecutivo(data);

  setTimeout(() => gerarPDFFluxoFinanceiro(data), 500);
  setTimeout(() => gerarPDFAnaliseVPL(data), 1000);

  if (data.comparativo) {
    setTimeout(() => gerarPDFComparativo(data), 1500);
  }
  if (data.venda) {
    setTimeout(() => gerarPDFVendaOperacao(data), 2000);
  }
  if (data.creditoLance) {
    setTimeout(() => gerarPDFCreditoLance(data), 2500);
  }
}

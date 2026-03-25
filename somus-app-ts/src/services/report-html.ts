/**
 * report-html.ts - Standalone HTML report generator
 * Somus Capital - Mesa de Produtos
 *
 * Generates self-contained .html files with ALL styles inlined.
 * - Dark theme for screen viewing (matching the app)
 * - Light theme for printing via @media print
 * - No external CSS/JS dependencies (except Google Fonts)
 */

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

// =============================================================================
// DOWNLOAD HELPER
// =============================================================================

export function downloadHTML(html: string, filename: string): void {
  const blob = new Blob([html], { type: 'text/html;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}

// =============================================================================
// HTML BUILDER
// =============================================================================

function buildHTML(title: string, content: string): string {
  return `<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>${title} - Somus Capital</title>
  <style>
    @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;700&display=swap');

    * { margin: 0; padding: 0; box-sizing: border-box; }
    body {
      font-family: 'DM Sans', sans-serif;
      background: #0A0F14;
      color: #E8ECF0;
      padding: 24px;
      max-width: 1200px;
      margin: 0 auto;
    }

    .header {
      background: linear-gradient(135deg, #1A7A3E, #0D5C2C);
      padding: 24px 32px;
      border-radius: 12px;
      margin-bottom: 24px;
      display: flex;
      justify-content: space-between;
      align-items: center;
    }
    .header h1 { color: white; font-size: 20px; font-weight: 700; }
    .header .subtitle { color: rgba(255,255,255,0.7); font-size: 12px; }
    .header .brand { color: rgba(255,255,255,0.5); font-size: 11px; }

    .kpi-grid {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
      gap: 12px;
      margin-bottom: 24px;
    }
    .kpi-card {
      background: #0F1419;
      border: 1px solid #1E2A3A;
      border-radius: 10px;
      padding: 16px;
      border-left: 3px solid #1A7A3E;
    }
    .kpi-card.purple { border-left-color: #7030A0; }
    .kpi-card.gold { border-left-color: #D4A017; }
    .kpi-card.red { border-left-color: #C00000; }
    .kpi-card .label { font-size: 11px; color: #8B95A5; margin-bottom: 4px; }
    .kpi-card .value { font-size: 20px; font-weight: 700; }

    .section { margin-bottom: 24px; }
    .section-title {
      font-size: 14px;
      font-weight: 700;
      color: #E8ECF0;
      margin-bottom: 12px;
      padding-bottom: 8px;
      border-bottom: 1px solid #1E2A3A;
    }

    table {
      width: 100%;
      border-collapse: collapse;
      font-size: 12px;
      margin-bottom: 16px;
    }
    th {
      background: #151C24;
      color: #8B95A5;
      padding: 8px 12px;
      text-align: left;
      font-weight: 600;
      border-bottom: 2px solid #1A7A3E;
    }
    td {
      padding: 6px 12px;
      border-bottom: 1px solid #1E2A3A;
    }
    tr:nth-child(even) { background: rgba(15, 20, 25, 0.5); }
    .text-right { text-align: right; }
    .text-center { text-align: center; }
    .text-green { color: #4ADE80; }
    .text-red { color: #EF4444; }
    .text-gold { color: #D4A017; }
    .text-purple { color: #7030A0; }
    .font-bold { font-weight: 700; }

    .badge {
      display: inline-block;
      padding: 8px 24px;
      border-radius: 8px;
      font-size: 16px;
      font-weight: 700;
      margin-bottom: 20px;
    }
    .badge-green { background: rgba(26,122,62,0.2); color: #4ADE80; border: 1px solid rgba(26,122,62,0.4); }
    .badge-red { background: rgba(192,0,0,0.2); color: #EF4444; border: 1px solid rgba(192,0,0,0.4); }

    .two-col { display: grid; grid-template-columns: 1fr 1fr; gap: 16px; }
    .panel { background: #0F1419; border: 1px solid #1E2A3A; border-radius: 10px; padding: 16px; }

    .label-value-table { width: 100%; border-collapse: collapse; margin-bottom: 16px; }
    .label-value-table td { padding: 8px 12px; border-bottom: 1px solid #1E2A3A; }
    .label-value-table tr:nth-child(even) { background: rgba(15, 20, 25, 0.5); }
    .label-value-table .lbl { font-weight: 600; color: #8B95A5; width: 40%; }
    .label-value-table .val { text-align: right; font-weight: 500; }

    .totals-row td { font-weight: 700 !important; border-top: 2px solid #1A7A3E !important; }

    .footer {
      margin-top: 32px;
      padding-top: 16px;
      border-top: 1px solid #1E2A3A;
      font-size: 10px;
      color: #5A6577;
      text-align: center;
    }

    @media print {
      body { background: white; color: #1F2937; padding: 0; }
      .header { background: #004D33; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
      .kpi-card { background: #F9FAFB; border-color: #E5E7EB; }
      .kpi-card .label { color: #6B7280; }
      .kpi-card .value { color: #1F2937; }
      th { background: #F3F4F6; color: #374151; border-bottom-color: #004D33; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
      td { border-color: #E5E7EB; color: #374151; }
      tr:nth-child(even) { background: #F9FAFB; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
      .panel { background: #F9FAFB; border-color: #E5E7EB; }
      .section-title { color: #1F2937; border-color: #E5E7EB; }
      .label-value-table td { border-color: #E5E7EB; color: #374151; }
      .label-value-table .lbl { color: #6B7280; }
      .label-value-table tr:nth-child(even) { background: #F9FAFB; }
      .text-green { color: #059669; }
      .text-red { color: #DC2626; }
      .text-gold { color: #B45309; }
      .text-purple { color: #7030A0; }
      .footer { color: #9CA3AF; }
      .badge-green { background: #ECFDF5; color: #059669; border-color: #A7F3D0; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
      .badge-red { background: #FEF2F2; color: #DC2626; border-color: #FECACA; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
    }
  </style>
</head>
<body>
  ${content}
  <div class="footer">
    Somus Capital &bull; Gerado em ${new Date().toLocaleString('pt-BR')} &bull; Documento confidencial
  </div>
</body>
</html>`;
}

// =============================================================================
// HTML COMPONENT HELPERS
// =============================================================================

function htmlHeader(title: string, subtitle?: string, brand?: string): string {
  return `<div class="header">
  <div>
    <h1>${title}</h1>
    ${subtitle ? `<div class="subtitle">${subtitle}</div>` : ''}
  </div>
  <div class="brand">${brand || 'SOMUS CAPITAL'}</div>
</div>`;
}

function htmlKpi(label: string, value: string, variant?: string): string {
  const cls = variant ? ` ${variant}` : '';
  return `<div class="kpi-card${cls}">
  <div class="label">${label}</div>
  <div class="value">${value}</div>
</div>`;
}

function htmlKpiGrid(items: { label: string; value: string; variant?: string }[]): string {
  return `<div class="kpi-grid">${items.map((i) => htmlKpi(i.label, i.value, i.variant)).join('\n')}</div>`;
}

function htmlSection(title: string, content: string): string {
  return `<div class="section">
  <div class="section-title">${title}</div>
  ${content}
</div>`;
}

function htmlBadge(text: string, isPositive: boolean): string {
  return `<div class="text-center" style="margin-bottom:20px">
  <span class="badge ${isPositive ? 'badge-green' : 'badge-red'}">${text}</span>
</div>`;
}

function htmlTable(
  headers: string[],
  rows: string[][],
  options?: { rightAlignFrom?: number; totalsRow?: boolean; id?: string }
): string {
  const raf = options?.rightAlignFrom ?? 999;
  let html = `<table${options?.id ? ` id="${options.id}"` : ''}>
<thead><tr>${headers.map((h, i) => `<th${i >= raf ? ' class="text-right"' : ''}>${h}</th>`).join('')}</tr></thead>
<tbody>`;

  rows.forEach((row, ri) => {
    const isTotals = options?.totalsRow && ri === rows.length - 1;
    html += `<tr${isTotals ? ' class="totals-row"' : ''}>`;
    row.forEach((cell, ci) => {
      const align = ci >= raf ? ' class="text-right"' : '';
      const bold = isTotals ? ' class="font-bold' + (ci >= raf ? ' text-right' : '') + '"' : align;
      html += `<td${bold}>${cell}</td>`;
    });
    html += '</tr>';
  });

  html += '</tbody></table>';
  return html;
}

function htmlLabelValueTable(items: { label: string; value: string; valueClass?: string }[]): string {
  let html = '<table class="label-value-table"><tbody>';
  for (const item of items) {
    const cls = item.valueClass ? ` class="val ${item.valueClass}"` : ' class="val"';
    html += `<tr><td class="lbl">${item.label}</td><td${cls}>${item.value}</td></tr>`;
  }
  html += '</tbody></table>';
  return html;
}

// =============================================================================
// 1. RESUMO EXECUTIVO
// =============================================================================

export function gerarHTMLResumoExecutivo(data: ReportData): string {
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

  const parcelaMedia = safeGet(m, 'parcela_media', 0);
  const parcelaMax = safeGet(m, 'parcela_maxima', 0);
  const parcelaMin = safeGet(m, 'parcela_minima', 0);

  const totalFC = safeGet(t, 'total_fundo_comum', 0);
  const totalTA = safeGet(t, 'total_taxa_adm', 0);
  const totalFR = safeGet(t, 'total_fundo_reserva', 0);
  const totalSeg = safeGet(t, 'total_seguro', 0);
  const totalAcess = safeGet(t, 'total_custos_acessorios', 0);

  let body = '';
  body += htmlHeader('Resumo Executivo', `${data.clienteNome} | ${data.assessor}`, 'SOMUS CAPITAL');
  body += htmlBadge(criaValor ? 'CRIA VALOR' : 'DESTROI VALOR', criaValor);

  // KPIs Row 1
  body += htmlKpiGrid([
    { label: 'Valor Credito', value: fmtBRL(valorCredito) },
    { label: 'Carta Liquida', value: fmtBRL(cartaLiquida) },
    { label: 'CET Anual', value: fmtPct(cetAnual * 100), variant: 'gold' },
    { label: 'TIR Anual', value: fmtPct(tirAnual * 100), variant: 'purple' },
  ]);

  // KPIs Row 2
  body += htmlKpiGrid([
    { label: 'Delta VPL', value: fmtBRL(deltaVpl), variant: deltaVpl >= 0 ? '' : 'red' },
    { label: 'Break-even Lance', value: fmtPct(breakEven * 100), variant: 'gold' },
    { label: 'Total Desembolsado', value: fmtBRL(totalPago), variant: 'red' },
    { label: 'Custo Total %', value: fmtPct(custoTotalPct * 100) },
  ]);

  // Dados da Operacao
  body += htmlSection('Dados da Operacao', htmlLabelValueTable([
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
  ]));

  // Composicao de Custos
  const custosRows: string[][] = [
    ['Fundo Comum', fmtBRL(totalFC), fmtPct(valorCredito > 0 ? (totalFC / valorCredito) * 100 : 0)],
    ['Taxa Administracao', fmtBRL(totalTA), fmtPct(safeGet(p, 'taxa_adm_pct', safeGet(p, 'taxa_adm', 0)))],
    ['Fundo Reserva', fmtBRL(totalFR), fmtPct(safeGet(p, 'fundo_reserva_pct', safeGet(p, 'fundo_reserva', 0)))],
    ['Seguro', fmtBRL(totalSeg), fmtPct(valorCredito > 0 ? (totalSeg / valorCredito) * 100 : 0)],
    ['Custos Acessorios', fmtBRL(totalAcess), fmtPct(valorCredito > 0 ? (totalAcess / valorCredito) * 100 : 0)],
    ['<strong>TOTAL</strong>', `<strong>${fmtBRL(totalPago)}</strong>`, `<strong>${fmtPct(custoTotalPct * 100)}</strong>`],
  ];
  body += htmlSection('Composicao de Custos',
    htmlTable(['Componente', 'Valor Total', '% do Credito'], custosRows, { rightAlignFrom: 1, totalsRow: true })
  );

  // Parcelas
  body += htmlSection('Resumo de Parcelas', htmlKpiGrid([
    { label: 'Parcela Media', value: fmtBRL(parcelaMedia) },
    { label: 'Parcela Maxima', value: fmtBRL(parcelaMax), variant: 'red' },
    { label: 'Parcela Minima', value: fmtBRL(parcelaMin) },
  ]));

  return buildHTML('Resumo Executivo - ' + data.clienteNome, body);
}

// =============================================================================
// 2. FLUXO FINANCEIRO
// =============================================================================

export function gerarHTMLFluxoFinanceiro(data: ReportData): string {
  const p = data.params;
  const t = getTotais(data.fluxo);
  const fluxoMensal = getFluxoMensal(data.fluxo);
  const valorCredito = safeGet(p, 'valor_credito', safeGet(p, 'valor_carta', 0));

  let body = '';
  body += htmlHeader('Fluxo Financeiro', `${data.clienteNome} | ${data.administradora}`, 'SOMUS CAPITAL');

  body += htmlKpiGrid([
    { label: 'Valor Credito', value: fmtBRL(valorCredito) },
    { label: 'Prazo', value: fmtMes(safeGet(p, 'prazo_meses', 0)) + ' meses' },
    { label: 'Contemplacao', value: 'Mes ' + fmtMes(safeGet(p, 'momento_contemplacao', safeGet(p, 'prazo_contemp', 0))) },
    { label: 'Total Pago', value: fmtBRL(safeGet(t, 'total_pago', 0)), variant: 'red' },
  ]);

  const headers = [
    'Mes', 'Amortizacao', 'Saldo Princ.', 'Tx Adm', 'Fdo Reserva',
    'Parcela Base', 'Reajuste', 'Parc. Reaj.', 'Seguro', 'Outros', 'Total', 'Fluxo Cx',
  ];

  const rows: string[][] = [];
  for (const f of fluxoMensal) {
    const fluxoCaixa = f.fluxo_caixa ?? 0;
    const fluxoClass = fluxoCaixa >= 0 ? 'text-green' : '';
    rows.push([
      fmtMes(f.mes),
      fmtBRL(f.amortizacao ?? 0),
      `<span class="text-purple">${fmtBRL(f.saldo_principal ?? f.saldo_devedor ?? 0)}</span>`,
      fmtBRL(f.valor_parcela_ta ?? f.taxa_adm_antecipada ?? 0),
      fmtBRL(f.valor_fundo_reserva ?? 0),
      fmtBRL(f.valor_parcela ?? 0),
      fmtPct((f.pct_reajuste ?? 0) * 100),
      fmtBRL(f.parcela_apos_reajuste ?? 0),
      fmtBRL(f.seguro_vida ?? 0),
      fmtBRL(f.outros_custos ?? 0),
      `<strong>${fmtBRL(f.parcela_com_seguro ?? f.parcela_apos_reajuste ?? 0)}</strong>`,
      `<span class="${fluxoClass} text-gold font-bold">${fmtBRL(fluxoCaixa)}</span>`,
    ]);
  }

  // Totals
  rows.push([
    '<strong>TOTAL</strong>', '', '', '', '', '', '', '',
    `<strong>${fmtBRL(safeGet(t, 'total_seguro', 0))}</strong>`,
    `<strong>${fmtBRL(safeGet(t, 'total_custos_acessorios', 0))}</strong>`,
    `<strong>${fmtBRL(safeGet(t, 'total_pago', 0))}</strong>`,
    '',
  ]);

  body += htmlSection('Fluxo Mensal Completo',
    `<div style="overflow-x:auto">${htmlTable(headers, rows, { rightAlignFrom: 1, totalsRow: true })}</div>`
  );

  return buildHTML('Fluxo Financeiro - ' + data.clienteNome, body);
}

// =============================================================================
// 3. ANALISE VPL
// =============================================================================

export function gerarHTMLAnaliseVPL(data: ReportData): string {
  const v = data.vpl || {};
  const p = data.params;
  const valorCredito = safeGet(p, 'valor_credito', safeGet(p, 'valor_carta', 0));
  const criaValor = safeGet(v, 'cria_valor', false);
  const deltaVpl = safeGet(v, 'delta_vpl', 0);

  let body = '';
  body += htmlHeader('Analise de VPL', `${data.clienteNome} | NASA Engine HD`, 'SOMUS CAPITAL');
  body += htmlBadge(criaValor ? 'CRIA VALOR' : 'DESTROI VALOR', criaValor);

  // Decomposicao
  const decompRows: string[][] = [
    ['B0 (VP Beneficio)', fmtBRL(safeGet(v, 'b0', 0)), fmtPct(valorCredito > 0 ? (safeGet(v, 'b0', 0) / valorCredito) * 100 : 0)],
    ['H0 (VP Custos pre-T)', fmtBRL(safeGet(v, 'h0', 0)), fmtPct(valorCredito > 0 ? (safeGet(v, 'h0', 0) / valorCredito) * 100 : 0)],
    ['D0 (VP Custos pos-T)', fmtBRL(safeGet(v, 'd0', 0)), fmtPct(valorCredito > 0 ? (safeGet(v, 'd0', 0) / valorCredito) * 100 : 0)],
    ['PV pos-T (na contemp)', fmtBRL(safeGet(v, 'pv_pos_t', 0)), fmtPct(valorCredito > 0 ? (safeGet(v, 'pv_pos_t', 0) / valorCredito) * 100 : 0)],
    [
      '<strong>Delta VPL</strong>',
      `<strong class="${deltaVpl >= 0 ? 'text-green' : 'text-red'}">${fmtBRL(deltaVpl)}</strong>`,
      `<strong>${fmtPct(valorCredito > 0 ? (deltaVpl / valorCredito) * 100 : 0)}</strong>`,
    ],
  ];

  body += htmlSection('Decomposicao do VPL',
    htmlTable(['Componente', 'Valor', '% do Credito'], decompRows, { rightAlignFrom: 1, totalsRow: true })
  );

  // Parametros
  body += htmlSection('Parametros da Analise', htmlKpiGrid([
    { label: 'ALM Anual', value: fmtPct(safeGet(p, 'alm_anual', 0)) },
    { label: 'Hurdle Anual', value: fmtPct(safeGet(p, 'hurdle_anual', 0)) },
    { label: 'Break-even Lance', value: fmtPct(safeGet(v, 'break_even_lance', 0) * 100), variant: 'gold' },
    { label: 'TMA', value: fmtPct(safeGet(p, 'tma', 0) * 100) },
  ]));

  // VP detail table
  const pvPreDetail: any[] = safeGet(v, 'pv_pre_t_detail', []);
  const pvPosDetail: any[] = safeGet(v, 'pv_pos_t_detail', []);

  if (pvPreDetail.length > 0 || pvPosDetail.length > 0) {
    const maxLen = Math.max(pvPreDetail.length, pvPosDetail.length);
    const vpRows: string[][] = [];
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
    body += htmlSection('VP por Periodo',
      `<div style="overflow-x:auto">${htmlTable(
        ['Mes', 'Pgto Pre-T', 'VP Pre-T', 'Pgto Pos-T', 'VP Pos-T'],
        vpRows,
        { rightAlignFrom: 1 }
      )}</div>`
    );
  }

  return buildHTML('Analise VPL - ' + data.clienteNome, body);
}

// =============================================================================
// 4. COMPARATIVO
// =============================================================================

export function gerarHTMLComparativo(data: ReportData): string {
  const comp = data.comparativo || {};

  const econNominal = safeGet(comp, 'economia_nominal', 0);
  const econVpl = safeGet(comp, 'economia_vpl', 0);

  let body = '';
  body += htmlHeader('Comparativo', `Consorcio vs Financiamento | ${data.clienteNome}`, 'SOMUS CAPITAL');

  // KPIs
  body += htmlKpiGrid([
    { label: 'Total Pago Consorcio', value: fmtBRL(safeGet(comp, 'total_pago_consorcio', 0)) },
    { label: 'Total Pago Financiamento', value: fmtBRL(safeGet(comp, 'total_pago_financiamento', 0)), variant: 'purple' },
    { label: 'Economia Nominal', value: fmtBRL(econNominal), variant: econNominal > 0 ? '' : 'red' },
  ]);

  body += htmlKpiGrid([
    { label: 'VPL Consorcio', value: fmtBRL(safeGet(comp, 'vpl_consorcio', 0)) },
    { label: 'VPL Financiamento', value: fmtBRL(safeGet(comp, 'vpl_financiamento', 0)), variant: 'purple' },
    { label: 'Economia VPL', value: fmtBRL(econVpl), variant: econVpl > 0 ? '' : 'red' },
  ]);

  // Comparison table
  const compRows: string[][] = [
    [
      'Total Pago',
      fmtBRL(safeGet(comp, 'total_pago_consorcio', 0)),
      fmtBRL(safeGet(comp, 'total_pago_financiamento', 0)),
      `<span class="${econNominal > 0 ? 'text-green' : 'text-red'}">${fmtBRL(econNominal)}</span>`,
    ],
    [
      'VPL (ALM)',
      fmtBRL(safeGet(comp, 'vpl_consorcio', 0)),
      fmtBRL(safeGet(comp, 'vpl_financiamento', 0)),
      `<span class="${econVpl > 0 ? 'text-green' : 'text-red'}">${fmtBRL(econVpl)}</span>`,
    ],
    [
      'VPL (TMA)',
      fmtBRL(safeGet(comp, 'vpl_consorcio_tma', 0)),
      fmtBRL(safeGet(comp, 'vpl_financiamento_tma', 0)),
      `<span class="${safeGet(comp, 'economia_vpl_tma', 0) > 0 ? 'text-green' : 'text-red'}">${fmtBRL(safeGet(comp, 'economia_vpl_tma', 0))}</span>`,
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
  ];

  body += htmlSection('Tabela Comparativa',
    htmlTable(['Indicador', 'Consorcio', 'Financiamento', 'Diferenca'], compRows, { rightAlignFrom: 1 })
  );

  // PV flows
  const pvCons: number[] = safeGet(comp, 'pv_consorcio', []);
  const pvFin: number[] = safeGet(comp, 'pv_financiamento', []);

  if (pvCons.length > 0 || pvFin.length > 0) {
    const maxLen = Math.max(pvCons.length, pvFin.length);
    const pvRows: string[][] = [];
    for (let i = 0; i < Math.min(maxLen, 60); i++) {
      const diff = (pvCons[i] ?? 0) - (pvFin[i] ?? 0);
      pvRows.push([
        fmtMes(i),
        fmtBRL(pvCons[i] ?? 0),
        fmtBRL(pvFin[i] ?? 0),
        `<span class="${diff > 0 ? 'text-green' : 'text-red'}">${fmtBRL(diff)}</span>`,
      ]);
    }
    body += htmlSection('Valor Presente dos Fluxos',
      `<div style="overflow-x:auto">${htmlTable(
        ['Mes', 'VP Consorcio', 'VP Financiamento', 'Diferenca'],
        pvRows,
        { rightAlignFrom: 1 }
      )}</div>`
    );
  }

  return buildHTML('Comparativo - ' + data.clienteNome, body);
}

// =============================================================================
// 5. VENDA DA OPERACAO
// =============================================================================

export function gerarHTMLVendaOperacao(data: ReportData): string {
  const vd = data.venda || {};

  const ganhoNominal = safeGet(vd, 'ganho_nominal', 0);
  const lucro = ganhoNominal >= 0;

  let body = '';
  body += htmlHeader('Venda da Operacao', data.clienteNome, 'SOMUS CAPITAL');
  body += htmlBadge(
    lucro ? `LUCRO: ${fmtBRL(ganhoNominal)}` : `PREJUIZO: ${fmtBRL(ganhoNominal)}`,
    lucro
  );

  // KPIs
  body += htmlKpiGrid([
    { label: 'VPL Vendedor', value: fmtBRL(safeGet(vd, 'vpl_vendedor', 0)) },
    { label: 'Ganho %', value: fmtPct(safeGet(vd, 'ganho_pct', 0)), variant: lucro ? '' : 'red' },
    { label: 'Prazo Medio', value: fmtNum(safeGet(vd, 'prazo_medio', 0), 1) + ' meses' },
  ]);
  body += htmlKpiGrid([
    { label: 'Ganho Mensal', value: fmtBRL(safeGet(vd, 'ganho_mensal', 0)) },
    { label: 'Margem Mensal', value: fmtPct(safeGet(vd, 'margem_mensal_pct', 0)) },
    { label: 'Total Investido', value: fmtBRL(safeGet(vd, 'total_investido', 0)), variant: 'red' },
  ]);
  body += htmlKpiGrid([
    { label: 'Valor Venda', value: fmtBRL(safeGet(vd, 'valor_venda', 0)), variant: 'gold' },
    { label: 'TIR Vendedor a.a.', value: fmtPct(safeGet(vd, 'tir_vendedor_anual', 0) * 100) },
    { label: 'TIR Comprador a.a.', value: fmtPct(safeGet(vd, 'tir_comprador_anual', 0) * 100), variant: 'purple' },
  ]);

  // Seller flow
  const cfVendedor: number[] = safeGet(vd, 'cashflow_vendedor', []);
  if (cfVendedor.length > 0) {
    const vendRows = cfVendedor.map((val, i) => [
      fmtMes(i),
      `<span class="${val >= 0 ? 'text-green' : ''}">${fmtBRL(val)}</span>`,
    ]);
    body += htmlSection('Fluxo do Vendedor',
      htmlTable(['Mes', 'Fluxo de Caixa'], vendRows, { rightAlignFrom: 1 })
    );
  }

  // Buyer flow
  const cfComprador: number[] = safeGet(vd, 'cashflow_comprador', []);
  if (cfComprador.length > 0) {
    const compRows = cfComprador.map((val, i) => [
      fmtMes(i),
      `<span class="${val >= 0 ? 'text-green' : ''}">${fmtBRL(val)}</span>`,
    ]);
    body += htmlSection('Fluxo do Comprador',
      htmlTable(['Mes', 'Fluxo de Caixa'], compRows, { rightAlignFrom: 1 })
    );
  }

  body += htmlKpiGrid([
    { label: 'VPL Comprador', value: fmtBRL(safeGet(vd, 'vpl_comprador', 0)), variant: 'purple' },
  ]);

  return buildHTML('Venda da Operacao - ' + data.clienteNome, body);
}

// =============================================================================
// 6. CREDITO PARA LANCE
// =============================================================================

export function gerarHTMLCreditoLance(data: ReportData): string {
  const cl = data.creditoLance || {};
  const cc = data.custoCombinado;

  let body = '';
  body += htmlHeader('Credito para Lance', data.clienteNome, 'SOMUS CAPITAL');

  // Financing params
  body += htmlSection('Parametros do Financiamento', htmlKpiGrid([
    { label: 'Valor Financiado', value: fmtBRL(safeGet(cl, 'valor', 0)) },
    { label: 'Total Pago', value: fmtBRL(safeGet(cl, 'total_pago', 0)), variant: 'red' },
    { label: 'Total Juros', value: fmtBRL(safeGet(cl, 'total_juros', 0)), variant: 'red' },
  ]));

  body += htmlKpiGrid([
    { label: 'IOF', value: fmtBRL(safeGet(cl, 'iof', 0)) },
    { label: 'CET Anual', value: fmtPct(safeGet(cl, 'cet_anual', 0) * 100), variant: 'gold' },
    { label: 'TIR Anual', value: fmtPct(safeGet(cl, 'tir_anual', 0) * 100), variant: 'purple' },
  ]);

  if (safeGet(cl, 'custos_iniciais', 0) > 0) {
    body += htmlKpiGrid([
      { label: 'Custos Iniciais (TAC + Garantia + Comissao)', value: fmtBRL(safeGet(cl, 'custos_iniciais', 0)), variant: 'gold' },
    ]);
  }

  if (safeGet(cl, 'valor_antecipacao', 0) > 0) {
    body += htmlKpiGrid([
      { label: 'Antecipacao no Mes', value: fmtMes(safeGet(cl, 'mes_antecipacao', 0)) },
      { label: 'Valor Antecipacao', value: fmtBRL(safeGet(cl, 'valor_antecipacao', 0)), variant: 'gold' },
    ]);
  }

  // Amortization table
  const parcelas: any[] = safeGet(cl, 'parcelas', []);
  if (parcelas.length > 0) {
    const amortRows = parcelas.map((row: any) => [
      fmtMes(row.mes),
      fmtBRL(row.parcela ?? 0),
      `<span class="text-red">${fmtBRL(row.juros ?? 0)}</span>`,
      fmtBRL(row.amortizacao ?? 0),
      `<span class="text-purple">${fmtBRL(row.saldo ?? 0)}</span>`,
    ]);
    body += htmlSection('Tabela de Amortizacao',
      `<div style="overflow-x:auto">${htmlTable(
        ['Mes', 'Parcela', 'Juros', 'Amortizacao', 'Saldo'],
        amortRows,
        { rightAlignFrom: 1 }
      )}</div>`
    );
  }

  // Combined cost
  if (cc) {
    body += htmlSection('Custo Combinado (Consorcio + Financiamento Lance)', '');
    body += htmlKpiGrid([
      { label: 'Total Consorcio', value: fmtBRL(safeGet(cc, 'total_pago_consorcio', 0)) },
      { label: 'Total Lance (Fin.)', value: fmtBRL(safeGet(cc, 'total_pago_lance', 0)), variant: 'purple' },
      { label: 'Total Combinado', value: fmtBRL(safeGet(cc, 'total_pago_combinado', 0)), variant: 'gold' },
    ]);
    body += htmlKpiGrid([
      { label: 'TIR Combinada Mensal', value: fmtPct(safeGet(cc, 'tir_mensal_combinado', 0) * 100) },
      { label: 'TIR Combinada Anual', value: fmtPct(safeGet(cc, 'tir_anual_combinado', 0) * 100), variant: 'purple' },
      { label: 'CET Combinado Anual', value: fmtPct(safeGet(cc, 'cet_anual_combinado', 0) * 100), variant: 'gold' },
    ]);
  }

  return buildHTML('Credito para Lance - ' + data.clienteNome, body);
}

/**
 * report-manager.ts - Unified report manager
 * Somus Capital - Mesa de Produtos
 *
 * Routes report generation to the appropriate generator (PDF, HTML, or both).
 */

import type { ReportData } from './report-types';

import {
  gerarPDFResumoExecutivo,
  gerarPDFFluxoFinanceiro,
  gerarPDFAnaliseVPL,
  gerarPDFComparativo,
  gerarPDFVendaOperacao,
  gerarPDFCreditoLance,
  gerarTodosPDFs,
} from './report-pdf';

import {
  gerarHTMLResumoExecutivo,
  gerarHTMLFluxoFinanceiro,
  gerarHTMLAnaliseVPL,
  gerarHTMLComparativo,
  gerarHTMLVendaOperacao,
  gerarHTMLCreditoLance,
  downloadHTML,
} from './report-html';

// =============================================================================
// TYPES
// =============================================================================

export type ReportFormat = 'pdf' | 'html' | 'both';
export type ReportType = 'resumo' | 'fluxo' | 'vpl' | 'comparativo' | 'venda' | 'credito';

// =============================================================================
// PDF GENERATORS MAP
// =============================================================================

const pdfGenerators: Record<ReportType, (data: ReportData) => void> = {
  resumo: gerarPDFResumoExecutivo,
  fluxo: gerarPDFFluxoFinanceiro,
  vpl: gerarPDFAnaliseVPL,
  comparativo: gerarPDFComparativo,
  venda: gerarPDFVendaOperacao,
  credito: gerarPDFCreditoLance,
};

// =============================================================================
// HTML GENERATORS MAP
// =============================================================================

const htmlGenerators: Record<ReportType, (data: ReportData) => string> = {
  resumo: gerarHTMLResumoExecutivo,
  fluxo: gerarHTMLFluxoFinanceiro,
  vpl: gerarHTMLAnaliseVPL,
  comparativo: gerarHTMLComparativo,
  venda: gerarHTMLVendaOperacao,
  credito: gerarHTMLCreditoLance,
};

// =============================================================================
// FILENAME MAP
// =============================================================================

const filenameMap: Record<ReportType, string> = {
  resumo: 'Resumo_Executivo',
  fluxo: 'Fluxo_Financeiro',
  vpl: 'Analise_VPL',
  comparativo: 'Comparativo',
  venda: 'Venda_Operacao',
  credito: 'Credito_Lance',
};

// =============================================================================
// REPORT AVAILABILITY
// =============================================================================

/** Check if a report type has the required data */
function isReportAvailable(type: ReportType, data: ReportData): boolean {
  switch (type) {
    case 'resumo':
    case 'fluxo':
    case 'vpl':
      return !!data.fluxo;
    case 'comparativo':
      return !!data.comparativo;
    case 'venda':
      return !!data.venda;
    case 'credito':
      return !!data.creditoLance;
    default:
      return false;
  }
}

// =============================================================================
// SINGLE REPORT
// =============================================================================

/**
 * Gera um relatorio no formato especificado.
 * @param type Tipo do relatorio
 * @param data Dados do relatorio (ReportData)
 * @param format 'pdf', 'html' ou 'both' (default: 'both')
 */
export function gerarRelatorio(
  type: ReportType,
  data: ReportData,
  format: ReportFormat = 'both'
): void {
  if (!isReportAvailable(type, data)) {
    console.warn(`[ReportManager] Dados insuficientes para relatorio: ${type}`);
    return;
  }

  const clienteSafe = data.clienteNome.replace(/\s+/g, '_');
  const filename = `${filenameMap[type]}_${clienteSafe}`;

  if (format === 'pdf' || format === 'both') {
    pdfGenerators[type](data);
  }

  if (format === 'html' || format === 'both') {
    const html = htmlGenerators[type](data);
    downloadHTML(html, `${filename}.html`);
  }
}

// =============================================================================
// ALL REPORTS
// =============================================================================

/**
 * Gera todos os relatorios disponiveis.
 * Verifica quais modulos tem dados antes de gerar.
 * @param data Dados do relatorio (ReportData)
 * @param format 'pdf', 'html' ou 'both' (default: 'both')
 */
export function gerarTodosRelatorios(
  data: ReportData,
  format: ReportFormat = 'both'
): void {
  const allTypes: ReportType[] = ['resumo', 'fluxo', 'vpl', 'comparativo', 'venda', 'credito'];

  const availableTypes = allTypes.filter((type) => isReportAvailable(type, data));

  if (availableTypes.length === 0) {
    console.warn('[ReportManager] Nenhum relatorio disponivel com os dados fornecidos.');
    return;
  }

  if (format === 'pdf' || format === 'both') {
    // Use staggered download for PDFs to avoid browser blocking
    availableTypes.forEach((type, index) => {
      setTimeout(() => pdfGenerators[type](data), index * 500);
    });
  }

  if (format === 'html' || format === 'both') {
    const clienteSafe = data.clienteNome.replace(/\s+/g, '_');
    availableTypes.forEach((type, index) => {
      const delay = format === 'both' ? (availableTypes.length + index) * 500 : index * 300;
      setTimeout(() => {
        const html = htmlGenerators[type](data);
        downloadHTML(html, `${filenameMap[type]}_${clienteSafe}.html`);
      }, delay);
    });
  }
}

/**
 * Retorna lista de relatorios disponiveis para os dados fornecidos.
 */
export function getRelatoriosDisponiveis(data: ReportData): ReportType[] {
  const allTypes: ReportType[] = ['resumo', 'fluxo', 'vpl', 'comparativo', 'venda', 'credito'];
  return allTypes.filter((type) => isReportAvailable(type, data));
}

/**
 * Retorna o HTML de um relatorio sem efetuar download (para preview).
 */
export function getHTMLPreview(type: ReportType, data: ReportData): string | null {
  if (!isReportAvailable(type, data)) return null;
  return htmlGenerators[type](data);
}

// Re-export types and utilities
export type { ReportData } from './report-types';
export { downloadHTML } from './report-html';

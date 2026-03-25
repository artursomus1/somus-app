/**
 * report-types.ts - Shared types for PDF and HTML report generators
 * Somus Capital - Mesa de Produtos
 */

export interface ReportData {
  // Metadata
  clienteNome: string;
  assessor: string;
  administradora: string;
  dataGeracao: string;

  // Core results (from engine)
  params: Record<string, any>;
  fluxo: Record<string, any>;       // from calcularFluxoCompleto / calcularFluxoConsorcio
  vpl: Record<string, any>;         // from calcularVplHd

  // Optional modules
  financiamento?: Record<string, any>;   // from calcularFinanciamentoStandalone
  comparativo?: Record<string, any>;     // from compararConsorcioFinanciamentoStandalone
  venda?: Record<string, any>;           // from calcularVendaOperacao
  creditoLance?: Record<string, any>;    // from calcularCreditoLance
  custoCombinado?: Record<string, any>;  // from calcularCustoCombinado
}

/** Formatting helpers shared across generators */
export function fmtBRL(v: number | undefined | null): string {
  if (v == null || isNaN(v)) return 'R$ 0,00';
  return v.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
}

export function fmtPct(v: number | undefined | null, decimals = 2): string {
  if (v == null || isNaN(v)) return '0,00%';
  return v.toFixed(decimals).replace('.', ',') + '%';
}

export function fmtNum(v: number | undefined | null, decimals = 2): string {
  if (v == null || isNaN(v)) return '0,00';
  return v.toLocaleString('pt-BR', {
    minimumFractionDigits: decimals,
    maximumFractionDigits: decimals,
  });
}

export function fmtMes(v: number | undefined | null): string {
  if (v == null || isNaN(v)) return '0';
  return v.toString();
}

/** Safe accessor for nested data */
export function safeGet(obj: Record<string, any> | undefined | null, key: string, fallback: any = 0): any {
  if (!obj) return fallback;
  return obj[key] ?? fallback;
}

/** Get totais from fluxo (handles both new and legacy shapes) */
export function getTotais(fluxo: Record<string, any>): Record<string, any> {
  return fluxo?.totais ?? fluxo ?? {};
}

/** Get metricas from fluxo */
export function getMetricas(fluxo: Record<string, any>): Record<string, any> {
  return fluxo?.metricas ?? fluxo ?? {};
}

/** Get fluxo mensal array */
export function getFluxoMensal(fluxo: Record<string, any>): any[] {
  return fluxo?.fluxo ?? fluxo?.fluxo_mensal ?? [];
}

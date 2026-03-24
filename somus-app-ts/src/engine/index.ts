/**
 * Engine module index - re-exports all engine components
 * Somus Capital - Mesa de Produtos
 */

// Core financial math
export {
  npv,
  irr,
  pmt,
  annualFromMonthly,
  monthlyFromAnnual,
  goalSeek,
} from './irr-npv';

// Main NASA engine
export {
  NasaEngine,
  FREQ_MAP,
  DEFAULT_CONFIG,
  DEFAULT_PARAMS,
  // Standalone compatibility functions
  calcularFluxoConsorcio,
  calcularVplHd,
  calcularFinanciamentoStandalone,
  compararConsorcioFinanciamentoStandalone,
} from './nasa-engine';

// Types from NASA engine
export type {
  NasaConfig,
  NasaParams,
  PeriodoDistribuicao,
  CustoAcessorio,
  FluxoMensal,
  FluxoTotais,
  FluxoMetricas,
  FluxoResult,
  VPLResult,
  ParcelaInfo,
  ResumoCliente,
  FinanciamentoParcelaRow,
  FinanciamentoResult,
} from './nasa-engine';

// Financing module
export {
  calcularFinanciamento,
  calcularCreditoLance,
  calcularIOF,
} from './financiamento';
export type {
  FinanciamentoParams,
  CreditoLanceParams,
} from './financiamento';

// Comparison module
export {
  compararConsorcioFinanciamento,
  calcularCustoCombinado,
  calcularVendaOperacao,
  calcularCreditoEquivalente,
} from './comparativo';
export type {
  ComparativoResult,
  VendaOperacaoParams,
  VendaResult,
  CustoCombinado,
} from './comparativo';

// Scenario manager
export { ScenarioManager } from './scenario-manager';
export type {
  Scenario,
  ScenarioComparison,
  ScenarioSummary,
} from './scenario-manager';

// Consolidation
export {
  consolidarCotas,
  consolidarCotasDetalhado,
  consolidarCotasComQuantidade,
} from './consolidador-cotas';
export type {
  GrupoInput,
  ConsolidacaoResult,
  GrupoDetalhadoInput,
  ConsolidacaoDetalhadaResult,
} from './consolidador-cotas';

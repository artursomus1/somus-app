/**
 * comparativo.ts - Comparison functions
 * Port of nasa_engine_hd.py comparison and sale analysis functions
 * Somus Capital - Mesa de Produtos
 *
 * Side-by-side consortium vs financing comparison,
 * combined cost analysis, operation sale, and equivalent credit.
 */

import { npv, irr, annualFromMonthly, monthlyFromAnnual, goalSeek } from './irr-npv';
import { NasaEngine } from './nasa-engine';
import type {
  FluxoResult,
  VPLResult,
  FinanciamentoResult,
} from './nasa-engine';
import { calcularFinanciamento } from './financiamento';
import type { FinanciamentoParams } from './financiamento';

// =============================================================================
// TYPES
// =============================================================================

export interface ComparativoResult {
  consorcio: FluxoResult;
  financiamento: FinanciamentoResult;
  // Nominal values
  total_pago_consorcio: number;
  total_pago_financiamento: number;
  economia_nominal: number;
  // VPL at ALM rate
  vpl_consorcio: number;
  vpl_financiamento: number;
  economia_vpl: number;
  // VPL at TMA rate
  vpl_consorcio_tma: number;
  vpl_financiamento_tma: number;
  economia_vpl_tma: number;
  // IRR
  tir_consorcio_mensal: number;
  tir_consorcio_anual: number;
  tir_financ_mensal: number;
  tir_financ_anual: number;
  // Ratios
  razao_vpl_consorcio: number;
  razao_vpl_financ: number;
  // PV flows (for charting)
  pv_consorcio: number[];
  pv_financiamento: number[];
}

export interface VendaOperacaoParams {
  momento_venda: number;
  valor_venda: number;
  tma: number;
}

export interface VendaResult {
  cashflow_vendedor: number[];
  cashflow_comprador: number[];
  vpl_vendedor: number;
  vpl_comprador: number;
  tir_vendedor_mensal: number;
  tir_vendedor_anual: number;
  tir_comprador_mensal: number;
  tir_comprador_anual: number;
  ganho_nominal: number;
  ganho_pct: number;
  total_investido: number;
  valor_venda: number;
  prazo_medio: number;
  ganho_mensal: number;
  margem_mensal_pct: number;
}

export interface CustoCombinado {
  cashflowCombinado: number[];
  totalPago: number;
  tirMensal: number;
  tirAnual: number;
}

// =============================================================================
// CONSORTIUM VS FINANCING COMPARISON
// =============================================================================

/**
 * Side-by-side comparison of consortium vs financing.
 * Includes nominal and present value flows.
 *
 * @param paramsConsorcio - consortium parameters (NasaParams format)
 * @param paramsFinanc - financing parameters
 * @param tma - monthly minimum attractiveness rate (decimal)
 * @returns complete comparison result
 */
export function compararConsorcioFinanciamento(
  paramsConsorcio: Record<string, any>,
  paramsFinanc: FinanciamentoParams,
  tma: number
): ComparativoResult {
  let tmaM = tma;
  if (tmaM > 1) tmaM = tmaM / 100; // Correct if passed as percentage

  const almA = (paramsConsorcio.alm_anual ?? 12.0) / 100;
  const almM = monthlyFromAnnual(almA);

  // Consortium calculation
  const engine = new NasaEngine();
  const fluxoC = engine.calcularFluxoCompleto(paramsConsorcio);
  const cfC = fluxoC.cashflow;
  const vplCAlm = npv(almM, cfC);
  const vplCTma = npv(tmaM, cfC);
  const tirC = irr(cfC, 0.005);

  // Financing calculation
  const fluxoF = calcularFinanciamento(paramsFinanc);
  const cfF = fluxoF.cashflow;
  const vplFAlm = npv(almM, cfF);
  const vplFTma = npv(tmaM, cfF);
  const tirF = irr(cfF, 0.008);

  // Nominal totals
  const totalCons = fluxoC.total_pago ?? fluxoC.totais?.total_pago ?? 0;
  const totalFin = fluxoF.total_pago;

  const cartaLiq = fluxoC.carta_liquida ?? fluxoC.totais?.carta_liquida ?? 0;
  const valorFin = fluxoF.valor;

  const razaoC = cartaLiq > 0 ? totalCons / cartaLiq : 0;
  const razaoF = valorFin > 0 ? totalFin / valorFin : 0;

  // PV of each installment (for charts)
  const pvCons = cfC.map((cf, i) =>
    tmaM > 0 ? cf / Math.pow(1 + tmaM, i) : cf
  );
  const pvFin = cfF.map((cf, i) =>
    tmaM > 0 ? cf / Math.pow(1 + tmaM, i) : cf
  );

  return {
    consorcio: fluxoC,
    financiamento: fluxoF,
    total_pago_consorcio: totalCons,
    total_pago_financiamento: totalFin,
    economia_nominal: totalFin - totalCons,
    vpl_consorcio: vplCAlm,
    vpl_financiamento: vplFAlm,
    economia_vpl: vplCAlm - vplFAlm,
    vpl_consorcio_tma: vplCTma,
    vpl_financiamento_tma: vplFTma,
    economia_vpl_tma: vplCTma - vplFTma,
    tir_consorcio_mensal: tirC,
    tir_consorcio_anual: annualFromMonthly(tirC),
    tir_financ_mensal: tirF,
    tir_financ_anual: annualFromMonthly(tirF),
    razao_vpl_consorcio: razaoC,
    razao_vpl_financ: razaoF,
    pv_consorcio: pvCons,
    pv_financiamento: pvFin,
  };
}

// =============================================================================
// COMBINED COST (Consortium + Lance Financing)
// =============================================================================

/**
 * Combines consortium flow with lance financing flow.
 * Produces a single merged cashflow with aggregate metrics.
 *
 * @param fluxoConsorcio - result from calcularFluxoCompleto
 * @param fluxoLance - result from lance financing
 * @returns combined cashflow and metrics
 */
export function calcularCustoCombinado(
  fluxoConsorcio: FluxoResult,
  fluxoLance: FinanciamentoResult
): CustoCombinado {
  const cfCons = fluxoConsorcio.cashflow ?? [];
  const cfLance = fluxoLance.cashflow ?? fluxoLance.cashflow_antecipado ?? [];

  const maxLen = Math.max(cfCons.length, cfLance.length);

  const cfCombinado: number[] = [];
  for (let i = 0; i < maxLen; i++) {
    const vCons = i < cfCons.length ? cfCons[i] : 0;
    const vLance = i < cfLance.length ? cfLance[i] : 0;
    cfCombinado.push(vCons + vLance);
  }

  const tirM = irr(cfCombinado, 0.005);
  const tirA = annualFromMonthly(tirM);

  const totalCons = fluxoConsorcio.total_pago ?? fluxoConsorcio.totais?.total_pago ?? 0;
  const totalLance = fluxoLance.custo_efetivo_total ?? fluxoLance.total_pago ?? 0;

  return {
    cashflowCombinado: cfCombinado,
    totalPago: totalCons + totalLance,
    tirMensal: tirM,
    tirAnual: tirA,
  };
}

// =============================================================================
// OPERATION SALE ANALYSIS
// =============================================================================

/**
 * Analyzes the sale of a consortium operation at a given point in time.
 * Computes seller and buyer perspectives including IRR and VPL.
 *
 * @param fluxo - result from calcularFluxoCompleto
 * @param params - sale parameters (timing, price, TMA)
 * @returns complete sale analysis from both perspectives
 */
export function calcularVendaOperacao(
  fluxo: FluxoResult,
  params: VendaOperacaoParams
): VendaResult {
  const momento = params.momento_venda ?? 0;
  const valorVenda = params.valor_venda ?? 0;
  const tma = params.tma ?? 0.01;

  const cfOriginal = fluxo.cashflow ?? [];

  // Seller flow: cashflows up to sale moment + sale proceeds
  const cfVendedor: number[] = [];
  let totalGasto = 0.0;

  for (let i = 0; i < Math.min(momento + 1, cfOriginal.length); i++) {
    cfVendedor.push(cfOriginal[i]);
    if (cfOriginal[i] < 0) {
      totalGasto += Math.abs(cfOriginal[i]);
    }
  }

  // Add sale value on last month
  if (cfVendedor.length > 0) {
    cfVendedor[cfVendedor.length - 1] += valorVenda;
  } else {
    cfVendedor.push(valorVenda);
  }

  const vplVendedor = npv(tma, cfVendedor);
  const tirVendedor = irr(cfVendedor, 0.01);

  const ganhoNominal = valorVenda - totalGasto;
  const ganhoPct = totalGasto > 0 ? (ganhoNominal / totalGasto) * 100 : 0;

  const prazoMedio = momento / 2; // simplification
  const ganhoMensal = momento > 0 ? ganhoNominal / momento : 0;
  const margemMensal = totalGasto > 0 ? (ganhoMensal / totalGasto) * 100 : 0;

  // Buyer flow: pays sale price, continues with remaining installments
  const cfComprador: number[] = [-valorVenda];
  for (let i = momento + 1; i < cfOriginal.length; i++) {
    cfComprador.push(cfOriginal[i]);
  }

  const tirComprador = cfComprador.length > 1 ? irr(cfComprador, 0.005) : 0;
  const vplComprador = npv(tma, cfComprador);

  return {
    cashflow_vendedor: cfVendedor,
    cashflow_comprador: cfComprador,
    vpl_vendedor: vplVendedor,
    vpl_comprador: vplComprador,
    tir_vendedor_mensal: tirVendedor,
    tir_vendedor_anual: annualFromMonthly(tirVendedor),
    tir_comprador_mensal: tirComprador,
    tir_comprador_anual: annualFromMonthly(tirComprador),
    ganho_nominal: ganhoNominal,
    ganho_pct: ganhoPct,
    total_investido: totalGasto,
    valor_venda: valorVenda,
    prazo_medio: prazoMedio,
    ganho_mensal: ganhoMensal,
    margem_mensal_pct: margemMensal,
  };
}

// =============================================================================
// EQUIVALENT CREDIT (GoalSeek)
// =============================================================================

/**
 * Finds the credit value that covers all operation costs.
 * Uses GoalSeek to solve iteratively: find X where total_paid(X) = X.
 *
 * @param params - consortium parameters (NasaParams format)
 * @returns equivalent credit value
 */
export function calcularCreditoEquivalente(
  params: Record<string, any>
): number {
  const creditoOriginal = params.valor_credito ?? 0;
  const engine = new NasaEngine();

  const custoTotalFunc = (creditoTeste: number): number => {
    const p = { ...params, valor_credito: creditoTeste };
    const fluxo = engine.calcularFluxoCompleto(p);
    return fluxo.total_pago ?? fluxo.totais?.total_pago ?? 0;
  };

  try {
    return goalSeek(
      custoTotalFunc,
      creditoOriginal,
      creditoOriginal * 0.1,
      creditoOriginal * 2.0,
      1.0,
      500
    );
  } catch {
    return creditoOriginal;
  }
}

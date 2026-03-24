/**
 * financiamento.ts - Financing calculations
 * Port of nasa_engine_hd.py financing functions
 * Somus Capital - Mesa de Produtos
 *
 * Supports SAC and Price amortization methods, grace period, IOF, TAC,
 * guarantee evaluation costs, and early payoff.
 */

import { irr, pmt, annualFromMonthly } from './irr-npv';

// =============================================================================
// CONSTANTS
// =============================================================================

const IOF_DIARIO = 0.000082; // 0.0082% a.d.
const IOF_ADICIONAL = 0.0038; // 0.38% on principal

// =============================================================================
// TYPES
// =============================================================================

export interface FinanciamentoParams {
  valor: number;
  prazo_meses: number;
  taxa_mensal_pct: number;
  metodo: 'price' | 'sac';
  carencia?: number;
  calcular_iof?: boolean;
  custos_adicionais?: Array<{
    descricao: string;
    valor: number;
    momento: number;
  }>;
}

export interface FinanciamentoParcelaRow {
  mes: number;
  parcela: number;
  juros: number;
  amortizacao: number;
  saldo: number;
}

export interface FinanciamentoResult {
  parcelas: FinanciamentoParcelaRow[];
  cashflow: number[];
  total_pago: number;
  total_juros: number;
  total_amortizado: number;
  valor: number;
  iof: number;
  custos_adicionais: number;
  custo_efetivo_total: number;
  tir_mensal: number;
  tir_anual: number;
  cet_anual: number;
  // Optional fields from lance financing
  cashflow_antecipado?: number[];
  valor_antecipacao?: number;
  mes_antecipacao?: number;
  custos_iniciais?: number;
}

export interface CreditoLanceParams {
  valor: number;
  prazo: number;
  taxa: number;
  metodo: 'price' | 'sac';
  carencia?: number;
  tac?: number;
  avaliacaoGarantia?: number;
  comissao?: number;
  pagamentoAntecipado?: { mes: number; valor: number };
}

// =============================================================================
// IOF CALCULATION
// =============================================================================

/**
 * Calculate IOF on a credit operation.
 * IOF = IOF_ADICIONAL (0.38%) on principal + IOF_DIARIO on each amortization.
 *
 * @param valor - principal amount
 * @param prazo - term in months
 * @returns total IOF
 */
export function calcularIOF(valor: number, prazo: number): number {
  // Simplified: assumes linear amortization for IOF calculation
  const amortMensal = prazo > 0 ? valor / prazo : 0;
  const iofAdicional = valor * IOF_ADICIONAL;

  let iofDiarioTotal = 0.0;
  for (let m = 1; m <= prazo; m++) {
    let dias = m * 30;
    dias = Math.min(dias, 365);
    iofDiarioTotal += amortMensal * IOF_DIARIO * dias;
  }

  return iofAdicional + iofDiarioTotal;
}

/**
 * Internal IOF calculation using actual amortization schedule.
 */
function calcularIOFDetalhado(
  valorPrincipal: number,
  parcelas: FinanciamentoParcelaRow[]
): number {
  const iofAdicional = valorPrincipal * IOF_ADICIONAL;

  let iofDiarioTotal = 0.0;
  for (const p of parcelas) {
    let dias = p.mes * 30; // approximation: 30 days/month
    dias = Math.min(dias, 365); // daily IOF capped at 365 days
    iofDiarioTotal += p.amortizacao * IOF_DIARIO * dias;
  }

  return iofAdicional + iofDiarioTotal;
}

// =============================================================================
// FULL FINANCING SIMULATION
// =============================================================================

/**
 * Complete financing simulation with SAC or Price amortization.
 * Supports: grace period, IOF, additional costs.
 *
 * @param params - financing parameters
 * @returns complete financing result with installments, cashflow, and metrics
 */
export function calcularFinanciamento(
  params: FinanciamentoParams
): FinanciamentoResult {
  const valor = params.valor ?? 0;
  const prazo = params.prazo_meses ?? 0;
  const taxa = (params.taxa_mensal_pct ?? 0) / 100;
  const metodo = (params.metodo ?? 'price').toLowerCase() as 'price' | 'sac';
  const carencia = params.carencia ?? 0;
  const calcIof = params.calcular_iof !== undefined ? params.calcular_iof : true;
  const custosAdd = params.custos_adicionais ?? [];

  let saldo = valor;
  const parcelas: FinanciamentoParcelaRow[] = [];
  const cashflow: number[] = [valor]; // month 0: receive principal
  let totalPago = 0.0;
  let totalJuros = 0.0;
  let totalAmort = 0.0;

  const prazoAmort = prazo - carencia;

  // PMT for Price method
  let pmtVal: number;
  if (metodo === 'price' && taxa > 0 && prazoAmort > 0) {
    pmtVal = pmt(taxa, prazoAmort, valor);
  } else if (prazoAmort > 0) {
    pmtVal = valor / prazoAmort;
  } else {
    pmtVal = 0;
  }

  for (let m = 1; m <= prazo; m++) {
    const juros = saldo * taxa;
    let amortM: number;
    let parcelaM: number;

    if (m <= carencia) {
      // Grace period: pay only interest
      amortM = 0.0;
      parcelaM = juros;
    } else if (metodo === 'price') {
      parcelaM = pmtVal;
      amortM = parcelaM - juros;
    } else {
      // SAC
      amortM = prazoAmort > 0 ? valor / prazoAmort : 0;
      parcelaM = amortM + juros;
    }

    saldo -= amortM;
    saldo = Math.max(0, saldo);

    totalPago += parcelaM;
    totalJuros += juros;
    totalAmort += amortM;

    parcelas.push({
      mes: m,
      parcela: parcelaM,
      juros,
      amortizacao: amortM,
      saldo,
    });
    cashflow.push(-parcelaM);
  }

  // IOF
  let iofTotal = 0.0;
  if (calcIof) {
    iofTotal = calcularIOFDetalhado(valor, parcelas);
  }

  // Additional costs
  let totalCustosAdd = 0;
  for (const c of custosAdd) {
    totalCustosAdd += c.valor ?? 0;
  }
  for (const c of custosAdd) {
    const mCusto = c.momento ?? 0;
    if (mCusto >= 0 && mCusto <= prazo) {
      if (mCusto === 0) {
        cashflow[0] -= c.valor ?? 0;
      } else if (mCusto <= cashflow.length - 1) {
        cashflow[mCusto] -= c.valor ?? 0;
      }
    }
  }

  // CET (Custo Efetivo Total)
  const cfCet = [...cashflow];
  if (calcIof) {
    cfCet[0] = valor - iofTotal - totalCustosAdd;
  }
  const tirM = irr(cfCet, 0.008);
  const tirA = annualFromMonthly(tirM);

  return {
    parcelas,
    cashflow,
    total_pago: totalPago,
    total_juros: totalJuros,
    total_amortizado: totalAmort,
    valor,
    iof: iofTotal,
    custos_adicionais: totalCustosAdd,
    custo_efetivo_total: totalPago + iofTotal + totalCustosAdd,
    tir_mensal: tirM,
    tir_anual: tirA,
    cet_anual: tirA,
  };
}

// =============================================================================
// LANCE FINANCING (Credit for Bid)
// =============================================================================

/**
 * Financing simulation to cover a consortium bid (lance).
 * Wraps calcularFinanciamento with additional lance-specific parameters.
 *
 * @param params - lance financing parameters
 * @returns financing result with optional early payoff data
 */
export function calcularCreditoLance(params: CreditoLanceParams): FinanciamentoResult {
  const valor = params.valor ?? 0;
  const prazo = params.prazo ?? 0;
  const taxa = params.taxa ?? 0;
  const metodo = (params.metodo ?? 'price') as 'price' | 'sac';
  const carencia = params.carencia ?? 0;
  const tac = params.tac ?? 0;
  const avalGarantia = params.avaliacaoGarantia ?? 0;
  const comissao = params.comissao ?? 0;
  const pagAntecipado = params.pagamentoAntecipado;

  // Build additional costs at month 0
  const custosAdd: Array<{ descricao: string; valor: number; momento: number }> = [];
  if (tac > 0) {
    custosAdd.push({ descricao: 'TAC', valor: tac, momento: 0 });
  }
  if (avalGarantia > 0) {
    custosAdd.push({ descricao: 'Avaliacao Garantia', valor: avalGarantia, momento: 0 });
  }
  if (comissao > 0) {
    custosAdd.push({ descricao: 'Comissao', valor: comissao, momento: 0 });
  }

  const finParams: FinanciamentoParams = {
    valor,
    prazo_meses: prazo,
    taxa_mensal_pct: taxa,
    metodo,
    carencia,
    calcular_iof: true,
    custos_adicionais: custosAdd,
  };

  const resultado = calcularFinanciamento(finParams);

  // Early payoff handling
  if (pagAntecipado && pagAntecipado.mes > 0 && pagAntecipado.mes < prazo) {
    const mesAntecipacao = pagAntecipado.mes;
    let saldoNoMes = 0;
    for (const p of resultado.parcelas) {
      if (p.mes === mesAntecipacao) {
        saldoNoMes = p.saldo;
        break;
      }
    }

    // Recalculate cashflow with early payoff
    const cfNovo = [...resultado.cashflow];
    for (let i = mesAntecipacao + 1; i < cfNovo.length; i++) {
      cfNovo[i] = 0;
    }
    if (mesAntecipacao < cfNovo.length) {
      cfNovo[mesAntecipacao] -= saldoNoMes;
    }

    resultado.cashflow_antecipado = cfNovo;
    resultado.valor_antecipacao = saldoNoMes;
    resultado.mes_antecipacao = mesAntecipacao;
  }

  resultado.custos_iniciais = tac + avalGarantia + comissao;

  return resultado;
}

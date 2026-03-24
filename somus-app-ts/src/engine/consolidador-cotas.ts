/**
 * consolidador-cotas.ts - Multi-group quota consolidation
 * Port of nasa_engine_hd.py consolidar_cotas function
 * Somus Capital - Mesa de Produtos
 *
 * Consolidates multiple consortium groups into weighted averages.
 */

// =============================================================================
// TYPES
// =============================================================================

export interface GrupoInput {
  grupo: string;
  valor: number;
  prazo: number;
  taxaAdm: number;
  fundoReserva: number;
}

export interface ConsolidacaoResult {
  totalCredito: number;
  prazoMedio: number;
  taxaAdmMedia: number;
  fundoReservaMedia: number;
}

export interface GrupoDetalhadoInput {
  grupo: string;
  valor: number;
  prazo: number;
  taxaAdm: number;
  fundoReserva: number;
  seguro?: number;
  lanceEmbutido?: number;
  lanceLivre?: number;
  momentoContemplacao?: number;
  reajusteAnual?: number;
}

export interface ConsolidacaoDetalhadaResult {
  totalCredito: number;
  prazoMedio: number;
  taxaAdmMedia: number;
  fundoReservaMedia: number;
  seguroMedio: number;
  lanceTotalPct: number;
  contemplacaoMedia: number;
  reajusteMedio: number;
  numGrupos: number;
  grupos: GrupoDetalhadoInput[];
}

// =============================================================================
// SIMPLE CONSOLIDATION
// =============================================================================

/**
 * Consolidates multiple consortium groups into weighted averages.
 * Weights are proportional to each group's credit value.
 *
 * @param grupos - array of group data
 * @returns consolidated weighted averages
 */
export function consolidarCotas(
  grupos: GrupoInput[]
): ConsolidacaoResult {
  if (!grupos || grupos.length === 0) {
    return {
      totalCredito: 0,
      prazoMedio: 0,
      taxaAdmMedia: 0,
      fundoReservaMedia: 0,
    };
  }

  let totalCredito = 0;
  let somaPrazoPonderado = 0;
  let somaTaxaAdmPonderada = 0;
  let somaFundoReservaPonderado = 0;

  for (const g of grupos) {
    const valor = g.valor ?? 0;
    totalCredito += valor;
    somaPrazoPonderado += (g.prazo ?? 0) * valor;
    somaTaxaAdmPonderada += (g.taxaAdm ?? 0) * valor;
    somaFundoReservaPonderado += (g.fundoReserva ?? 0) * valor;
  }

  if (totalCredito === 0) {
    return {
      totalCredito: 0,
      prazoMedio: 0,
      taxaAdmMedia: 0,
      fundoReservaMedia: 0,
    };
  }

  return {
    totalCredito,
    prazoMedio: somaPrazoPonderado / totalCredito,
    taxaAdmMedia: somaTaxaAdmPonderada / totalCredito,
    fundoReservaMedia: somaFundoReservaPonderado / totalCredito,
  };
}

// =============================================================================
// DETAILED CONSOLIDATION
// =============================================================================

/**
 * Consolidates multiple consortium groups with full parameter set.
 * All averages are credit-value-weighted.
 *
 * @param grupos - array of detailed group data
 * @returns detailed consolidated result
 */
export function consolidarCotasDetalhado(
  grupos: GrupoDetalhadoInput[]
): ConsolidacaoDetalhadaResult {
  if (!grupos || grupos.length === 0) {
    return {
      totalCredito: 0,
      prazoMedio: 0,
      taxaAdmMedia: 0,
      fundoReservaMedia: 0,
      seguroMedio: 0,
      lanceTotalPct: 0,
      contemplacaoMedia: 0,
      reajusteMedio: 0,
      numGrupos: 0,
      grupos: [],
    };
  }

  let totalCredito = 0;
  let somaPrazoPonderado = 0;
  let somaTaxaAdmPonderada = 0;
  let somaFundoReservaPonderado = 0;
  let somaSeguroPonderado = 0;
  let somaLancePonderado = 0;
  let somaContempPonderado = 0;
  let somaReajustePonderado = 0;

  for (const g of grupos) {
    const valor = g.valor ?? 0;
    totalCredito += valor;
    somaPrazoPonderado += (g.prazo ?? 0) * valor;
    somaTaxaAdmPonderada += (g.taxaAdm ?? 0) * valor;
    somaFundoReservaPonderado += (g.fundoReserva ?? 0) * valor;
    somaSeguroPonderado += (g.seguro ?? 0) * valor;
    somaLancePonderado +=
      ((g.lanceEmbutido ?? 0) + (g.lanceLivre ?? 0)) * valor;
    somaContempPonderado += (g.momentoContemplacao ?? 0) * valor;
    somaReajustePonderado += (g.reajusteAnual ?? 0) * valor;
  }

  if (totalCredito === 0) {
    return {
      totalCredito: 0,
      prazoMedio: 0,
      taxaAdmMedia: 0,
      fundoReservaMedia: 0,
      seguroMedio: 0,
      lanceTotalPct: 0,
      contemplacaoMedia: 0,
      reajusteMedio: 0,
      numGrupos: grupos.length,
      grupos,
    };
  }

  return {
    totalCredito,
    prazoMedio: somaPrazoPonderado / totalCredito,
    taxaAdmMedia: somaTaxaAdmPonderada / totalCredito,
    fundoReservaMedia: somaFundoReservaPonderado / totalCredito,
    seguroMedio: somaSeguroPonderado / totalCredito,
    lanceTotalPct: somaLancePonderado / totalCredito,
    contemplacaoMedia: somaContempPonderado / totalCredito,
    reajusteMedio: somaReajustePonderado / totalCredito,
    numGrupos: grupos.length,
    grupos,
  };
}

// =============================================================================
// QUANTITY-WEIGHTED CONSOLIDATION
// =============================================================================

/**
 * Consolidates groups with explicit quantity weights.
 * Useful when a client holds multiple quotas of the same group.
 *
 * @param grupos - array of group data with quantities
 * @returns consolidated result using quantity-weighted values
 */
export function consolidarCotasComQuantidade(
  grupos: Array<GrupoInput & { quantidade: number }>
): ConsolidacaoResult & { quantidadeTotal: number } {
  if (!grupos || grupos.length === 0) {
    return {
      totalCredito: 0,
      prazoMedio: 0,
      taxaAdmMedia: 0,
      fundoReservaMedia: 0,
      quantidadeTotal: 0,
    };
  }

  let totalCredito = 0;
  let quantidadeTotal = 0;
  let somaPrazoPonderado = 0;
  let somaTaxaAdmPonderada = 0;
  let somaFundoReservaPonderado = 0;

  for (const g of grupos) {
    const valor = (g.valor ?? 0) * (g.quantidade ?? 1);
    const qtd = g.quantidade ?? 1;
    totalCredito += valor;
    quantidadeTotal += qtd;
    somaPrazoPonderado += (g.prazo ?? 0) * valor;
    somaTaxaAdmPonderada += (g.taxaAdm ?? 0) * valor;
    somaFundoReservaPonderado += (g.fundoReserva ?? 0) * valor;
  }

  if (totalCredito === 0) {
    return {
      totalCredito: 0,
      prazoMedio: 0,
      taxaAdmMedia: 0,
      fundoReservaMedia: 0,
      quantidadeTotal,
    };
  }

  return {
    totalCredito,
    prazoMedio: somaPrazoPonderado / totalCredito,
    taxaAdmMedia: somaTaxaAdmPonderada / totalCredito,
    fundoReservaMedia: somaFundoReservaPonderado / totalCredito,
    quantidadeTotal,
  };
}

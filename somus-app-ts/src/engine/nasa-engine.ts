/**
 * nasa-engine.ts - Full consortium calculation engine
 * Complete port of nasa_engine_hd.py (NasaEngineHD class)
 * Somus Capital - Mesa de Produtos
 *
 * Replaces the NASA NOVA HD VPL spreadsheet (21 tabs, 150+ named ranges).
 * Produces IDENTICAL numerical results to the Python version.
 */

import { npv, irr, pmt, annualFromMonthly, monthlyFromAnnual, goalSeek } from './irr-npv';

// =============================================================================
// CONSTANTS
// =============================================================================

export const FREQ_MAP: Record<string, number> = {
  Mensal: 1,
  Bimestral: 2,
  Trimestral: 3,
  Semestral: 6,
  Anual: 12,
};

const IOF_DIARIO = 0.000082; // 0.0082% a.d.
const IOF_ADICIONAL = 0.0038; // 0.38% on principal
const MAX_MESES = 420; // 35 years

// =============================================================================
// TYPES
// =============================================================================

export interface NasaConfig {
  /** Base for insurance: "saldo_devedor" or "valor_credito" */
  seguroBase: string;
  /** Timing of TA anticipation: "junto_1a_parcela", "na_contemplacao", "diluida" */
  momentoAntecipacaoTa: string;
  /** Timing of embedded bid: "na_contemplacao" or "desde_inicio" */
  momentoLanceEmbutido: string;
  /** Base for embedded bid: "credito_original", "original+txadm", "atualizado", "saldo_devedor" */
  baseCalculoLanceEmbutido: string;
  /** Base for free bid */
  baseCalculoLanceLivre: string;
  /** Update credit value by readjustment */
  atualizarValorCredito: boolean;
  /** Base for reserve fund */
  baseCalculoFundoReserva: string;
  /** Method for IRR: "fluxo_original" or "fluxo_ajustado" */
  metodoTir: string;
}

export interface PeriodoDistribuicao {
  start: number;
  end: number;
  fc_pct: number;
  ta_pct: number;
  fr_pct: number;
}

export interface CustoAcessorio {
  descricao: string;
  valor: number;
  momento: number;
}

export interface NasaParams {
  valor_credito: number;
  prazo_meses: number;
  taxa_adm_pct: number;
  fundo_reserva_pct: number;
  momento_contemplacao: number;

  periodos: PeriodoDistribuicao[];

  lance_embutido_pct: number;
  lance_livre_pct: number;
  lance_embutido_valor: number;
  lance_livre_valor: number;

  reajuste_pre_pct: number;
  reajuste_pos_pct: number;
  reajuste_pre_freq: string;
  reajuste_pos_freq: string;

  seguro_vida_pct: number;
  seguro_vida_inicio: number;

  antecipacao_ta_pct: number;
  antecipacao_ta_parcelas: number;

  taxa_vp_credito: number;
  tma: number;
  alm_anual: number;
  hurdle_anual: number;

  custos_acessorios: CustoAcessorio[];

  // Legacy compat fields
  valor_carta?: number;
  taxa_adm?: number;
  fundo_reserva?: number;
  seguro?: number;
  prazo_contemp?: number;
  parcela_red_pct?: number;
  lance_livre_pct_compat?: number;
  lance_embutido_pct_compat?: number;
  correcao_anual?: number;
}

export interface FluxoMensal {
  mes: number;
  meses_restantes: number;
  valor_base_fc: number;
  lance_embutido: number;
  lance_livre: number;
  valor_base_final: number;
  pct_mensal_fc: number;
  pct_acum_fc: number;
  amortizacao: number;
  saldo_principal: number;
  taxa_adm_antecipada: number;
  pct_ta_mensal: number;
  pct_ta_acum: number;
  valor_parcela_ta: number;
  pct_fr_mensal: number;
  pct_fr_acum: number;
  pct_fr_base: number;
  pct_fr_calc: number;
  fr_saldo: number;
  valor_fundo_reserva: number;
  valor_parcela: number;
  saldo_devedor: number;
  peso_parcela: number;
  pct_reajuste: number;
  pct_reajuste_acum: number;
  parcela_apos_reajuste: number;
  saldo_devedor_reajustado: number;
  seguro_vida: number;
  parcela_com_seguro: number;
  outros_custos: number;
  carta_credito_original: number;
  carta_credito_reajustada: number;
  fluxo_caixa: number;
  fluxo_caixa_tir: number;
  credito_recebido: number;
  fator_reajuste: number;
}

export interface FluxoTotais {
  total_pago: number;
  total_fundo_comum: number;
  total_taxa_adm: number;
  total_fundo_reserva: number;
  total_seguro: number;
  total_custos_acessorios: number;
  carta_liquida: number;
  lance_embutido_valor: number;
  lance_livre_valor: number;
}

export interface FluxoMetricas {
  tir_mensal: number;
  tir_anual: number;
  cet_anual: number;
  parcela_media: number;
  parcela_maxima: number;
  parcela_minima: number;
  custo_total_pct: number;
}

export interface FluxoResult {
  fluxo: FluxoMensal[];
  cashflow: number[];
  cashflow_tir: number[];
  totais: FluxoTotais;
  metricas: FluxoMetricas;
  // Legacy compat
  fluxo_mensal: FluxoMensal[];
  cashflow_consorcio: number[];
  total_pago: number;
  carta_liquida: number;
  lance_livre_valor: number;
  lance_embutido_valor: number;
}

export interface VPLResult {
  b0: number;
  h0: number;
  d0: number;
  pv_pos_t: number;
  pv_pos_t_at_contemp: number;
  delta_vpl: number;
  cria_valor: boolean;
  break_even_lance: number;
  tir_mensal: number;
  tir_anual: number;
  cet_anual: number;
  vpl_total: number;
  pv_pre_t_detail: Array<{ mes: number; valor: number; pv: number }>;
  pv_pos_t_detail: Array<{ mes: number; valor: number; pv: number }>;
}

export interface ParcelaInfo {
  mes: number;
  fundo_comum: number;
  taxa_adm: number;
  fundo_reserva: number;
  parcela_base: number;
  reajuste: number;
  parcela_reajustada: number;
  seguro: number;
  parcela_total: number;
  outros_custos: number;
  desembolso_total: number;
  saldo_devedor: number;
  lance_embutido: number;
  lance_livre: number;
  credito_recebido: number;
}

export interface ResumoCliente {
  valor_credito: number;
  prazo_meses: number;
  momento_contemplacao: number;
  carta_liquida: number;
  lance_embutido: number;
  lance_embutido_pct: number;
  lance_livre: number;
  lance_livre_pct: number;
  lance_total: number;
  lance_total_pct: number;
  primeira_parcela: number;
  ultima_parcela: number;
  parcela_media: number;
  parcela_maxima: number;
  parcela_minima: number;
  total_pago: number;
  total_fundo_comum: number;
  total_taxa_adm: number;
  total_fundo_reserva: number;
  total_seguro: number;
  total_custos_acessorios: number;
  custo_total_pct: number;
  taxa_adm_pct: number;
  fundo_reserva_pct: number;
  seguro_pct: number;
  tir_mensal: number;
  tir_anual: number;
  cet_anual: number;
  b0: number;
  h0: number;
  d0: number;
  pv_pos_t: number;
  delta_vpl: number;
  cria_valor: boolean;
  vpl_total: number;
  break_even_lance: number;
  reajuste_pre_pct: number;
  reajuste_pos_pct: number;
  reajuste_pre_freq: string;
  reajuste_pos_freq: string;
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
  // Optional fields added by calcularCreditoLance
  cashflow_antecipado?: number[];
  valor_antecipacao?: number;
  mes_antecipacao?: number;
  custos_iniciais?: number;
}

// =============================================================================
// DEFAULT CONFIGURATION
// =============================================================================

export const DEFAULT_CONFIG: NasaConfig = {
  seguroBase: 'saldo_devedor',
  momentoAntecipacaoTa: 'junto_1a_parcela',
  momentoLanceEmbutido: 'na_contemplacao',
  baseCalculoLanceEmbutido: 'credito_original',
  baseCalculoLanceLivre: 'credito_original',
  atualizarValorCredito: true,
  baseCalculoFundoReserva: 'credito_original',
  metodoTir: 'fluxo_original',
};

export const DEFAULT_PARAMS: NasaParams = {
  valor_credito: 500000.0,
  prazo_meses: 200,
  taxa_adm_pct: 20.0,
  fundo_reserva_pct: 3.0,
  momento_contemplacao: 36,
  periodos: [
    { start: 1, end: 200, fc_pct: 1.0, ta_pct: 100.0, fr_pct: 100.0 },
  ],
  lance_embutido_pct: 0.0,
  lance_livre_pct: 0.0,
  lance_embutido_valor: 0.0,
  lance_livre_valor: 0.0,
  reajuste_pre_pct: 0.0,
  reajuste_pos_pct: 0.0,
  reajuste_pre_freq: 'Anual',
  reajuste_pos_freq: 'Anual',
  seguro_vida_pct: 0.0,
  seguro_vida_inicio: 1,
  antecipacao_ta_pct: 0.0,
  antecipacao_ta_parcelas: 1,
  taxa_vp_credito: 0.0,
  tma: 0.01,
  alm_anual: 12.0,
  hurdle_anual: 12.0,
  custos_acessorios: [],
};

// =============================================================================
// HELPER: fill defaults from partial input
// =============================================================================

function resolveParams(input: Record<string, any>): Record<string, any> {
  return { ...DEFAULT_PARAMS, ...input };
}

// =============================================================================
// NASA ENGINE HD
// =============================================================================

export class NasaEngine {
  config: NasaConfig;

  constructor(config?: Partial<NasaConfig>) {
    this.config = { ...DEFAULT_CONFIG, ...(config ?? {}) };
  }

  // -----------------------------------------------------------------
  // Internal helpers
  // -----------------------------------------------------------------

  /**
   * Resolve lance value (absolute takes priority over percentage).
   */
  private resolveLance(
    params: Record<string, any>,
    tipo: 'embutido' | 'livre'
  ): number {
    const credito = params.valor_credito ?? 0;
    if (tipo === 'embutido') {
      const val = params.lance_embutido_valor ?? 0;
      if (val && val > 0) return val;
      const pctVal = (params.lance_embutido_pct ?? 0) / 100;
      return credito * pctVal;
    } else {
      const val = params.lance_livre_valor ?? 0;
      if (val && val > 0) return val;
      const pctVal = (params.lance_livre_pct ?? 0) / 100;
      return credito * pctVal;
    }
  }

  /**
   * Build accumulated readjustment factor vector for months 0..prazo.
   */
  private resolveReajusteSchedule(
    prazo: number,
    contemp: number,
    params: Record<string, any>
  ): number[] {
    const preRate = (params.reajuste_pre_pct ?? 0) / 100;
    const posRate = (params.reajuste_pos_pct ?? 0) / 100;
    const preFreq = FREQ_MAP[params.reajuste_pre_freq ?? 'Anual'] ?? 12;
    const posFreq = FREQ_MAP[params.reajuste_pos_freq ?? 'Anual'] ?? 12;

    const fatores = new Array(prazo + 1).fill(1.0);
    let acum = 0.0;

    for (let m = 1; m <= prazo; m++) {
      let rate: number;
      let freq: number;
      if (m <= contemp) {
        rate = preRate;
        freq = preFreq;
      } else {
        rate = posRate;
        freq = posFreq;
      }

      if (freq > 0 && m % freq === 0) {
        acum = (1 + acum) * (1 + rate) - 1;
      }

      fatores[m] = 1 + acum;
    }

    return fatores;
  }

  /**
   * Build per-period distribution using GoalSeek.
   * Returns arrays of monthly % for FC, TA, FR (indices 0..prazo).
   */
  private buildPeriodDistribution(
    params: Record<string, any>
  ): [number[], number[], number[]] {
    const prazo = params.prazo_meses ?? 200;
    const periodos: PeriodoDistribuicao[] = params.periodos ?? [];

    if (!periodos || periodos.length === 0) {
      const pctVal = prazo > 0 ? 1.0 / prazo : 0;
      const uniform = [0.0, ...new Array(prazo).fill(pctVal)];
      return [uniform, [...uniform], [...uniform]];
    }

    const fcDist = this.solveDistribution(prazo, periodos, 'fc_pct');
    const taDist = this.solveDistribution(prazo, periodos, 'ta_pct');
    const frDist = this.solveDistribution(prazo, periodos, 'fr_pct');

    return [fcDist, taDist, frDist];
  }

  /**
   * Solve distribution for one component using GoalSeek.
   */
  private solveDistribution(
    prazo: number,
    periodos: PeriodoDistribuicao[],
    key: 'fc_pct' | 'ta_pct' | 'fr_pct'
  ): number[] {
    const pesosBrutos = new Array(prazo + 1).fill(0.0);
    for (const p of periodos) {
      const start = Math.max(1, p.start ?? 1);
      const end = Math.min(prazo, p.end ?? prazo);
      let peso: number = (p as any)[key] ?? 1.0;
      if ((key === 'ta_pct' || key === 'fr_pct') && peso > 1) {
        peso = peso / 100.0;
      }
      for (let m = start; m <= end; m++) {
        pesosBrutos[m] = peso;
      }
    }

    let totalPeso = 0;
    for (let m = 1; m <= prazo; m++) {
      totalPeso += pesosBrutos[m];
    }

    if (totalPeso === 0) {
      const pctVal = prazo > 0 ? 1.0 / prazo : 0;
      return [0.0, ...new Array(prazo).fill(pctVal)];
    }

    const somaFunc = (mult: number): number => {
      let s = 0;
      for (let m = 1; m <= prazo; m++) {
        s += mult * pesosBrutos[m];
      }
      return s;
    };

    const multiplicador = goalSeek(
      somaFunc,
      1.0,
      0.0,
      10.0 / Math.max(totalPeso, 1e-10),
      1e-12,
      500
    );

    const dist = new Array(prazo + 1).fill(0.0);
    for (let m = 1; m <= prazo; m++) {
      dist[m] = multiplicador * pesosBrutos[m];
    }

    return dist;
  }

  // -----------------------------------------------------------------
  // 1. FULL CASH FLOW (35 columns)
  // -----------------------------------------------------------------

  calcularFluxoCompleto(inputParams: Record<string, any>): FluxoResult {
    const params = resolveParams(inputParams);
    const credito = params.valor_credito ?? 0;
    const prazo = params.prazo_meses ?? 200;
    const contemp = params.momento_contemplacao ?? 36;
    const taxaAdmTotal = (params.taxa_adm_pct ?? 20.0) / 100;
    const fundoReservaTotal = (params.fundo_reserva_pct ?? 3.0) / 100;
    const seguroPct = (params.seguro_vida_pct ?? 0) / 100;
    const seguroInicio = params.seguro_vida_inicio ?? 1;

    const lanceEmbVal = this.resolveLance(params, 'embutido');
    const lanceLivreVal = this.resolveLance(params, 'livre');

    const antecipacaoTaPct = (params.antecipacao_ta_pct ?? 0) / 100;
    const antecipacaoTaParcelas = Math.max(1, params.antecipacao_ta_parcelas ?? 1);

    const custos: CustoAcessorio[] = params.custos_acessorios ?? [];

    const [fcDist, taDist, frDist] = this.buildPeriodDistribution(params);
    const fatoresReajuste = this.resolveReajusteSchedule(prazo, contemp, params);

    const valorTaTotal = credito * taxaAdmTotal;
    const valorFrTotal = credito * fundoReservaTotal;

    const taAntecipadaTotal = valorTaTotal * antecipacaoTaPct;
    const taAntecipadaPorParcela =
      antecipacaoTaParcelas > 0 ? taAntecipadaTotal / antecipacaoTaParcelas : 0;

    const fluxo: FluxoMensal[] = [];
    let saldoPrincipal = credito;
    let saldoDevedor = credito;
    let creditoReajustado = credito;
    let totalPago = 0.0;
    let totalFc = 0.0;
    let totalTa = 0.0;
    let totalFr = 0.0;
    let totalSeguro = 0.0;
    let totalCustosAcessorios = 0.0;
    let cumFcPct = 0.0;

    const cashflow: number[] = [];
    const cashflowTir: number[] = [];

    for (let m = 0; m <= prazo; m++) {
      if (m === 0) {
        const row: FluxoMensal = {
          mes: 0,
          meses_restantes: prazo,
          valor_base_fc: credito,
          lance_embutido: 0.0,
          lance_livre: 0.0,
          valor_base_final: credito,
          pct_mensal_fc: 0.0,
          pct_acum_fc: 0.0,
          amortizacao: 0.0,
          saldo_principal: credito,
          taxa_adm_antecipada: 0.0,
          pct_ta_mensal: 0.0,
          pct_ta_acum: 0.0,
          valor_parcela_ta: 0.0,
          pct_fr_mensal: 0.0,
          pct_fr_acum: 0.0,
          pct_fr_base: 0.0,
          pct_fr_calc: 0.0,
          fr_saldo: 0.0,
          valor_fundo_reserva: 0.0,
          valor_parcela: 0.0,
          saldo_devedor: credito,
          peso_parcela: 0.0,
          pct_reajuste: 0.0,
          pct_reajuste_acum: 0.0,
          parcela_apos_reajuste: 0.0,
          saldo_devedor_reajustado: credito,
          seguro_vida: 0.0,
          parcela_com_seguro: 0.0,
          outros_custos: 0.0,
          carta_credito_original: credito,
          carta_credito_reajustada: credito,
          fluxo_caixa: 0.0,
          fluxo_caixa_tir: 0.0,
          credito_recebido: 0.0,
          fator_reajuste: 1.0,
        };

        saldoPrincipal = credito;
        saldoDevedor = credito;

        fluxo.push(row);
        cashflow.push(0.0);
        cashflowTir.push(0.0);
        continue;
      }

      // --- Col D: Base Value Fundo Comum ---
      let valorBaseFc = credito;
      if (this.config.atualizarValorCredito) {
        valorBaseFc = credito * fatoresReajuste[m];
        creditoReajustado = valorBaseFc;
      }

      // --- Col E: Lance Embutido ---
      let lanceEmbMes = 0.0;
      if (this.config.momentoLanceEmbutido === 'na_contemplacao') {
        if (m === contemp) {
          lanceEmbMes = -lanceEmbVal;
        }
      } else {
        if (m === 1) {
          lanceEmbMes = -lanceEmbVal;
        }
      }

      // --- Col F: Lance Livre ---
      let lanceLivreMes = 0.0;
      if (m === contemp) {
        lanceLivreMes = -lanceLivreVal;
      }

      // --- Col G: Final Base Value ---
      const valorBaseFinal = valorBaseFc + lanceEmbMes + lanceLivreMes;

      // --- Col H: % Monthly FC ---
      const pctMensalFc = m < fcDist.length ? fcDist[m] : 0.0;

      // --- Col I: % Accumulated FC ---
      cumFcPct += pctMensalFc;

      // --- Col J: Amortization ---
      const prevPctAcum = fluxo[m - 1].pct_acum_fc;
      const prevSaldo = saldoPrincipal;

      const denom = 1 - prevPctAcum;
      let amort = 0.0;
      if (Math.abs(denom) > 1e-12 && pctMensalFc > 0) {
        amort = -(prevSaldo / denom) * pctMensalFc;
      }

      amort += lanceEmbMes + lanceLivreMes;

      // --- Col K: Principal Balance ---
      saldoPrincipal = prevSaldo + amort;
      if (Math.abs(saldoPrincipal) < 0.01) {
        saldoPrincipal = 0.0;
      }

      // --- Col L: Anticipated TA ---
      let taAntecipadaMes = 0.0;
      if (antecipacaoTaPct > 0) {
        if (this.config.momentoAntecipacaoTa === 'junto_1a_parcela') {
          if (m >= 1 && m <= antecipacaoTaParcelas) {
            taAntecipadaMes = taAntecipadaPorParcela;
          }
        } else if (this.config.momentoAntecipacaoTa === 'na_contemplacao') {
          if (m === contemp) {
            taAntecipadaMes = taAntecipadaTotal;
          }
        } else if (this.config.momentoAntecipacaoTa === 'diluida') {
          if (m >= 1 && m <= prazo) {
            taAntecipadaMes = taAntecipadaTotal / prazo;
          }
        }
      }

      // --- Col M/N: % TA monthly and accumulated ---
      const pctTaMensal = m < taDist.length ? taDist[m] : 0.0;
      const pctTaEfetivo = pctTaMensal * (1 - antecipacaoTaPct);
      const taAcumPrev = fluxo[m - 1].pct_ta_acum;
      const pctTaAcum = taAcumPrev + pctTaMensal;

      // --- Col O: TA Installment Value ---
      const valorParcelaTa = valorBaseFc * pctTaEfetivo + taAntecipadaMes;

      // --- Col P-T: Reserve Fund ---
      const pctFrMensal = m < frDist.length ? frDist[m] : 0.0;
      const frAcumPrev = fluxo[m - 1].pct_fr_acum;
      const pctFrAcum = frAcumPrev + pctFrMensal;

      // Col U: Reserve Fund Value (linearly distributes total FR)
      const valorFundoReserva = valorFrTotal * pctFrMensal;

      // --- Col V: Total Installment ---
      const amortPura = amort - lanceEmbMes - lanceLivreMes;
      const valorParcela = Math.abs(amortPura) + valorParcelaTa + valorFundoReserva;

      // --- Col W: Outstanding Balance ---
      saldoDevedor =
        Math.abs(saldoPrincipal) +
        valorTaTotal * (1 - pctTaAcum) +
        valorFrTotal * (1 - pctFrAcum);

      // --- Col X: Installment Weight ---
      const peso = credito > 0 ? valorParcela / credito : 0;

      // --- Col Y-Z: Readjustment ---
      const fatorReaj = fatoresReajuste[m];
      const pctReajPeriod =
        m > 0 && fatoresReajuste[m - 1] > 0
          ? fatorReaj / fatoresReajuste[m - 1] - 1
          : 0;

      // --- Col AA: Installment After Readjustment ---
      const parcelaReajustada = valorParcela * fatorReaj;

      // --- Col AB: Readjusted Outstanding Balance ---
      const saldoDevReaj = saldoDevedor * fatorReaj;

      // --- Col AC: Life Insurance ---
      let seguroMes = 0.0;
      if (seguroPct > 0 && m >= seguroInicio) {
        if (this.config.seguroBase === 'saldo_devedor') {
          seguroMes = Math.abs(saldoDevReaj) * seguroPct;
        } else {
          seguroMes = creditoReajustado * seguroPct;
        }
      }

      // --- Col AD: Installment with Insurance ---
      const parcelaComSeguro = parcelaReajustada + seguroMes;

      // --- Col AE: Other Costs ---
      let custoMes = 0;
      for (const c of custos) {
        if ((c.momento ?? -1) === m) {
          custoMes += c.valor ?? 0;
        }
      }
      totalCustosAcessorios += custoMes;

      // --- Col AF-AG: Credit Letter ---
      const cartaCreditoReajustada = credito * fatorReaj;

      // --- Credit received ---
      let creditoRecebido = 0.0;
      if (m === contemp) {
        const cartaLiquida = credito - lanceEmbVal;
        creditoRecebido = this.config.atualizarValorCredito
          ? cartaLiquida * fatorReaj
          : cartaLiquida;
      }

      // --- Col AH: Cash Flow ---
      const desembolso = parcelaComSeguro + custoMes;
      const lanceLivreReal =
        lanceLivreMes !== 0 ? Math.abs(lanceLivreMes) * fatorReaj : 0;
      const lanceEmbReal =
        lanceEmbMes !== 0 ? Math.abs(lanceEmbMes) * fatorReaj : 0;

      const fluxoMes = creditoRecebido - desembolso - lanceLivreReal;

      // --- Col AI: Cash Flow for IRR ---
      let fluxoTir: number;
      if (this.config.metodoTir === 'fluxo_ajustado') {
        fluxoTir = fluxoMes;
      } else {
        fluxoTir = creditoRecebido - desembolso - lanceLivreReal;
      }

      // Accumulators
      totalPago += desembolso + lanceLivreReal;
      totalFc += Math.abs(amortPura) * fatorReaj;
      totalTa += valorParcelaTa * fatorReaj;
      totalFr += valorFundoReserva * fatorReaj;
      totalSeguro += seguroMes;

      const row: FluxoMensal = {
        mes: m,
        meses_restantes: prazo - m,
        valor_base_fc: valorBaseFc,
        lance_embutido: lanceEmbMes,
        lance_livre: lanceLivreMes,
        valor_base_final: valorBaseFinal,
        pct_mensal_fc: pctMensalFc,
        pct_acum_fc: cumFcPct,
        amortizacao: amort,
        saldo_principal: saldoPrincipal,
        taxa_adm_antecipada: taAntecipadaMes,
        pct_ta_mensal: pctTaMensal,
        pct_ta_acum: pctTaAcum,
        valor_parcela_ta: valorParcelaTa,
        pct_fr_mensal: pctFrMensal,
        pct_fr_acum: pctFrAcum,
        pct_fr_base: fundoReservaTotal,
        pct_fr_calc: pctFrMensal * fundoReservaTotal,
        fr_saldo: 0.0,
        valor_fundo_reserva: valorFundoReserva,
        valor_parcela: valorParcela,
        saldo_devedor: saldoDevedor,
        peso_parcela: peso,
        pct_reajuste: pctReajPeriod,
        pct_reajuste_acum: fatorReaj - 1,
        parcela_apos_reajuste: parcelaReajustada,
        saldo_devedor_reajustado: saldoDevReaj,
        seguro_vida: seguroMes,
        parcela_com_seguro: parcelaComSeguro,
        outros_custos: custoMes,
        carta_credito_original: credito,
        carta_credito_reajustada: cartaCreditoReajustada,
        fluxo_caixa: fluxoMes,
        fluxo_caixa_tir: fluxoTir,
        credito_recebido: creditoRecebido,
        fator_reajuste: fatorReaj,
      };

      fluxo.push(row);
      cashflow.push(fluxoMes);
      cashflowTir.push(fluxoTir);
    }

    // --- Metrics ---
    const tirM = irr(cashflowTir, 0.005);
    const tirA = annualFromMonthly(tirM);

    const parcelasReaj = fluxo
      .filter((r) => r.mes > 0)
      .map((r) => r.parcela_apos_reajuste);
    const parcelaMedia =
      parcelasReaj.length > 0
        ? parcelasReaj.reduce((a, b) => a + b, 0) / parcelasReaj.length
        : 0;
    const parcelaMax = parcelasReaj.length > 0 ? Math.max(...parcelasReaj) : 0;
    const positiveParc = parcelasReaj.filter((p) => p > 0);
    const parcelaMin = positiveParc.length > 0 ? Math.min(...positiveParc) : 0;

    const cartaLiquida = credito - lanceEmbVal;

    return {
      fluxo,
      cashflow,
      cashflow_tir: cashflowTir,
      totais: {
        total_pago: totalPago,
        total_fundo_comum: totalFc,
        total_taxa_adm: totalTa,
        total_fundo_reserva: totalFr,
        total_seguro: totalSeguro,
        total_custos_acessorios: totalCustosAcessorios,
        carta_liquida: cartaLiquida,
        lance_embutido_valor: lanceEmbVal,
        lance_livre_valor: lanceLivreVal,
      },
      metricas: {
        tir_mensal: tirM,
        tir_anual: tirA,
        cet_anual: tirA,
        parcela_media: parcelaMedia,
        parcela_maxima: parcelaMax,
        parcela_minima: parcelaMin,
        custo_total_pct: credito > 0 ? (totalPago / credito) * 100 : 0,
      },
      fluxo_mensal: fluxo,
      cashflow_consorcio: cashflow,
      total_pago: totalPago,
      carta_liquida: cartaLiquida,
      lance_livre_valor: lanceLivreVal,
      lance_embutido_valor: lanceEmbVal,
    };
  }

  // -----------------------------------------------------------------
  // 2. VPL HD (Goal-Based with dual rates)
  // -----------------------------------------------------------------

  calcularVPLHD(
    inputParams: Record<string, any>,
    fluxoInput?: FluxoResult | null
  ): VPLResult {
    const params = resolveParams(inputParams);
    const fluxoResult = fluxoInput ?? this.calcularFluxoCompleto(params);

    const almA = (params.alm_anual ?? 12.0) / 100;
    const hurdleA = (params.hurdle_anual ?? 12.0) / 100;
    const almM = monthlyFromAnnual(almA);
    const hurdleM = monthlyFromAnnual(hurdleA);

    const contemp = params.momento_contemplacao ?? 36;
    const cartaLiquida = fluxoResult.carta_liquida ?? fluxoResult.totais?.carta_liquida ?? 0;

    const fluxoMensal = fluxoResult.fluxo ?? fluxoResult.fluxo_mensal ?? [];

    // B0: PV of credit received at contemplation
    let creditoRecebido = 0;
    for (const f of fluxoMensal) {
      const cr = f.credito_recebido ?? 0;
      if (cr > 0) {
        creditoRecebido = cr;
        break;
      }
    }
    if (creditoRecebido === 0) {
      creditoRecebido = cartaLiquida;
    }

    const b0 = contemp > 0
      ? creditoRecebido / Math.pow(1 + almM, contemp)
      : creditoRecebido;

    // H0: PV of pre-contemplation payments + bids
    let h0 = 0.0;
    const pvPreTDetail: Array<{ mes: number; valor: number; pv: number }> = [];
    for (const f of fluxoMensal) {
      const mes = f.mes ?? 0;
      if (mes > 0 && mes <= contemp) {
        let pagamento = f.parcela_com_seguro ?? 0;
        pagamento += f.outros_custos ?? 0;
        const lance = Math.abs(f.lance_livre ?? 0);
        const totalM = pagamento + lance;
        const pvM = totalM / Math.pow(1 + almM, mes);
        h0 += pvM;
        pvPreTDetail.push({ mes, valor: totalM, pv: pvM });
      }
    }

    const d0 = b0 - h0;

    // PV of post-contemplation payments (discounted at hurdle)
    let pvPosTAtContemp = 0.0;
    const pvPosTDetail: Array<{ mes: number; valor: number; pv: number }> = [];
    for (const f of fluxoMensal) {
      const mes = f.mes ?? 0;
      if (mes > contemp) {
        let pagamento = f.parcela_com_seguro ?? 0;
        pagamento += f.outros_custos ?? 0;
        const mesesApos = mes - contemp;
        const pvM = pagamento / Math.pow(1 + hurdleM, mesesApos);
        pvPosTAtContemp += pvM;
        pvPosTDetail.push({ mes, valor: pagamento, pv: pvM });
      }
    }

    const pvPosT = contemp > 0
      ? pvPosTAtContemp / Math.pow(1 + almM, contemp)
      : pvPosTAtContemp;

    const deltaVpl = d0 - pvPosT;
    const criaValor = deltaVpl >= 0;

    const cf = fluxoResult.cashflow_tir ?? fluxoResult.cashflow ?? [];
    const tirM = irr(cf, 0.005);
    const tirA = annualFromMonthly(tirM);
    const vplTotal = npv(almM, cf);

    const beLance = this.buscarBreakEvenLanceHD(params, almM, hurdleM);

    return {
      b0,
      h0,
      d0,
      pv_pos_t: pvPosT,
      pv_pos_t_at_contemp: pvPosTAtContemp,
      delta_vpl: deltaVpl,
      cria_valor: criaValor,
      break_even_lance: beLance,
      tir_mensal: tirM,
      tir_anual: tirA,
      cet_anual: tirA,
      vpl_total: vplTotal,
      pv_pre_t_detail: pvPreTDetail,
      pv_pos_t_detail: pvPosTDetail,
    };
  }

  /**
   * Binary search for lance that zeros Delta VPL.
   */
  private buscarBreakEvenLanceHD(
    params: Record<string, any>,
    almM: number,
    hurdleM: number
  ): number {
    const self = this;

    const calcDelta = (lancePct: number): number => {
      const p = { ...params, lance_livre_pct: lancePct, lance_livre_valor: 0 };
      const fl = self.calcularFluxoCompleto(p as any);
      const contemp = (p as any).momento_contemplacao ?? (p as any).momentoContemplacao ?? 36;
      const cartaLiq = fl.carta_liquida;
      const fluxoArr = fl.fluxo;

      let creditoRec = 0;
      for (const f of fluxoArr) {
        const cr = f.credito_recebido ?? 0;
        if (cr > 0) { creditoRec = cr; break; }
      }
      if (creditoRec === 0) creditoRec = cartaLiq;

      const b0 = contemp > 0
        ? creditoRec / Math.pow(1 + almM, contemp)
        : creditoRec;

      let h0 = 0.0;
      for (const f of fluxoArr) {
        const mes = f.mes;
        if (mes > 0 && mes <= contemp) {
          const pag = (f.parcela_com_seguro ?? 0) + (f.outros_custos ?? 0);
          const lance = Math.abs(f.lance_livre ?? 0);
          h0 += (pag + lance) / Math.pow(1 + almM, mes);
        }
      }

      const d0 = b0 - h0;

      let pvPos = 0.0;
      for (const f of fluxoArr) {
        const mes = f.mes;
        if (mes > contemp) {
          const pag = (f.parcela_com_seguro ?? 0) + (f.outros_custos ?? 0);
          pvPos += pag / Math.pow(1 + hurdleM, mes - contemp);
        }
      }
      if (contemp > 0) {
        pvPos /= Math.pow(1 + almM, contemp);
      }

      return d0 - pvPos;
    };

    try {
      return goalSeek(calcDelta, 0.0, 0.0, 90.0, 0.01, 500);
    } catch {
      return 0.0;
    }
  }

  // -----------------------------------------------------------------
  // 3. FINANCING (SAC/Price with IOF)
  // -----------------------------------------------------------------

  calcularFinanciamento(params: Record<string, any>): FinanciamentoResult {
    const valor = params.valor ?? 0;
    const prazo = params.prazo_meses ?? 0;
    const taxa = (params.taxa_mensal_pct ?? 0) / 100;
    const metodo = (params.metodo ?? 'price').toLowerCase();
    const carencia = params.carencia ?? 0;
    const calcIof = params.calcular_iof !== undefined ? params.calcular_iof : true;
    const custosAdd: CustoAcessorio[] = params.custos_adicionais ?? [];

    let saldo = valor;
    const parcelas: FinanciamentoParcelaRow[] = [];
    const cashflowFin: number[] = [valor];
    let totalPagoFin = 0.0;
    let totalJuros = 0.0;
    let totalAmort = 0.0;

    const prazoAmort = prazo - carencia;

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
        amortM = 0.0;
        parcelaM = juros;
      } else if (metodo === 'price') {
        parcelaM = pmtVal;
        amortM = parcelaM - juros;
      } else {
        amortM = prazoAmort > 0 ? valor / prazoAmort : 0;
        parcelaM = amortM + juros;
      }

      saldo -= amortM;
      saldo = Math.max(0, saldo);

      totalPagoFin += parcelaM;
      totalJuros += juros;
      totalAmort += amortM;

      parcelas.push({
        mes: m,
        parcela: parcelaM,
        juros,
        amortizacao: amortM,
        saldo,
      });
      cashflowFin.push(-parcelaM);
    }

    let iofTotal = 0.0;
    if (calcIof) {
      iofTotal = this.calcularIOFInternal(valor, parcelas);
    }

    let totalCustosAdd = 0;
    for (const c of custosAdd) {
      totalCustosAdd += c.valor ?? 0;
    }
    for (const c of custosAdd) {
      const mCusto = c.momento ?? 0;
      if (mCusto >= 0 && mCusto <= prazo) {
        if (mCusto === 0) {
          cashflowFin[0] -= c.valor ?? 0;
        } else if (mCusto <= cashflowFin.length - 1) {
          cashflowFin[mCusto] -= c.valor ?? 0;
        }
      }
    }

    const cfCet = [...cashflowFin];
    if (calcIof) {
      cfCet[0] = valor - iofTotal - totalCustosAdd;
    }
    const tirM = irr(cfCet, 0.008);
    const tirA = annualFromMonthly(tirM);

    return {
      parcelas,
      cashflow: cashflowFin,
      total_pago: totalPagoFin,
      total_juros: totalJuros,
      total_amortizado: totalAmort,
      valor,
      iof: iofTotal,
      custos_adicionais: totalCustosAdd,
      custo_efetivo_total: totalPagoFin + iofTotal + totalCustosAdd,
      tir_mensal: tirM,
      tir_anual: tirA,
      cet_anual: tirA,
    };
  }

  private calcularIOFInternal(
    valorPrincipal: number,
    parcelas: Array<{ mes: number; amortizacao: number }>
  ): number {
    const iofAdicional = valorPrincipal * IOF_ADICIONAL;
    let iofDiarioTotal = 0.0;
    for (const p of parcelas) {
      let dias = p.mes * 30;
      dias = Math.min(dias, 365);
      iofDiarioTotal += p.amortizacao * IOF_DIARIO * dias;
    }
    return iofAdicional + iofDiarioTotal;
  }

  // -----------------------------------------------------------------
  // 4. LANCE FINANCING
  // -----------------------------------------------------------------

  calcularCreditoLance(params: Record<string, any>): FinanciamentoResult {
    const valor = params.valor_lance ?? 0;
    const prazo = params.prazo_meses ?? 0;
    const taxa = (params.taxa_mensal_pct ?? 0) / 100;
    const metodo = (params.metodo ?? 'price').toLowerCase();
    const carencia = params.carencia ?? 0;
    const tac = params.tac ?? 0;
    const avalGarantia = params.avaliacao_garantia ?? 0;
    const comissao = params.comissao ?? 0;
    const calcIof = params.calcular_iof !== undefined ? params.calcular_iof : true;
    const antecipacao = params.antecipacao_mes ?? 0;

    const custosAdd: CustoAcessorio[] = [];
    if (tac > 0)
      custosAdd.push({ descricao: 'TAC', valor: tac, momento: 0 });
    if (avalGarantia > 0)
      custosAdd.push({ descricao: 'Avaliacao Garantia', valor: avalGarantia, momento: 0 });
    if (comissao > 0)
      custosAdd.push({ descricao: 'Comissao', valor: comissao, momento: 0 });

    const finParams = {
      valor,
      prazo_meses: prazo,
      taxa_mensal_pct: taxa * 100,
      metodo,
      carencia,
      calcular_iof: calcIof,
      custos_adicionais: custosAdd,
    };

    const resultado = this.calcularFinanciamento(finParams);

    if (antecipacao > 0 && antecipacao < prazo) {
      let saldoNoMes = 0;
      for (const p of resultado.parcelas) {
        if (p.mes === antecipacao) {
          saldoNoMes = p.saldo;
          break;
        }
      }

      const cfNovo = [...resultado.cashflow];
      for (let i = antecipacao + 1; i < cfNovo.length; i++) {
        cfNovo[i] = 0;
      }
      if (antecipacao < cfNovo.length) {
        cfNovo[antecipacao] -= saldoNoMes;
      }

      resultado.cashflow_antecipado = cfNovo;
      resultado.valor_antecipacao = saldoNoMes;
      resultado.mes_antecipacao = antecipacao;
    }

    resultado.custos_iniciais = tac + avalGarantia + comissao;

    return resultado;
  }

  // -----------------------------------------------------------------
  // 5. COMBINED COST
  // -----------------------------------------------------------------

  calcularCustoCombinado(
    fluxoConsorcio: FluxoResult,
    fluxoLance: FinanciamentoResult
  ): {
    cashflow_combinado: number[];
    total_pago_consorcio: number;
    total_pago_lance: number;
    total_pago_combinado: number;
    tir_mensal_combinado: number;
    tir_anual_combinado: number;
    cet_anual_combinado: number;
  } {
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
      cashflow_combinado: cfCombinado,
      total_pago_consorcio: totalCons,
      total_pago_lance: totalLance,
      total_pago_combinado: totalCons + totalLance,
      tir_mensal_combinado: tirM,
      tir_anual_combinado: tirA,
      cet_anual_combinado: tirA,
    };
  }

  // -----------------------------------------------------------------
  // 6. CONSORTIUM VS FINANCING COMPARISON
  // -----------------------------------------------------------------

  compararConsorcioFinanciamento(
    paramsCons: Record<string, any>,
    paramsFin: Record<string, any>
  ): Record<string, any> {
    let tmaM = paramsCons.tma ?? 0.01;
    if (tmaM > 1) tmaM = tmaM / 100;
    const almA = (paramsCons.alm_anual ?? 12.0) / 100;
    const almM = monthlyFromAnnual(almA);

    const fluxoC = this.calcularFluxoCompleto(paramsCons);
    const cfC = fluxoC.cashflow;
    const vplCAlm = npv(almM, cfC);
    const vplCTma = npv(tmaM, cfC);
    const tirC = irr(cfC, 0.005);

    const fluxoF = this.calcularFinanciamento(paramsFin);
    const cfF = fluxoF.cashflow;
    const vplFAlm = npv(almM, cfF);
    const vplFTma = npv(tmaM, cfF);
    const tirF = irr(cfF, 0.008);

    const totalCons = fluxoC.total_pago ?? fluxoC.totais?.total_pago ?? 0;
    const totalFin = fluxoF.total_pago;

    const cartaLiq = fluxoC.carta_liquida ?? fluxoC.totais?.carta_liquida ?? 0;
    const valorFin = fluxoF.valor;

    const razaoC = cartaLiq > 0 ? totalCons / cartaLiq : 0;
    const razaoF = valorFin > 0 ? totalFin / valorFin : 0;

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

  // -----------------------------------------------------------------
  // 7. OPERATION SALE
  // -----------------------------------------------------------------

  calcularVendaOperacao(
    fluxoResult: FluxoResult,
    paramsVenda: Record<string, any>
  ): Record<string, any> {
    const momento = paramsVenda.momento_venda ?? 0;
    const valorVenda = paramsVenda.valor_venda ?? 0;
    const tma = paramsVenda.tma ?? 0.01;

    const cfOriginal = fluxoResult.cashflow ?? [];

    const cfVendedor: number[] = [];
    let totalGasto = 0.0;

    for (let i = 0; i < Math.min(momento + 1, cfOriginal.length); i++) {
      cfVendedor.push(cfOriginal[i]);
      if (cfOriginal[i] < 0) {
        totalGasto += Math.abs(cfOriginal[i]);
      }
    }

    if (cfVendedor.length > 0) {
      cfVendedor[cfVendedor.length - 1] += valorVenda;
    } else {
      cfVendedor.push(valorVenda);
    }

    const vplVendedor = npv(tma, cfVendedor);
    const tirVendedor = irr(cfVendedor, 0.01);

    const ganhoNominal = valorVenda - totalGasto;
    const ganhoPct = totalGasto > 0 ? (ganhoNominal / totalGasto) * 100 : 0;

    const prazoMedio = momento / 2;
    const ganhoMensal = momento > 0 ? ganhoNominal / momento : 0;
    const margemMensal = totalGasto > 0 ? (ganhoMensal / totalGasto) * 100 : 0;

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

  // -----------------------------------------------------------------
  // 8. EQUIVALENT CREDIT (GoalSeek)
  // -----------------------------------------------------------------

  calcularCreditoEquivalente(params: Record<string, any>): number {
    const creditoOriginal = params.valor_credito ?? 0;

    const custoTotalFunc = (creditoTeste: number): number => {
      const p = { ...params, valor_credito: creditoTeste };
      const fluxo = this.calcularFluxoCompleto(p);
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

  // -----------------------------------------------------------------
  // 9. GENERATE INSTALLMENTS
  // -----------------------------------------------------------------

  gerarParcelas(fluxoResult: FluxoResult): ParcelaInfo[] {
    const fluxoMensal = fluxoResult.fluxo ?? fluxoResult.fluxo_mensal ?? [];
    const parcelas: ParcelaInfo[] = [];

    for (const f of fluxoMensal) {
      if ((f.mes ?? 0) === 0) continue;

      parcelas.push({
        mes: f.mes,
        fundo_comum: Math.abs(
          (f.amortizacao ?? 0) - (f.lance_embutido ?? 0) - (f.lance_livre ?? 0)
        ),
        taxa_adm: f.valor_parcela_ta ?? 0,
        fundo_reserva: f.valor_fundo_reserva ?? 0,
        parcela_base: f.valor_parcela ?? 0,
        reajuste: f.fator_reajuste ?? 1.0,
        parcela_reajustada: f.parcela_apos_reajuste ?? 0,
        seguro: f.seguro_vida ?? 0,
        parcela_total: f.parcela_com_seguro ?? 0,
        outros_custos: f.outros_custos ?? 0,
        desembolso_total: (f.parcela_com_seguro ?? 0) + (f.outros_custos ?? 0),
        saldo_devedor: f.saldo_devedor_reajustado ?? f.saldo_devedor ?? 0,
        lance_embutido: Math.abs(f.lance_embutido ?? 0),
        lance_livre: Math.abs(f.lance_livre ?? 0),
        credito_recebido: f.credito_recebido ?? 0,
      });
    }

    return parcelas;
  }

  // -----------------------------------------------------------------
  // 10. CLIENT SUMMARY
  // -----------------------------------------------------------------

  gerarResumoCliente(
    inputParams: Record<string, any>,
    fluxoResult: FluxoResult,
    vplResult?: VPLResult | null
  ): ResumoCliente {
    const params = resolveParams(inputParams);
    const vpl = vplResult ?? this.calcularVPLHD(params, fluxoResult);

    const totais = fluxoResult.totais;
    const metricas = fluxoResult.metricas;
    const credito = params.valor_credito ?? 0;
    const prazo = params.prazo_meses ?? 0;
    const contemp = params.momento_contemplacao ?? 0;

    const lanceEmb = totais.lance_embutido_valor ?? 0;
    const lanceLivre = totais.lance_livre_valor ?? 0;
    const cartaLiq = totais.carta_liquida ?? credito;
    const totalPago = totais.total_pago ?? 0;

    const fluxoMensal = fluxoResult.fluxo ?? fluxoResult.fluxo_mensal ?? [];
    const parcelasVals = fluxoMensal
      .filter((f) => (f.mes ?? 0) > 0)
      .map((f) => f.parcela_com_seguro ?? f.parcela_apos_reajuste ?? 0);

    const primeiraParcela = parcelasVals.length > 0 ? parcelasVals[0] : 0;
    const ultimaParcela =
      parcelasVals.length > 0 ? parcelasVals[parcelasVals.length - 1] : 0;

    return {
      valor_credito: credito,
      prazo_meses: prazo,
      momento_contemplacao: contemp,
      carta_liquida: cartaLiq,
      lance_embutido: lanceEmb,
      lance_embutido_pct: credito > 0 ? (lanceEmb / credito) * 100 : 0,
      lance_livre: lanceLivre,
      lance_livre_pct: credito > 0 ? (lanceLivre / credito) * 100 : 0,
      lance_total: lanceEmb + lanceLivre,
      lance_total_pct: credito > 0 ? ((lanceEmb + lanceLivre) / credito) * 100 : 0,
      primeira_parcela: primeiraParcela,
      ultima_parcela: ultimaParcela,
      parcela_media: metricas.parcela_media ?? 0,
      parcela_maxima: metricas.parcela_maxima ?? 0,
      parcela_minima: metricas.parcela_minima ?? 0,
      total_pago: totalPago,
      total_fundo_comum: totais.total_fundo_comum ?? 0,
      total_taxa_adm: totais.total_taxa_adm ?? 0,
      total_fundo_reserva: totais.total_fundo_reserva ?? 0,
      total_seguro: totais.total_seguro ?? 0,
      total_custos_acessorios: totais.total_custos_acessorios ?? 0,
      custo_total_pct: metricas.custo_total_pct ?? 0,
      taxa_adm_pct: params.taxa_adm_pct ?? 0,
      fundo_reserva_pct: params.fundo_reserva_pct ?? 0,
      seguro_pct: params.seguro_vida_pct ?? 0,
      tir_mensal: vpl.tir_mensal ?? metricas.tir_mensal ?? 0,
      tir_anual: vpl.tir_anual ?? metricas.tir_anual ?? 0,
      cet_anual: vpl.cet_anual ?? metricas.cet_anual ?? 0,
      b0: vpl.b0 ?? 0,
      h0: vpl.h0 ?? 0,
      d0: vpl.d0 ?? 0,
      pv_pos_t: vpl.pv_pos_t ?? 0,
      delta_vpl: vpl.delta_vpl ?? 0,
      cria_valor: vpl.cria_valor ?? false,
      vpl_total: vpl.vpl_total ?? 0,
      break_even_lance: vpl.break_even_lance ?? 0,
      reajuste_pre_pct: params.reajuste_pre_pct ?? 0,
      reajuste_pos_pct: params.reajuste_pos_pct ?? 0,
      reajuste_pre_freq: params.reajuste_pre_freq ?? 'Anual',
      reajuste_pos_freq: params.reajuste_pos_freq ?? 'Anual',
    };
  }

  // -----------------------------------------------------------------
  // 11. MULTI-GROUP QUOTA CONSOLIDATION
  // -----------------------------------------------------------------

  consolidarCotas(
    grupos: Array<{ params: Record<string, any>; peso?: number }>
  ): Record<string, any> {
    if (!grupos || grupos.length === 0) return {};

    const fluxos: FluxoResult[] = [];
    const pesos: number[] = [];

    for (const g of grupos) {
      const p = g.params ?? g;
      const peso = g.peso ?? 1;
      const fl = this.calcularFluxoCompleto(p);
      fluxos.push(fl);
      pesos.push(peso);
    }

    const pesoTotal = pesos.reduce((a, b) => a + b, 0);
    const maxPrazo = Math.max(...fluxos.map((fl) => fl.cashflow.length));

    const cfConsolidado = new Array(maxPrazo).fill(0.0);
    for (let idx = 0; idx < fluxos.length; idx++) {
      const cf = fluxos[idx].cashflow;
      for (let i = 0; i < cf.length; i++) {
        cfConsolidado[i] += cf[i] * pesos[idx];
      }
    }

    let totalPago = 0;
    let totalCredito = 0;
    for (let idx = 0; idx < fluxos.length; idx++) {
      totalPago += (fluxos[idx].total_pago ?? fluxos[idx].totais?.total_pago ?? 0) * pesos[idx];
      totalCredito += (fluxos[idx].carta_liquida ?? fluxos[idx].totais?.carta_liquida ?? 0) * pesos[idx];
    }

    const tirM = irr(cfConsolidado, 0.005);
    const tirA = annualFromMonthly(tirM);

    return {
      cashflow_consolidado: cfConsolidado,
      total_pago: totalPago,
      total_credito: totalCredito,
      tir_mensal: tirM,
      tir_anual: tirA,
      num_cotas: grupos.length,
      peso_total: pesoTotal,
      fluxos_individuais: fluxos,
      custo_total_pct: totalCredito > 0 ? (totalPago / totalCredito) * 100 : 0,
    };
  }
}

// =============================================================================
// STANDALONE FUNCTIONS (compatibility with nasa_engine.py API)
// =============================================================================

let _engineDefault: NasaEngine | null = null;

function getEngine(): NasaEngine {
  if (_engineDefault === null) {
    _engineDefault = new NasaEngine();
  }
  return _engineDefault;
}

export function calcularFluxoConsorcio(params: Record<string, any>): Record<string, any> {
  const p: Record<string, any> = {};
  if ('valor_carta' in params) p.valor_credito = params.valor_carta;
  if ('valor_credito' in params) p.valor_credito = params.valor_credito;

  p.prazo_meses = params.prazo_meses ?? 200;

  if ('taxa_adm' in params) p.taxa_adm_pct = params.taxa_adm;
  if ('taxa_adm_pct' in params) p.taxa_adm_pct = params.taxa_adm_pct;

  if ('fundo_reserva' in params) p.fundo_reserva_pct = params.fundo_reserva;
  if ('fundo_reserva_pct' in params) p.fundo_reserva_pct = params.fundo_reserva_pct;

  p.momento_contemplacao = params.prazo_contemp ?? params.momento_contemplacao ?? 36;

  if ('seguro' in params) p.seguro_vida_pct = params.seguro;
  if ('seguro_vida_pct' in params) p.seguro_vida_pct = params.seguro_vida_pct;

  p.lance_embutido_pct = params.lance_embutido_pct ?? 0;
  p.lance_livre_pct = params.lance_livre_pct ?? 0;

  const redPct = params.parcela_red_pct ?? 100;
  const contemp = p.momento_contemplacao;
  const prazo = p.prazo_meses;

  if (redPct !== 100) {
    const fcPctF1 = redPct / 100.0;
    p.periodos = [
      { start: 1, end: contemp, fc_pct: fcPctF1, ta_pct: 1.0, fr_pct: 1.0 },
      { start: contemp + 1, end: prazo, fc_pct: 1.0, ta_pct: 1.0, fr_pct: 1.0 },
    ];
  }

  const corr = params.correcao_anual ?? 0;
  if (corr > 0) {
    p.reajuste_pre_pct = corr;
    p.reajuste_pos_pct = corr;
    p.reajuste_pre_freq = 'Anual';
    p.reajuste_pos_freq = 'Anual';
  }

  p.alm_anual = params.alm_anual ?? 12.0;
  p.hurdle_anual = params.hurdle_anual ?? 12.0;

  const engine = getEngine();
  const resultado = engine.calcularFluxoCompleto(p);

  const fluxoCompat: Record<string, any>[] = [];
  for (const f of resultado.fluxo) {
    fluxoCompat.push({
      mes: f.mes,
      parcela: f.parcela_com_seguro ?? f.parcela_apos_reajuste ?? 0,
      fundo_comum: Math.abs(f.amortizacao ?? 0),
      taxa_adm: f.valor_parcela_ta ?? 0,
      fundo_reserva: f.valor_fundo_reserva ?? 0,
      seguro: f.seguro_vida ?? 0,
      lance: Math.abs(f.lance_livre ?? 0),
      credito: f.credito_recebido ?? 0,
      fluxo_liquido: f.fluxo_caixa ?? 0,
      fator_correcao: f.fator_reajuste ?? 1.0,
    });
  }

  const result: Record<string, any> = { ...resultado };
  result.fluxo_mensal = fluxoCompat;
  result.cashflow_consorcio = resultado.cashflow;
  result.parcela_f1_base = fluxoCompat.length > 1 ? fluxoCompat[1].parcela : 0;
  result.parcela_f2_base = fluxoCompat.length > contemp + 1 ? fluxoCompat[contemp + 1].parcela : 0;
  result.meses_restantes = prazo - contemp;

  return result;
}

export function calcularVplHd(
  params: Record<string, any>,
  fluxoResult: FluxoResult
): VPLResult {
  const engine = getEngine();
  const p = { ...params };
  if ('valor_carta' in p && !('valor_credito' in p)) p.valor_credito = p.valor_carta;
  if ('prazo_contemp' in p && !('momento_contemplacao' in p)) p.momento_contemplacao = p.prazo_contemp;
  return engine.calcularVPLHD(p, fluxoResult);
}

export function calcularFinanciamentoStandalone(
  valor: number,
  prazoMeses: number,
  taxaMensalPct: number,
  metodo: string = 'price'
): FinanciamentoResult {
  const engine = getEngine();
  return engine.calcularFinanciamento({
    valor,
    prazo_meses: prazoMeses,
    taxa_mensal_pct: taxaMensalPct,
    metodo,
    calcular_iof: false,
  });
}

export function compararConsorcioFinanciamentoStandalone(
  paramsConsorcio: Record<string, any>,
  paramsFinanc: Record<string, any>
): Record<string, any> {
  const engine = getEngine();
  const pf: Record<string, any> = {};
  if ('valor' in paramsFinanc) pf.valor = paramsFinanc.valor;
  pf.prazo_meses = paramsFinanc.prazo_meses ?? 0;
  pf.taxa_mensal_pct = paramsFinanc.taxa_mensal_pct ?? 0;
  pf.metodo = paramsFinanc.metodo ?? 'price';
  pf.calcular_iof = false;
  return engine.compararConsorcioFinanciamento(paramsConsorcio, pf);
}

import { describe, it, expect } from 'vitest';
import { NasaEngine, DEFAULT_PARAMS } from '../../src/engine/nasa-engine';
import { npv, monthlyFromAnnual } from '../../src/engine/irr-npv';

// =============================================================================
// Basic consortium flow
// =============================================================================

describe('NasaEngine - calcularFluxoCompleto', () => {
  const engine = new NasaEngine();

  it('basic flow: R$1,000,000, 120 months, 18% admin, 2% reserve, contemplation at 60', () => {
    const params = {
      valor_credito: 1000000,
      prazo_meses: 120,
      taxa_adm_pct: 18,
      fundo_reserva_pct: 2,
      momento_contemplacao: 60,
      periodos: [
        { start: 1, end: 120, fc_pct: 1.0, ta_pct: 100.0, fr_pct: 100.0 },
      ],
      lance_embutido_pct: 0,
      lance_livre_pct: 0,
      lance_embutido_valor: 0,
      lance_livre_valor: 0,
      reajuste_pre_pct: 0,
      reajuste_pos_pct: 0,
      reajuste_pre_freq: 'Anual',
      reajuste_pos_freq: 'Anual',
      seguro_vida_pct: 0,
      seguro_vida_inicio: 1,
      antecipacao_ta_pct: 0,
      antecipacao_ta_parcelas: 1,
      taxa_vp_credito: 0,
      tma: 0.01,
      alm_anual: 12.0,
      hurdle_anual: 12.0,
      custos_acessorios: [],
    };

    const result = engine.calcularFluxoCompleto(params);

    // total_pago should be more than the credit value (admin + reserve)
    expect(result.total_pago).toBeGreaterThan(1000000);
    // carta_liquida = valor_credito - lance_embutido = 1000000
    expect(result.carta_liquida).toBeCloseTo(1000000, 2);
    // fluxo should have prazo + 1 entries (month 0 to 120)
    expect(result.fluxo.length).toBe(121);
    // cashflow should also have prazo + 1 entries
    expect(result.cashflow.length).toBe(121);
    // Total paid should be greater than the credit value (includes fees)
    expect(result.total_pago).toBeGreaterThan(1000000);
  });

  it('with lance embutido 30%: carta_liquida = 70% of carta', () => {
    const params = {
      valor_credito: 1000000,
      prazo_meses: 120,
      taxa_adm_pct: 18,
      fundo_reserva_pct: 2,
      momento_contemplacao: 60,
      periodos: [
        { start: 1, end: 120, fc_pct: 1.0, ta_pct: 100.0, fr_pct: 100.0 },
      ],
      lance_embutido_pct: 30,
      lance_livre_pct: 0,
      lance_embutido_valor: 0,
      lance_livre_valor: 0,
      reajuste_pre_pct: 0,
      reajuste_pos_pct: 0,
      reajuste_pre_freq: 'Anual',
      reajuste_pos_freq: 'Anual',
      seguro_vida_pct: 0,
      seguro_vida_inicio: 1,
      antecipacao_ta_pct: 0,
      antecipacao_ta_parcelas: 1,
      taxa_vp_credito: 0,
      tma: 0.01,
      alm_anual: 12.0,
      hurdle_anual: 12.0,
      custos_acessorios: [],
    };

    const result = engine.calcularFluxoCompleto(params);

    expect(result.carta_liquida).toBeCloseTo(700000, 2);
    expect(result.lance_embutido_valor).toBeCloseTo(300000, 2);
  });

  it('with lance livre 20%: paid at contemplation month', () => {
    const params = {
      valor_credito: 1000000,
      prazo_meses: 120,
      taxa_adm_pct: 18,
      fundo_reserva_pct: 2,
      momento_contemplacao: 60,
      periodos: [
        { start: 1, end: 120, fc_pct: 1.0, ta_pct: 100.0, fr_pct: 100.0 },
      ],
      lance_embutido_pct: 0,
      lance_livre_pct: 20,
      lance_embutido_valor: 0,
      lance_livre_valor: 0,
      reajuste_pre_pct: 0,
      reajuste_pos_pct: 0,
      reajuste_pre_freq: 'Anual',
      reajuste_pos_freq: 'Anual',
      seguro_vida_pct: 0,
      seguro_vida_inicio: 1,
      antecipacao_ta_pct: 0,
      antecipacao_ta_parcelas: 1,
      taxa_vp_credito: 0,
      tma: 0.01,
      alm_anual: 12.0,
      hurdle_anual: 12.0,
      custos_acessorios: [],
    };

    const result = engine.calcularFluxoCompleto(params);

    expect(result.lance_livre_valor).toBeCloseTo(200000, 2);
    // Lance livre appears at contemplation month
    const contempRow = result.fluxo.find(r => r.mes === 60);
    expect(contempRow).toBeDefined();
    expect(contempRow!.lance_livre).toBeCloseTo(-200000, 2);
  });

  it('with both lances: verify total lances', () => {
    const params = {
      valor_credito: 1000000,
      prazo_meses: 120,
      taxa_adm_pct: 18,
      fundo_reserva_pct: 2,
      momento_contemplacao: 60,
      periodos: [
        { start: 1, end: 120, fc_pct: 1.0, ta_pct: 100.0, fr_pct: 100.0 },
      ],
      lance_embutido_pct: 25,
      lance_livre_pct: 15,
      lance_embutido_valor: 0,
      lance_livre_valor: 0,
      reajuste_pre_pct: 0,
      reajuste_pos_pct: 0,
      reajuste_pre_freq: 'Anual',
      reajuste_pos_freq: 'Anual',
      seguro_vida_pct: 0,
      seguro_vida_inicio: 1,
      antecipacao_ta_pct: 0,
      antecipacao_ta_parcelas: 1,
      taxa_vp_credito: 0,
      tma: 0.01,
      alm_anual: 12.0,
      hurdle_anual: 12.0,
      custos_acessorios: [],
    };

    const result = engine.calcularFluxoCompleto(params);

    expect(result.lance_embutido_valor).toBeCloseTo(250000, 2);
    expect(result.lance_livre_valor).toBeCloseTo(150000, 2);
    expect(result.carta_liquida).toBeCloseTo(750000, 2);
  });

  it('cashflow sums correctly: total outflows minus credit received', () => {
    const params = {
      valor_credito: 500000,
      prazo_meses: 200,
      taxa_adm_pct: 20,
      fundo_reserva_pct: 3,
      momento_contemplacao: 36,
      periodos: [
        { start: 1, end: 200, fc_pct: 1.0, ta_pct: 100.0, fr_pct: 100.0 },
      ],
      lance_embutido_pct: 0,
      lance_livre_pct: 0,
      lance_embutido_valor: 0,
      lance_livre_valor: 0,
      reajuste_pre_pct: 0,
      reajuste_pos_pct: 0,
      reajuste_pre_freq: 'Anual',
      reajuste_pos_freq: 'Anual',
      seguro_vida_pct: 0,
      seguro_vida_inicio: 1,
      antecipacao_ta_pct: 0,
      antecipacao_ta_parcelas: 1,
      taxa_vp_credito: 0,
      tma: 0.01,
      alm_anual: 12.0,
      hurdle_anual: 12.0,
      custos_acessorios: [],
    };

    const result = engine.calcularFluxoCompleto(params);

    // Cashflow length should be prazo + 1
    expect(result.cashflow.length).toBe(201);
    // Sum of cashflow should be negative (paid more than received since admin + reserve > 0)
    const totalCf = result.cashflow.reduce((a, b) => a + b, 0);
    expect(totalCf).toBeLessThan(0);
  });
});

// =============================================================================
// VPL HD
// =============================================================================

describe('NasaEngine - calcularVPLHD', () => {
  const engine = new NasaEngine();

  it('VPL HD with ALM=12%, Hurdle=12%', () => {
    const params = {
      valor_credito: 500000,
      prazo_meses: 200,
      taxa_adm_pct: 20,
      fundo_reserva_pct: 3,
      momento_contemplacao: 36,
      periodos: [
        { start: 1, end: 200, fc_pct: 1.0, ta_pct: 100.0, fr_pct: 100.0 },
      ],
      lance_embutido_pct: 0,
      lance_livre_pct: 0,
      lance_embutido_valor: 0,
      lance_livre_valor: 0,
      reajuste_pre_pct: 0,
      reajuste_pos_pct: 0,
      reajuste_pre_freq: 'Anual',
      reajuste_pos_freq: 'Anual',
      seguro_vida_pct: 0,
      seguro_vida_inicio: 1,
      antecipacao_ta_pct: 0,
      antecipacao_ta_parcelas: 1,
      taxa_vp_credito: 0,
      tma: 0.01,
      alm_anual: 12.0,
      hurdle_anual: 12.0,
      custos_acessorios: [],
    };

    const fluxo = engine.calcularFluxoCompleto(params);
    const vpl = engine.calcularVPLHD(params, fluxo);

    // b0, h0, d0 should be defined numbers
    expect(isFinite(vpl.b0)).toBe(true);
    expect(isFinite(vpl.h0)).toBe(true);
    expect(isFinite(vpl.d0)).toBe(true);
    expect(isFinite(vpl.delta_vpl)).toBe(true);

    // cria_valor is boolean
    expect(typeof vpl.cria_valor).toBe('boolean');

    // break_even_lance should be between 0 and 100
    expect(vpl.break_even_lance).toBeGreaterThanOrEqual(0);
    expect(vpl.break_even_lance).toBeLessThanOrEqual(100);

    // d0 = b0 - h0
    expect(vpl.d0).toBeCloseTo(vpl.b0 - vpl.h0, 2);

    // delta_vpl = d0 - pv_pos_t
    expect(vpl.delta_vpl).toBeCloseTo(vpl.d0 - vpl.pv_pos_t, 2);
  });
});

// =============================================================================
// Correction / Readjustment
// =============================================================================

describe('NasaEngine - with correction', () => {
  const engine = new NasaEngine();

  it('3% annual correction: parcelas grow over time', () => {
    const params = {
      valor_credito: 500000,
      prazo_meses: 60,
      taxa_adm_pct: 18,
      fundo_reserva_pct: 2,
      momento_contemplacao: 24,
      periodos: [
        { start: 1, end: 60, fc_pct: 1.0, ta_pct: 100.0, fr_pct: 100.0 },
      ],
      lance_embutido_pct: 0,
      lance_livre_pct: 0,
      lance_embutido_valor: 0,
      lance_livre_valor: 0,
      reajuste_pre_pct: 3,
      reajuste_pos_pct: 3,
      reajuste_pre_freq: 'Anual',
      reajuste_pos_freq: 'Anual',
      seguro_vida_pct: 0,
      seguro_vida_inicio: 1,
      antecipacao_ta_pct: 0,
      antecipacao_ta_parcelas: 1,
      taxa_vp_credito: 0,
      tma: 0.01,
      alm_anual: 12.0,
      hurdle_anual: 12.0,
      custos_acessorios: [],
    };

    const result = engine.calcularFluxoCompleto(params);

    // Month 12 should have fator_reajuste = 1.03 (first annual adjustment)
    const month12 = result.fluxo.find(r => r.mes === 12);
    expect(month12).toBeDefined();
    expect(month12!.fator_reajuste).toBeCloseTo(1.03, 4);

    // Month 24 should have fator_reajuste = 1.03^2 = 1.0609
    const month24 = result.fluxo.find(r => r.mes === 24);
    expect(month24).toBeDefined();
    expect(month24!.fator_reajuste).toBeCloseTo(1.0609, 4);

    // Parcela at month 13 should be higher than month 1 (post-readjustment)
    const month1 = result.fluxo.find(r => r.mes === 1);
    const month13 = result.fluxo.find(r => r.mes === 13);
    expect(month13!.parcela_apos_reajuste).toBeGreaterThan(month1!.parcela_apos_reajuste);
  });
});

// =============================================================================
// Default params flow
// =============================================================================

describe('NasaEngine - default params', () => {
  const engine = new NasaEngine();

  it('default params produce valid flow', () => {
    const result = engine.calcularFluxoCompleto(DEFAULT_PARAMS);

    expect(result.fluxo.length).toBe(201); // 0..200
    expect(result.total_pago).toBeGreaterThan(0);
    expect(result.carta_liquida).toBe(500000);
    expect(result.metricas.parcela_media).toBeGreaterThan(0);
    expect(result.metricas.parcela_maxima).toBeGreaterThanOrEqual(result.metricas.parcela_media);
  });
});

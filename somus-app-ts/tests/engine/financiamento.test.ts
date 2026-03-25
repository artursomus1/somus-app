import { describe, it, expect } from 'vitest';
import {
  calcularFinanciamento,
  calcularCreditoLance,
  calcularIOF,
} from '../../src/engine/financiamento';
import { pmt } from '../../src/engine/irr-npv';

// =============================================================================
// Price method
// =============================================================================

describe('calcularFinanciamento - Price', () => {
  it('R$100,000, 120 months, 1%: PMT matches and saldo ends near 0', () => {
    const result = calcularFinanciamento({
      valor: 100000,
      prazo_meses: 120,
      taxa_mensal_pct: 1,
      metodo: 'price',
      calcular_iof: false,
    });

    const expectedPmt = pmt(0.01, 120, 100000);
    expect(result.parcelas[0].parcela).toBeCloseTo(expectedPmt, 2);

    // Final saldo should be near 0
    const lastParcela = result.parcelas[result.parcelas.length - 1];
    expect(lastParcela.saldo).toBeCloseTo(0, 0);

    // total_juros = total_pago - valor
    expect(result.total_juros).toBeCloseTo(result.total_pago - result.valor, 0);
  });

  it('all parcelas should be equal in Price method', () => {
    const result = calcularFinanciamento({
      valor: 50000,
      prazo_meses: 24,
      taxa_mensal_pct: 0.8,
      metodo: 'price',
      calcular_iof: false,
    });

    const firstParcela = result.parcelas[0].parcela;
    for (const p of result.parcelas) {
      expect(p.parcela).toBeCloseTo(firstParcela, 2);
    }
  });

  it('total_pago > valor (interest makes it cost more)', () => {
    const result = calcularFinanciamento({
      valor: 200000,
      prazo_meses: 60,
      taxa_mensal_pct: 1.5,
      metodo: 'price',
      calcular_iof: false,
    });

    expect(result.total_pago).toBeGreaterThan(200000);
  });
});

// =============================================================================
// SAC method
// =============================================================================

describe('calcularFinanciamento - SAC', () => {
  it('amortization is constant', () => {
    const result = calcularFinanciamento({
      valor: 120000,
      prazo_meses: 12,
      taxa_mensal_pct: 1,
      metodo: 'sac',
      calcular_iof: false,
    });

    const expectedAmort = 120000 / 12;
    for (const p of result.parcelas) {
      expect(p.amortizacao).toBeCloseTo(expectedAmort, 2);
    }
  });

  it('parcelas decrease over time', () => {
    const result = calcularFinanciamento({
      valor: 100000,
      prazo_meses: 24,
      taxa_mensal_pct: 1,
      metodo: 'sac',
      calcular_iof: false,
    });

    for (let i = 1; i < result.parcelas.length; i++) {
      expect(result.parcelas[i].parcela).toBeLessThan(result.parcelas[i - 1].parcela);
    }
  });

  it('final saldo near 0', () => {
    const result = calcularFinanciamento({
      valor: 300000,
      prazo_meses: 120,
      taxa_mensal_pct: 0.75,
      metodo: 'sac',
      calcular_iof: false,
    });

    const lastParcela = result.parcelas[result.parcelas.length - 1];
    expect(lastParcela.saldo).toBeCloseTo(0, 0);
  });
});

// =============================================================================
// IOF
// =============================================================================

describe('calcularIOF', () => {
  it('IOF is positive for non-zero values', () => {
    const iof = calcularIOF(100000, 12);
    expect(iof).toBeGreaterThan(0);
  });

  it('IOF increases with prazo', () => {
    const iof12 = calcularIOF(100000, 12);
    const iof60 = calcularIOF(100000, 60);
    expect(iof60).toBeGreaterThan(iof12);
  });

  it('IOF is 0 for zero valor', () => {
    const iof = calcularIOF(0, 12);
    expect(iof).toBe(0);
  });
});

describe('calcularFinanciamento with IOF', () => {
  it('IOF is calculated and included in custo_efetivo_total', () => {
    const result = calcularFinanciamento({
      valor: 100000,
      prazo_meses: 60,
      taxa_mensal_pct: 1,
      metodo: 'price',
      calcular_iof: true,
    });

    expect(result.iof).toBeGreaterThan(0);
    expect(result.custo_efetivo_total).toBeGreaterThan(result.total_pago);
  });
});

// =============================================================================
// Grace period (carencia)
// =============================================================================

describe('calcularFinanciamento with carencia', () => {
  it('during carencia, only interest is paid (no amortization)', () => {
    const result = calcularFinanciamento({
      valor: 100000,
      prazo_meses: 24,
      taxa_mensal_pct: 1,
      metodo: 'price',
      carencia: 6,
      calcular_iof: false,
    });

    // First 6 months: amortization should be 0
    for (let i = 0; i < 6; i++) {
      expect(result.parcelas[i].amortizacao).toBe(0);
      expect(result.parcelas[i].parcela).toBeCloseTo(result.parcelas[i].juros, 2);
    }

    // After carencia: amortization should start
    expect(result.parcelas[6].amortizacao).toBeGreaterThan(0);
  });
});

// =============================================================================
// Credito Lance
// =============================================================================

describe('calcularCreditoLance', () => {
  it('basic lance financing with TAC and costs', () => {
    const result = calcularCreditoLance({
      valor: 50000,
      prazo: 60,
      taxa: 1.2,
      metodo: 'price',
      tac: 1500,
      avaliacaoGarantia: 800,
      comissao: 500,
    });

    expect(result.custos_iniciais).toBe(2800);
    expect(result.total_pago).toBeGreaterThan(50000);
  });

  it('lance financing with early payoff', () => {
    const result = calcularCreditoLance({
      valor: 80000,
      prazo: 120,
      taxa: 0.9,
      metodo: 'price',
      pagamentoAntecipado: { mes: 36, valor: 0 },
    });

    expect(result.mes_antecipacao).toBe(36);
    expect(result.valor_antecipacao).toBeDefined();
    // Cashflow antecipado should zero out after month 36
    if (result.cashflow_antecipado) {
      for (let i = 37; i < result.cashflow_antecipado.length; i++) {
        expect(result.cashflow_antecipado[i]).toBe(0);
      }
    }
  });
});

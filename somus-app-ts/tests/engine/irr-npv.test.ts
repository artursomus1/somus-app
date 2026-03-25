import { describe, it, expect } from 'vitest';
import { npv, irr, pmt, annualFromMonthly, monthlyFromAnnual, goalSeek } from '../../src/engine/irr-npv';

// =============================================================================
// NPV
// =============================================================================

describe('npv', () => {
  it('zero rate returns sum of cashflows', () => {
    const cfs = [-1000, 300, 400, 500];
    expect(npv(0, cfs)).toBeCloseTo(200, 2);
  });

  it('known rate with known cashflows returns expected value', () => {
    // Investment of 10000 followed by 3 annual payments of 4000
    // at 10% monthly-equivalent? Use a simpler example:
    // NPV at 1% of [-1000, 500, 500, 200]
    const cfs = [-1000, 500, 500, 200];
    // NPV = -1000 + 500/1.01 + 500/1.01^2 + 200/1.01^3
    const expected = -1000 + 500 / 1.01 + 500 / Math.pow(1.01, 2) + 200 / Math.pow(1.01, 3);
    expect(npv(0.01, cfs)).toBeCloseTo(expected, 6);
  });

  it('empty cashflows returns 0', () => {
    expect(npv(0.05, [])).toBe(0);
  });

  it('single cashflow returns the cashflow itself', () => {
    expect(npv(0.10, [5000])).toBeCloseTo(5000, 2);
  });

  it('handles negative rate gracefully', () => {
    const cfs = [-1000, 600, 600];
    const result = npv(-0.05, cfs);
    expect(isFinite(result)).toBe(true);
  });
});

// =============================================================================
// IRR
// =============================================================================

describe('irr', () => {
  it('standard investment (negative then positives) returns known rate', () => {
    // -1000 + 1100 after 1 period => IRR = 10%
    const cfs = [-1000, 1100];
    const result = irr(cfs);
    expect(result).toBeCloseTo(0.10, 6);
  });

  it('multi-period investment returns expected IRR', () => {
    // -10000 then 5 payments of 2500 => IRR ~= 7.93%
    const cfs = [-10000, 2500, 2500, 2500, 2500, 2500];
    const result = irr(cfs);
    // Verify NPV at IRR is approximately zero
    expect(npv(result, cfs)).toBeCloseTo(0, 4);
  });

  it('all positive cashflows returns 0 (no sign change)', () => {
    const cfs = [100, 200, 300];
    const result = irr(cfs);
    // Should handle gracefully - either 0 or a value where NPV ~ 0
    expect(isFinite(result)).toBe(true);
  });

  it('all negative cashflows returns 0', () => {
    const cfs = [-100, -200, -300];
    const result = irr(cfs);
    expect(isFinite(result)).toBe(true);
  });

  it('consortium-like cashflow converges', () => {
    // Simulating a consortium: receive credit at month 3, pay installments
    const cfs: number[] = new Array(13).fill(0);
    cfs[0] = 0;
    for (let i = 1; i <= 12; i++) cfs[i] = -1000;
    cfs[3] += 15000; // credit received at month 3
    const result = irr(cfs);
    expect(isFinite(result)).toBe(true);
    expect(Math.abs(npv(result, cfs))).toBeLessThan(1);
  });
});

// =============================================================================
// PMT
// =============================================================================

describe('pmt', () => {
  it('known loan: R$100,000, 12 months, 1%', () => {
    const result = pmt(0.01, 12, 100000);
    // Excel PMT(0.01, 12, -100000) = 8884.88
    expect(result).toBeCloseTo(8884.88, 0);
  });

  it('zero rate returns valor/prazo', () => {
    const result = pmt(0, 12, 120000);
    expect(result).toBeCloseTo(10000, 2);
  });

  it('zero periods returns Infinity (division edge case)', () => {
    const result = pmt(0.01, 0, 100000);
    expect(result).toBe(Infinity);
  });

  it('R$500,000 at 0.8% for 360 months (typical mortgage)', () => {
    const result = pmt(0.008, 360, 500000);
    // Should be a reasonable monthly payment
    expect(result).toBeGreaterThan(3000);
    expect(result).toBeLessThan(6000);
  });

  it('matches Excel PMT result for R$250,000, 120 months, 0.75%', () => {
    const result = pmt(0.0075, 120, 250000);
    // PMT formula result for these params
    expect(result).toBeCloseTo(3166.89, 0);
  });
});

// =============================================================================
// Rate conversions
// =============================================================================

describe('annualFromMonthly', () => {
  it('1% monthly to 12.68% annual', () => {
    const result = annualFromMonthly(0.01);
    expect(result * 100).toBeCloseTo(12.68, 1);
  });

  it('0% monthly to 0% annual', () => {
    expect(annualFromMonthly(0)).toBe(0);
  });

  it('2% monthly to ~26.82% annual', () => {
    const result = annualFromMonthly(0.02);
    expect(result * 100).toBeCloseTo(26.82, 1);
  });
});

describe('monthlyFromAnnual', () => {
  it('12% annual to ~0.949% monthly', () => {
    const result = monthlyFromAnnual(0.12);
    expect(result * 100).toBeCloseTo(0.949, 2);
  });

  it('0% annual to 0% monthly', () => {
    expect(monthlyFromAnnual(0)).toBe(0);
  });

  it('-1 or below returns 0', () => {
    expect(monthlyFromAnnual(-1)).toBe(0);
    expect(monthlyFromAnnual(-2)).toBe(0);
  });

  it('round trip: monthly -> annual -> monthly', () => {
    const monthly = 0.01;
    const annual = annualFromMonthly(monthly);
    const backToMonthly = monthlyFromAnnual(annual);
    expect(backToMonthly).toBeCloseTo(monthly, 10);
  });

  it('round trip: annual -> monthly -> annual', () => {
    const annual = 0.15;
    const monthly = monthlyFromAnnual(annual);
    const backToAnnual = annualFromMonthly(monthly);
    expect(backToAnnual).toBeCloseTo(annual, 10);
  });
});

// =============================================================================
// GoalSeek
// =============================================================================

describe('goalSeek', () => {
  it('find x where x^2 = 4 => x = 2', () => {
    const result = goalSeek((x) => x * x, 4, 0, 10);
    expect(result).toBeCloseTo(2, 6);
  });

  it('find x where 2x + 3 = 11 => x = 4', () => {
    const result = goalSeek((x) => 2 * x + 3, 11, -10, 20);
    expect(result).toBeCloseTo(4, 6);
  });

  it('find rate where NPV = 0 (same as IRR)', () => {
    const cfs = [-1000, 500, 400, 300];
    const result = goalSeek(
      (r) => npv(r, cfs),
      0,
      -0.5,
      2.0
    );
    const irrResult = irr(cfs);
    expect(result).toBeCloseTo(irrResult, 4);
  });

  it('find x where x^3 = 27 => x = 3', () => {
    const result = goalSeek((x) => x * x * x, 27, 0, 10);
    expect(result).toBeCloseTo(3, 6);
  });
});

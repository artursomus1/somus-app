import { describe, it, expect } from 'vitest';
import {
  fmtCurrency,
  fmtCurrencyShort,
  fmtPct,
  fmtPctAm,
  fmtPctAa,
  fmtMes,
  fmtMesAno,
  fmtDate,
  parseBRNumber,
  fmtMeses,
  fmtHorizonte,
} from '../../src/utils/format';

// =============================================================================
// fmtCurrency
// =============================================================================

describe('fmtCurrency', () => {
  it('formats positive value with R$ prefix', () => {
    const result = fmtCurrency(1234567.89);
    // Should contain "R$" and the formatted number
    expect(result).toContain('R$');
    expect(result).toContain('1.234.567');
    expect(result).toContain('89');
  });

  it('formats zero', () => {
    const result = fmtCurrency(0);
    expect(result).toContain('R$');
    expect(result).toContain('0');
  });

  it('formats negative value', () => {
    const result = fmtCurrency(-500);
    expect(result).toContain('R$');
    expect(result).toContain('500');
  });

  it('formats small values', () => {
    const result = fmtCurrency(42.5);
    expect(result).toContain('R$');
    expect(result).toContain('42');
    expect(result).toContain('50');
  });
});

// =============================================================================
// fmtCurrencyShort
// =============================================================================

describe('fmtCurrencyShort', () => {
  it('millions are abbreviated with M', () => {
    const result = fmtCurrencyShort(1234567);
    expect(result).toContain('R$');
    expect(result).toContain('M');
  });

  it('hundreds of thousands are abbreviated with K', () => {
    const result = fmtCurrencyShort(500000);
    expect(result).toContain('R$');
    expect(result).toContain('K');
  });

  it('small values use full format', () => {
    const result = fmtCurrencyShort(1500);
    expect(result).toContain('R$');
    // Should not contain K or M
    expect(result).not.toContain('K');
    expect(result).not.toContain('M');
  });

  it('billions are abbreviated with B', () => {
    const result = fmtCurrencyShort(2500000000);
    expect(result).toContain('R$');
    expect(result).toContain('B');
  });
});

// =============================================================================
// fmtPct
// =============================================================================

describe('fmtPct', () => {
  it('formats percentage with 2 decimals by default', () => {
    const result = fmtPct(12.5);
    expect(result).toContain('12');
    expect(result).toContain('50');
    expect(result).toContain('%');
  });

  it('formats with custom decimals', () => {
    const result = fmtPct(12.5, 1);
    expect(result).toContain('12');
    expect(result).toContain('5');
    expect(result).toContain('%');
  });

  it('formats zero', () => {
    const result = fmtPct(0);
    expect(result).toContain('0');
    expect(result).toContain('%');
  });
});

// =============================================================================
// fmtPctAm / fmtPctAa
// =============================================================================

describe('fmtPctAm', () => {
  it('adds a.m. suffix', () => {
    expect(fmtPctAm(1.5)).toContain('a.m.');
  });
});

describe('fmtPctAa', () => {
  it('adds a.a. suffix', () => {
    expect(fmtPctAa(19.56)).toContain('a.a.');
  });
});

// =============================================================================
// fmtMes
// =============================================================================

describe('fmtMes', () => {
  it('formats month ordinal', () => {
    expect(fmtMes(60)).toBe('60o mes');
  });
});

// =============================================================================
// fmtMesAno
// =============================================================================

describe('fmtMesAno', () => {
  it('formats date as month year in Portuguese', () => {
    const result = fmtMesAno(new Date(2026, 2, 24)); // March 2026
    expect(result).toBe('Marco 2026');
  });

  it('January', () => {
    expect(fmtMesAno(new Date(2025, 0, 1))).toBe('Janeiro 2025');
  });

  it('December', () => {
    expect(fmtMesAno(new Date(2024, 11, 31))).toBe('Dezembro 2024');
  });
});

// =============================================================================
// fmtDate
// =============================================================================

describe('fmtDate', () => {
  it('formats date as DD/MM/YYYY', () => {
    expect(fmtDate(new Date(2026, 2, 24))).toBe('24/03/2026');
  });

  it('pads single digits', () => {
    expect(fmtDate(new Date(2025, 0, 5))).toBe('05/01/2025');
  });
});

// =============================================================================
// parseBRNumber
// =============================================================================

describe('parseBRNumber', () => {
  it('parses Brazilian number format', () => {
    expect(parseBRNumber('1.234,56')).toBeCloseTo(1234.56, 2);
  });

  it('parses plain integer', () => {
    expect(parseBRNumber('1234')).toBe(1234);
  });

  it('parses with currency prefix (stripped)', () => {
    expect(parseBRNumber('R$ 1.234,56')).toBeCloseTo(1234.56, 2);
  });

  it('returns 0 for empty string', () => {
    expect(parseBRNumber('')).toBe(0);
  });

  it('returns 0 for non-numeric string', () => {
    expect(parseBRNumber('abc')).toBe(0);
  });

  it('handles negative numbers', () => {
    expect(parseBRNumber('-500,00')).toBeCloseTo(-500, 2);
  });
});

// =============================================================================
// fmtMeses
// =============================================================================

describe('fmtMeses', () => {
  it('singular', () => {
    expect(fmtMeses(1)).toBe('1 mes');
  });

  it('plural', () => {
    expect(fmtMeses(225)).toBe('225 meses');
  });
});

// =============================================================================
// fmtHorizonte
// =============================================================================

describe('fmtHorizonte', () => {
  it('225 months = 18 anos e 9 meses', () => {
    expect(fmtHorizonte(225)).toBe('18 anos e 9 meses');
  });

  it('12 months = 1 ano', () => {
    expect(fmtHorizonte(12)).toBe('1 ano');
  });

  it('5 months = 5 meses', () => {
    expect(fmtHorizonte(5)).toBe('5 meses');
  });

  it('24 months = 2 anos', () => {
    expect(fmtHorizonte(24)).toBe('2 anos');
  });

  it('13 months = 1 ano e 1 mes', () => {
    expect(fmtHorizonte(13)).toBe('1 ano e 1 mes');
  });

  it('1 month = 1 mes', () => {
    expect(fmtHorizonte(1)).toBe('1 mes');
  });

  it('25 months = 2 anos e 1 mes', () => {
    expect(fmtHorizonte(25)).toBe('2 anos e 1 mes');
  });
});

import { describe, it, expect } from 'vitest';
import {
  consolidarCotas,
  consolidarCotasDetalhado,
  consolidarCotasComQuantidade,
} from '../../src/engine/consolidador-cotas';

// =============================================================================
// Simple consolidation
// =============================================================================

describe('consolidarCotas', () => {
  it('single group returns same values', () => {
    const result = consolidarCotas([
      { grupo: 'G1', valor: 500000, prazo: 120, taxaAdm: 18, fundoReserva: 3 },
    ]);

    expect(result.totalCredito).toBe(500000);
    expect(result.prazoMedio).toBe(120);
    expect(result.taxaAdmMedia).toBe(18);
    expect(result.fundoReservaMedia).toBe(3);
  });

  it('two equal groups return same values', () => {
    const result = consolidarCotas([
      { grupo: 'G1', valor: 500000, prazo: 120, taxaAdm: 18, fundoReserva: 3 },
      { grupo: 'G2', valor: 500000, prazo: 120, taxaAdm: 18, fundoReserva: 3 },
    ]);

    expect(result.totalCredito).toBe(1000000);
    expect(result.prazoMedio).toBe(120);
    expect(result.taxaAdmMedia).toBe(18);
    expect(result.fundoReservaMedia).toBe(3);
  });

  it('weighted averages are correct', () => {
    // G1: 300k, prazo=120, taxa=20, fr=3
    // G2: 700k, prazo=200, taxa=15, fr=2
    // Weighted avg prazo = (300k*120 + 700k*200) / 1M = (36M + 140M)/1M = 176
    const result = consolidarCotas([
      { grupo: 'G1', valor: 300000, prazo: 120, taxaAdm: 20, fundoReserva: 3 },
      { grupo: 'G2', valor: 700000, prazo: 200, taxaAdm: 15, fundoReserva: 2 },
    ]);

    expect(result.totalCredito).toBe(1000000);
    expect(result.prazoMedio).toBeCloseTo(176, 2);
    expect(result.taxaAdmMedia).toBeCloseTo(16.5, 2); // (300k*20+700k*15)/1M
    expect(result.fundoReservaMedia).toBeCloseTo(2.3, 2); // (300k*3+700k*2)/1M
  });

  it('empty groups returns zeros', () => {
    const result = consolidarCotas([]);

    expect(result.totalCredito).toBe(0);
    expect(result.prazoMedio).toBe(0);
    expect(result.taxaAdmMedia).toBe(0);
    expect(result.fundoReservaMedia).toBe(0);
  });

  it('groups with zero valor returns zeros', () => {
    const result = consolidarCotas([
      { grupo: 'G1', valor: 0, prazo: 120, taxaAdm: 18, fundoReserva: 3 },
    ]);

    expect(result.totalCredito).toBe(0);
    expect(result.prazoMedio).toBe(0);
  });
});

// =============================================================================
// Detailed consolidation
// =============================================================================

describe('consolidarCotasDetalhado', () => {
  it('detailed consolidation includes all fields', () => {
    const result = consolidarCotasDetalhado([
      {
        grupo: 'G1',
        valor: 400000,
        prazo: 180,
        taxaAdm: 18,
        fundoReserva: 3,
        seguro: 0.03,
        lanceEmbutido: 25,
        lanceLivre: 10,
        momentoContemplacao: 36,
        reajusteAnual: 5,
      },
      {
        grupo: 'G2',
        valor: 600000,
        prazo: 200,
        taxaAdm: 15,
        fundoReserva: 2,
        seguro: 0.02,
        lanceEmbutido: 20,
        lanceLivre: 5,
        momentoContemplacao: 48,
        reajusteAnual: 3,
      },
    ]);

    expect(result.totalCredito).toBe(1000000);
    expect(result.numGrupos).toBe(2);
    expect(result.grupos.length).toBe(2);

    // Weighted averages
    expect(result.prazoMedio).toBeCloseTo((400000 * 180 + 600000 * 200) / 1000000, 2);
    expect(result.taxaAdmMedia).toBeCloseTo((400000 * 18 + 600000 * 15) / 1000000, 2);
    expect(result.seguroMedio).toBeCloseTo((400000 * 0.03 + 600000 * 0.02) / 1000000, 6);
    expect(result.contemplacaoMedia).toBeCloseTo((400000 * 36 + 600000 * 48) / 1000000, 2);
    expect(result.reajusteMedio).toBeCloseTo((400000 * 5 + 600000 * 3) / 1000000, 2);

    // lanceTotalPct = weighted average of (lanceEmbutido + lanceLivre)
    const lanceG1 = 25 + 10;
    const lanceG2 = 20 + 5;
    expect(result.lanceTotalPct).toBeCloseTo((400000 * lanceG1 + 600000 * lanceG2) / 1000000, 2);
  });

  it('empty detailed returns zeros', () => {
    const result = consolidarCotasDetalhado([]);
    expect(result.totalCredito).toBe(0);
    expect(result.numGrupos).toBe(0);
    expect(result.grupos.length).toBe(0);
  });
});

// =============================================================================
// Quantity-weighted consolidation
// =============================================================================

describe('consolidarCotasComQuantidade', () => {
  it('quantity multiplies the value correctly', () => {
    const result = consolidarCotasComQuantidade([
      { grupo: 'G1', valor: 100000, prazo: 120, taxaAdm: 18, fundoReserva: 3, quantidade: 3 },
    ]);

    expect(result.totalCredito).toBe(300000);
    expect(result.quantidadeTotal).toBe(3);
    expect(result.prazoMedio).toBe(120);
    expect(result.taxaAdmMedia).toBe(18);
  });

  it('multiple groups with quantities', () => {
    const result = consolidarCotasComQuantidade([
      { grupo: 'G1', valor: 100000, prazo: 120, taxaAdm: 20, fundoReserva: 3, quantidade: 2 },
      { grupo: 'G2', valor: 200000, prazo: 180, taxaAdm: 15, fundoReserva: 2, quantidade: 1 },
    ]);

    // Total: 2*100k + 1*200k = 400k
    expect(result.totalCredito).toBe(400000);
    expect(result.quantidadeTotal).toBe(3);

    // Weighted prazo: (200k*120 + 200k*180) / 400k = (24M + 36M)/400k = 150
    expect(result.prazoMedio).toBeCloseTo(150, 2);
  });

  it('empty returns zeros', () => {
    const result = consolidarCotasComQuantidade([]);
    expect(result.totalCredito).toBe(0);
    expect(result.quantidadeTotal).toBe(0);
  });
});

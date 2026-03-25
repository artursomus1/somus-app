import { describe, it, expect, beforeEach } from 'vitest';
import { ScenarioManager } from '../../src/engine/scenario-manager';

describe('ScenarioManager', () => {
  let manager: ScenarioManager;

  beforeEach(() => {
    manager = new ScenarioManager();
  });

  // ===========================================================================
  // Save / Load / Remove
  // ===========================================================================

  describe('save/load/remove', () => {
    it('save returns an index and load retrieves the scenario', () => {
      const idx = manager.save(
        'Cenario 1',
        { valor_credito: 500000, prazo_meses: 200 },
        { total_pago: 650000, tir_mensal: 0.005 }
      );

      expect(idx).toBe(1);
      expect(manager.count).toBe(1);

      const loaded = manager.load(idx);
      expect(loaded).not.toBeNull();
      expect(loaded!.name).toBe('Cenario 1');
      expect(loaded!.params.valor_credito).toBe(500000);
      expect(loaded!.results.total_pago).toBe(650000);
    });

    it('load returns null for non-existent index', () => {
      expect(manager.load(999)).toBeNull();
    });

    it('remove deletes the scenario', () => {
      const idx = manager.save('Test', {}, {});
      expect(manager.count).toBe(1);

      manager.remove(idx);
      expect(manager.count).toBe(0);
      expect(manager.load(idx)).toBeNull();
    });

    it('clear removes all scenarios', () => {
      manager.save('A', {}, {});
      manager.save('B', {}, {});
      manager.save('C', {}, {});
      expect(manager.count).toBe(3);

      manager.clear();
      expect(manager.count).toBe(0);
    });

    it('saved scenario is a deep copy (no mutation)', () => {
      const params = { valor_credito: 100000 };
      const results = { total_pago: 120000 };
      const idx = manager.save('Test', params, results);

      // Mutate original
      params.valor_credito = 999;
      results.total_pago = 999;

      const loaded = manager.load(idx)!;
      expect(loaded.params.valor_credito).toBe(100000);
      expect(loaded.results.total_pago).toBe(120000);
    });
  });

  // ===========================================================================
  // Max 10 scenarios
  // ===========================================================================

  describe('max 10 scenarios', () => {
    it('throws error when saving 11th scenario', () => {
      for (let i = 0; i < 10; i++) {
        manager.save(`Cenario ${i + 1}`, {}, {});
      }
      expect(manager.count).toBe(10);
      expect(manager.isFull).toBe(true);

      expect(() => manager.save('Cenario 11', {}, {})).toThrow(
        /Maximo de 10 cenarios/
      );
    });

    it('after removing one, can save again', () => {
      for (let i = 0; i < 10; i++) {
        manager.save(`Cenario ${i + 1}`, {}, {});
      }

      manager.remove(1);
      expect(manager.isFull).toBe(false);

      // Should not throw
      const idx = manager.save('New', {}, {});
      expect(idx).toBeGreaterThan(0);
      expect(manager.count).toBe(10);
    });
  });

  // ===========================================================================
  // Export / Import JSON
  // ===========================================================================

  describe('export/import JSON', () => {
    it('export produces valid JSON and import restores scenarios', () => {
      manager.save('A', { valor_credito: 100000 }, { total_pago: 120000 });
      manager.save('B', { valor_credito: 200000 }, { total_pago: 250000 });

      const json = manager.exportJSON();
      const parsed = JSON.parse(json);
      expect(parsed.version).toBe(1);
      expect(parsed.scenarios.length).toBe(2);

      // Create new manager and import
      const manager2 = new ScenarioManager();
      manager2.importJSON(json);

      expect(manager2.count).toBe(2);
      const scenarioA = manager2.load(1)!;
      expect(scenarioA.name).toBe('A');
      expect(scenarioA.params.valor_credito).toBe(100000);
    });

    it('import with invalid JSON throws error', () => {
      expect(() => manager.importJSON('{}')).toThrow(/invalido/);
    });

    it('import replaces existing scenarios', () => {
      manager.save('Old', {}, {});
      expect(manager.count).toBe(1);

      const json = JSON.stringify({
        version: 1,
        scenarios: [
          { index: 1, name: 'Imported', params: {}, results: {}, timestamp: Date.now() },
        ],
        nextId: 2,
      });

      manager.importJSON(json);
      expect(manager.count).toBe(1);
      expect(manager.load(1)!.name).toBe('Imported');
    });
  });

  // ===========================================================================
  // Compare
  // ===========================================================================

  describe('compare', () => {
    it('compare returns correct best VPL and TIR', () => {
      manager.save(
        'Low VPL',
        { valor_credito: 100000 },
        { delta_vpl: 10000, tir_anual: 0.08 }
      );
      manager.save(
        'High VPL',
        { valor_credito: 200000 },
        { delta_vpl: 50000, tir_anual: 0.15 }
      );
      manager.save(
        'Mid VPL High TIR',
        { valor_credito: 150000 },
        { delta_vpl: 30000, tir_anual: 0.20 }
      );

      const comparison = manager.compare([1, 2, 3]);

      expect(comparison.cenarios.length).toBe(3);
      expect(comparison.melhorVPL).toBe(2); // delta_vpl=50000
      expect(comparison.melhorTIR).toBe(3); // tir_anual=0.20
    });

    it('compare ignores non-existent indices', () => {
      manager.save('A', {}, { delta_vpl: 100, tir_anual: 0.1 });

      const comparison = manager.compare([1, 999]);
      expect(comparison.cenarios.length).toBe(1);
    });
  });

  // ===========================================================================
  // List
  // ===========================================================================

  describe('list', () => {
    it('returns scenarios sorted by index', () => {
      manager.save('C', {}, {});
      manager.save('A', {}, {});
      manager.save('B', {}, {});

      const list = manager.list();
      expect(list.length).toBe(3);
      expect(list[0].index).toBeLessThan(list[1].index);
      expect(list[1].index).toBeLessThan(list[2].index);
    });
  });
});

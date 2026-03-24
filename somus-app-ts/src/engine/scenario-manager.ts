/**
 * scenario-manager.ts - Scenario management
 * Port of nasa_engine_hd.py ScenarioManager class
 * Somus Capital - Mesa de Produtos
 *
 * Saves, loads, compares, and manages up to 10 simulation scenarios.
 */

import type { FluxoResult, VPLResult, NasaParams } from './nasa-engine';

// =============================================================================
// TYPES
// =============================================================================

export interface Scenario {
  index: number;
  name: string;
  params: Record<string, any>;
  results: Record<string, any>;
  vpl?: VPLResult;
  timestamp: number;
}

export interface ScenarioComparison {
  cenarios: Scenario[];
  melhorVPL: number;
  melhorTIR: number;
}

export interface ScenarioSummary {
  index: number;
  name: string;
  valor_credito: number;
  prazo_meses: number;
  total_pago: number;
  tir_mensal: number;
  tir_anual: number;
  delta_vpl: number;
  parcela_media: number;
}

// =============================================================================
// SCENARIO MANAGER
// =============================================================================

export class ScenarioManager {
  private scenarios: Map<number, Scenario> = new Map();
  private nextId: number = 1;
  private maxScenarios: number = 10;

  /**
   * Save a scenario. Returns the assigned index.
   *
   * @param nome - scenario name
   * @param params - simulation parameters
   * @param resultado - flow result
   * @param vpl - optional VPL result
   * @returns scenario index
   * @throws Error if max scenarios reached
   */
  save(
    nome: string,
    params: Record<string, any>,
    resultado: FluxoResult | Record<string, any>,
    vpl?: VPLResult
  ): number {
    if (this.scenarios.size >= this.maxScenarios) {
      throw new Error(
        `Maximo de ${this.maxScenarios} cenarios atingido. Limpe antes de salvar.`
      );
    }

    const idx = this.nextId;

    // Deep clone params and results to avoid mutation
    const scenario: Scenario = {
      index: idx,
      name: nome,
      params: JSON.parse(JSON.stringify(params)),
      results: JSON.parse(JSON.stringify(resultado)),
      vpl: vpl ? JSON.parse(JSON.stringify(vpl)) : undefined,
      timestamp: Date.now(),
    };

    this.scenarios.set(idx, scenario);
    this.nextId++;

    return idx;
  }

  /**
   * Load a scenario by index.
   *
   * @param index - scenario index
   * @returns deep copy of the scenario, or null if not found
   */
  load(index: number): Scenario | null {
    const sc = this.scenarios.get(index);
    if (!sc) return null;
    return JSON.parse(JSON.stringify(sc));
  }

  /**
   * Remove a scenario by index.
   *
   * @param index - scenario index to remove
   */
  remove(index: number): void {
    this.scenarios.delete(index);
  }

  /**
   * Remove all scenarios.
   */
  clear(): void {
    this.scenarios.clear();
    this.nextId = 1;
  }

  /**
   * List all saved scenarios (summary info).
   *
   * @returns array of scenarios sorted by index
   */
  list(): Scenario[] {
    const sorted = Array.from(this.scenarios.entries())
      .sort(([a], [b]) => a - b)
      .map(([, sc]) => JSON.parse(JSON.stringify(sc)));
    return sorted;
  }

  /**
   * Compare selected scenarios by index.
   * Returns the selected scenarios along with best VPL and IRR indices.
   *
   * @param indices - array of scenario indices to compare
   * @returns comparison result with best metrics
   */
  compare(indices: number[]): ScenarioComparison {
    const cenarios: Scenario[] = [];
    let melhorVPL = -1;
    let melhorVPLValue = -Infinity;
    let melhorTIR = -1;
    let melhorTIRValue = -Infinity;

    for (const idx of indices) {
      const sc = this.scenarios.get(idx);
      if (!sc) continue;

      cenarios.push(JSON.parse(JSON.stringify(sc)));

      // Extract delta_vpl
      const deltaVpl =
        sc.vpl?.delta_vpl ??
        (sc.results as any)?.delta_vpl ??
        0;
      if (deltaVpl > melhorVPLValue) {
        melhorVPLValue = deltaVpl;
        melhorVPL = idx;
      }

      // Extract TIR
      const tirAnual =
        sc.vpl?.tir_anual ??
        (sc.results as any)?.metricas?.tir_anual ??
        (sc.results as any)?.tir_anual ??
        0;
      if (tirAnual > melhorTIRValue) {
        melhorTIRValue = tirAnual;
        melhorTIR = idx;
      }
    }

    return {
      cenarios,
      melhorVPL,
      melhorTIR,
    };
  }

  /**
   * Get detailed summary for comparison.
   * Returns structured data for each scenario index.
   */
  getSummaries(indices: number[]): ScenarioSummary[] {
    const summaries: ScenarioSummary[] = [];

    for (const idx of indices) {
      const sc = this.scenarios.get(idx);
      if (!sc) continue;

      const results = sc.results as any;
      summaries.push({
        index: idx,
        name: sc.name,
        valor_credito: sc.params.valor_credito ?? 0,
        prazo_meses: sc.params.prazo_meses ?? 0,
        total_pago:
          results.total_pago ??
          results.totais?.total_pago ??
          0,
        tir_mensal:
          results.tir_mensal ??
          results.metricas?.tir_mensal ??
          0,
        tir_anual:
          results.tir_anual ??
          results.metricas?.tir_anual ??
          0,
        delta_vpl:
          sc.vpl?.delta_vpl ??
          results.delta_vpl ??
          0,
        parcela_media:
          results.parcela_media ??
          results.metricas?.parcela_media ??
          0,
      });
    }

    return summaries;
  }

  /**
   * Export all scenarios as JSON string.
   *
   * @returns JSON string
   */
  exportJSON(): string {
    const data = {
      version: 1,
      exportDate: new Date().toISOString(),
      scenarios: Array.from(this.scenarios.entries()).map(([idx, sc]) => ({
        ...sc,
        index: idx,
      })),
      nextId: this.nextId,
    };
    return JSON.stringify(data, null, 2);
  }

  /**
   * Import scenarios from JSON string.
   * Replaces all existing scenarios.
   *
   * @param json - JSON string from exportJSON
   */
  importJSON(json: string): void {
    const data = JSON.parse(json);
    if (!data || !data.scenarios) {
      throw new Error('Formato JSON invalido para importacao de cenarios.');
    }

    this.scenarios.clear();

    for (const sc of data.scenarios) {
      const idx = sc.index ?? this.nextId;
      this.scenarios.set(idx, {
        index: idx,
        name: sc.name ?? `Cenario ${idx}`,
        params: sc.params ?? {},
        results: sc.results ?? {},
        vpl: sc.vpl ?? undefined,
        timestamp: sc.timestamp ?? Date.now(),
      });
    }

    this.nextId = data.nextId ?? this.scenarios.size + 1;
  }

  /**
   * Get the number of saved scenarios.
   */
  get count(): number {
    return this.scenarios.size;
  }

  /**
   * Check if max scenarios has been reached.
   */
  get isFull(): boolean {
    return this.scenarios.size >= this.maxScenarios;
  }
}

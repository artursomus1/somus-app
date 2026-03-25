import { useMemo, useCallback, useRef } from 'react';
import { NasaEngine } from '@/engine/nasa-engine';

/**
 * Hook que fornece uma instancia singleton do NasaEngine
 * e cacheia os ultimos resultados de cada calculo.
 */
export function useNasaEngine() {
  const engine = useMemo(() => new NasaEngine(), []);
  const lastResults = useRef<Record<string, any>>({});

  const calcularFluxo = useCallback(
    (params: Record<string, any>) => {
      const result = engine.calcularFluxoCompleto(params);
      lastResults.current.fluxo = result;
      lastResults.current.fluxoParams = params;
      return result;
    },
    [engine],
  );

  const calcularVPL = useCallback(
    (...args: Parameters<NasaEngine['calcularVPLHD']>) => {
      const result = engine.calcularVPLHD(...args);
      lastResults.current.vpl = result;
      return result;
    },
    [engine],
  );

  const calcularFinanciamento = useCallback(
    (params: Record<string, any>) => {
      const result = engine.calcularFinanciamento(params);
      lastResults.current.financiamento = result;
      lastResults.current.financiamentoParams = params;
      return result;
    },
    [engine],
  );

  const calcularCreditoLance = useCallback(
    (params: Record<string, any>) => {
      const result = engine.calcularCreditoLance(params);
      lastResults.current.creditoLance = result;
      return result;
    },
    [engine],
  );

  const calcularCustoCombinado = useCallback(
    (...args: Parameters<NasaEngine['calcularCustoCombinado']>) => {
      const result = engine.calcularCustoCombinado(...args);
      lastResults.current.custoCombinado = result;
      return result;
    },
    [engine],
  );

  const calcularVendaOperacao = useCallback(
    (...args: Parameters<NasaEngine['calcularVendaOperacao']>) => {
      const result = engine.calcularVendaOperacao(...args);
      lastResults.current.venda = result;
      return result;
    },
    [engine],
  );

  const calcularComparativo = useCallback(
    (...args: Parameters<NasaEngine['calcularCreditoEquivalente']>) => {
      const result = engine.calcularCreditoEquivalente(...args);
      lastResults.current.comparativo = result;
      return result;
    },
    [engine],
  );

  return {
    engine,
    calcularFluxo,
    calcularVPL,
    calcularFinanciamento,
    calcularCreditoLance,
    calcularCustoCombinado,
    calcularVendaOperacao,
    calcularComparativo,
    lastResults,
  };
}

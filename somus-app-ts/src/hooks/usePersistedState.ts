import { useState, useCallback } from 'react';
import { saveData, loadData } from '@/services/storage';

/**
 * React hook que persiste state automaticamente.
 * Carrega do storage na inicializacao e salva a cada mudanca.
 * Funciona tanto em Electron (via IPC fs) quanto em browser (localStorage).
 *
 * @param key Chave de armazenamento (sera prefixada com 'somus_')
 * @param defaultValue Valor padrao caso nao exista dados salvos
 * @returns Tupla [state, setState] compativel com useState
 */
export function usePersistedState<T>(
  key: string,
  defaultValue: T,
): [T, (value: T | ((prev: T) => T)) => void] {
  const [state, setState] = useState<T>(() => {
    try {
      return loadData<T>(key, defaultValue);
    } catch {
      return defaultValue;
    }
  });

  const setPersistedState = useCallback(
    (value: T | ((prev: T) => T)) => {
      setState((prev) => {
        const newValue =
          typeof value === 'function'
            ? (value as (prev: T) => T)(prev)
            : value;
        try {
          saveData(key, newValue);
        } catch (err) {
          console.error(`[usePersistedState] Erro ao persistir "${key}":`, err);
        }
        return newValue;
      });
    },
    [key],
  );

  return [state, setPersistedState];
}

/**
 * Servico de persistencia local usando electron-store
 * com fallback para localStorage no browser
 */

let electronStore: any = null;

// Tenta carregar electron-store (funciona apenas em contexto Electron)
try {
  if (typeof window !== 'undefined' && (window as any).electronStore) {
    electronStore = (window as any).electronStore;
  }
} catch {
  // fallback para localStorage
}

/**
 * Salva dados no armazenamento local
 */
export function saveData(key: string, data: any): void {
  try {
    if (electronStore) {
      electronStore.set(key, data);
    } else {
      localStorage.setItem(key, JSON.stringify(data));
    }
  } catch (err) {
    console.error(`[Storage] Erro ao salvar "${key}":`, err);
  }
}

/**
 * Carrega dados do armazenamento local
 */
export function loadData<T>(key: string, defaultValue: T): T {
  try {
    if (electronStore) {
      const val = electronStore.get(key);
      return val !== undefined ? val : defaultValue;
    } else {
      const raw = localStorage.getItem(key);
      if (raw === null) return defaultValue;
      return JSON.parse(raw) as T;
    }
  } catch (err) {
    console.error(`[Storage] Erro ao carregar "${key}":`, err);
    return defaultValue;
  }
}

/**
 * Remove dados do armazenamento local
 */
export function deleteData(key: string): void {
  try {
    if (electronStore) {
      electronStore.delete(key);
    } else {
      localStorage.removeItem(key);
    }
  } catch (err) {
    console.error(`[Storage] Erro ao deletar "${key}":`, err);
  }
}

/**
 * Lista todas as chaves armazenadas
 */
export function listKeys(): string[] {
  try {
    if (electronStore) {
      return Object.keys(electronStore.store || {});
    } else {
      return Object.keys(localStorage);
    }
  } catch {
    return [];
  }
}

/**
 * Limpa todos os dados armazenados
 */
export function clearAll(): void {
  try {
    if (electronStore) {
      electronStore.clear();
    } else {
      localStorage.clear();
    }
  } catch (err) {
    console.error('[Storage] Erro ao limpar dados:', err);
  }
}

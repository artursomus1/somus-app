/**
 * Servico de persistencia local
 * - Em Electron: usa IPC fs para salvar em somus_data/*.json no userData
 * - Em browser: usa localStorage com prefixo somus_
 * Fallback automatico para localStorage em caso de erro
 */

const STORAGE_PREFIX = 'somus_';
const DATA_DIR = 'somus_data';

// ── Helpers ────────────────────────────────────────────────────────────────────

function isElectron(): boolean {
  return typeof window !== 'undefined' && !!window.electron?.fs;
}

async function ensureDataDir(): Promise<void> {
  if (!isElectron()) return;
  try {
    const exists = await window.electron!.fs.exists(DATA_DIR);
    if (!exists) {
      await window.electron!.fs.mkdir(DATA_DIR);
    }
  } catch {
    // ignore - will fallback to localStorage
  }
}

// Track if dir was already ensured this session
let dirEnsured = false;

// ── Sync API (localStorage only) ──────────────────────────────────────────────

function saveToLocalStorage(key: string, data: any): void {
  localStorage.setItem(STORAGE_PREFIX + key, JSON.stringify(data));
}

function loadFromLocalStorage<T>(key: string, defaultValue: T): T {
  const raw = localStorage.getItem(STORAGE_PREFIX + key);
  if (raw === null) return defaultValue;
  return JSON.parse(raw) as T;
}

// ── Public API ─────────────────────────────────────────────────────────────────

/**
 * Salva dados no armazenamento local.
 * Em Electron, salva async via IPC e tambem no localStorage como backup.
 * Em browser, salva apenas no localStorage.
 */
export function saveData(key: string, data: any): void {
  try {
    // Always save to localStorage as immediate cache
    saveToLocalStorage(key, data);

    if (isElectron()) {
      // Fire-and-forget async write to fs
      const filePath = `${DATA_DIR}/${key}.json`;
      const jsonStr = JSON.stringify(data, null, 2);

      const doWrite = async () => {
        if (!dirEnsured) {
          await ensureDataDir();
          dirEnsured = true;
        }
        await window.electron!.fs.writeFile(filePath, jsonStr);
      };

      doWrite().catch((err) => {
        console.error(`[Storage] Erro ao salvar "${key}" via IPC:`, err);
      });
    }
  } catch (err) {
    console.error(`[Storage] Erro ao salvar "${key}":`, err);
    // Last resort fallback
    try {
      saveToLocalStorage(key, data);
    } catch {
      // nothing more we can do
    }
  }
}

/**
 * Carrega dados do armazenamento local (sincrono).
 * Leitura eh feita do localStorage para manter API sincrona.
 * Os dados sao mantidos em sync pelo saveData que grava em ambos.
 */
export function loadData<T>(key: string, defaultValue: T): T {
  try {
    const raw = localStorage.getItem(STORAGE_PREFIX + key);
    if (raw === null) return defaultValue;
    return JSON.parse(raw) as T;
  } catch (err) {
    console.error(`[Storage] Erro ao carregar "${key}":`, err);
    return defaultValue;
  }
}

/**
 * Carrega dados de forma assincrona - tenta Electron fs primeiro,
 * depois fallback para localStorage.
 */
export async function loadDataAsync<T>(key: string, defaultValue: T): Promise<T> {
  try {
    if (isElectron()) {
      const filePath = `${DATA_DIR}/${key}.json`;
      const exists = await window.electron!.fs.exists(filePath);
      if (exists) {
        const result = await window.electron!.fs.readFile(filePath);
        if (result.success && result.data) {
          const parsed = JSON.parse(result.data) as T;
          // Sync to localStorage
          saveToLocalStorage(key, parsed);
          return parsed;
        }
      }
    }
    // Fallback to localStorage
    return loadFromLocalStorage(key, defaultValue);
  } catch (err) {
    console.error(`[Storage] Erro ao carregar async "${key}":`, err);
    // Try localStorage as last resort
    try {
      return loadFromLocalStorage(key, defaultValue);
    } catch {
      return defaultValue;
    }
  }
}

/**
 * Remove dados do armazenamento local
 */
export function deleteData(key: string): void {
  try {
    localStorage.removeItem(STORAGE_PREFIX + key);

    if (isElectron()) {
      // Overwrite with empty to "delete" (no fs.delete in preload)
      const filePath = `${DATA_DIR}/${key}.json`;
      window.electron!.fs.writeFile(filePath, '').catch(() => {});
    }
  } catch (err) {
    console.error(`[Storage] Erro ao deletar "${key}":`, err);
  }
}

/**
 * Lista todas as chaves armazenadas (apenas localStorage - sincrono)
 */
export function listKeys(): string[] {
  try {
    return Object.keys(localStorage)
      .filter((k) => k.startsWith(STORAGE_PREFIX))
      .map((k) => k.slice(STORAGE_PREFIX.length));
  } catch {
    return [];
  }
}

/**
 * Limpa todos os dados com prefixo somus_
 */
export function clearAll(): void {
  try {
    const keys = Object.keys(localStorage).filter((k) => k.startsWith(STORAGE_PREFIX));
    keys.forEach((k) => localStorage.removeItem(k));
  } catch (err) {
    console.error('[Storage] Erro ao limpar dados:', err);
  }
}

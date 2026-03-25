import { create } from 'zustand';
import { saveData, loadData } from '@/services/storage';
import type { Operacao } from '@/types';

// ── Types ────────────────────────────────────────────────────────────────────

interface SavedScenario {
  id: number;
  nome: string;
  params: Record<string, any>;
  parcelaF1: number;
  parcelaF2: number;
  totalPago: number;
  cartaLiquida: number;
  tirMensal: number;
  cetAnual: number;
  lanceLivreValor: number;
  lanceEmbutidoValor: number;
  timestamp: string;
}

interface CustoAcessorio {
  id: string;
  descricao: string;
  valor: number;
  momento: number;
}

// ── Store ────────────────────────────────────────────────────────────────────

interface DataStore {
  // Last calculation results (shared between pages, not persisted)
  lastParams: Record<string, any> | null;
  lastFluxo: any | null;
  lastVPL: any | null;
  lastFinanciamento: any | null;
  lastComparativo: any | null;
  lastVenda: any | null;
  lastCreditoLance: any | null;
  lastCustoCombinado: any | null;

  // Persisted data
  operacoes: Operacao[];
  cenarios: SavedScenario[];
  custosAcessorios: CustoAcessorio[];

  // ── Actions: calculation results ──────────────────────────────────
  setLastFluxo: (params: Record<string, any>, fluxo: any, vpl?: any) => void;
  setLastFinanciamento: (result: any) => void;
  setLastComparativo: (result: any) => void;
  setLastVenda: (result: any) => void;
  setLastCreditoLance: (result: any) => void;
  setLastCustoCombinado: (result: any) => void;

  // ── Actions: operacoes ────────────────────────────────────────────
  setOperacoes: (ops: Operacao[]) => void;
  addOperacao: (op: Operacao) => void;
  updateOperacao: (id: string, op: Operacao) => void;
  removeOperacao: (id: string) => void;

  // ── Actions: cenarios ─────────────────────────────────────────────
  setCenarios: (list: SavedScenario[]) => void;
  addCenario: (cenario: SavedScenario) => void;
  removeCenario: (id: number) => void;

  // ── Actions: custos acessorios ────────────────────────────────────
  setCustosAcessorios: (list: CustoAcessorio[]) => void;
  addCustoAcessorio: (custo: CustoAcessorio) => void;
  removeCustoAcessorio: (id: string) => void;

  // ── Hydrate from storage ──────────────────────────────────────────
  hydrate: () => void;
}

export const useDataStore = create<DataStore>((set, get) => ({
  // ── Initial state ───────────────────────────────────────────────────────
  lastParams: null,
  lastFluxo: null,
  lastVPL: null,
  lastFinanciamento: null,
  lastComparativo: null,
  lastVenda: null,
  lastCreditoLance: null,
  lastCustoCombinado: null,

  operacoes: [],
  cenarios: [],
  custosAcessorios: [],

  // ── Calculation results ─────────────────────────────────────────────────

  setLastFluxo: (params, fluxo, vpl) =>
    set({ lastParams: params, lastFluxo: fluxo, lastVPL: vpl ?? null }),

  setLastFinanciamento: (result) =>
    set({ lastFinanciamento: result }),

  setLastComparativo: (result) =>
    set({ lastComparativo: result }),

  setLastVenda: (result) =>
    set({ lastVenda: result }),

  setLastCreditoLance: (result) =>
    set({ lastCreditoLance: result }),

  setLastCustoCombinado: (result) =>
    set({ lastCustoCombinado: result }),

  // ── Operacoes ───────────────────────────────────────────────────────────

  setOperacoes: (ops) => {
    set({ operacoes: ops });
    saveData('operacoes', ops);
  },

  addOperacao: (op) => {
    const updated = [...get().operacoes, op];
    set({ operacoes: updated });
    saveData('operacoes', updated);
  },

  updateOperacao: (id, op) => {
    const updated = get().operacoes.map((o) => (o.id === id ? op : o));
    set({ operacoes: updated });
    saveData('operacoes', updated);
  },

  removeOperacao: (id) => {
    const updated = get().operacoes.filter((o) => o.id !== id);
    set({ operacoes: updated });
    saveData('operacoes', updated);
  },

  // ── Cenarios ────────────────────────────────────────────────────────────

  setCenarios: (list) => {
    set({ cenarios: list });
    saveData('cenarios', list);
  },

  addCenario: (cenario) => {
    const updated = [...get().cenarios, cenario];
    set({ cenarios: updated });
    saveData('cenarios', updated);
  },

  removeCenario: (id) => {
    const updated = get().cenarios.filter((c) => c.id !== id);
    set({ cenarios: updated });
    saveData('cenarios', updated);
  },

  // ── Custos acessorios ───────────────────────────────────────────────────

  setCustosAcessorios: (list) => {
    set({ custosAcessorios: list });
    saveData('custos_acessorios', list);
  },

  addCustoAcessorio: (custo) => {
    const updated = [...get().custosAcessorios, custo];
    set({ custosAcessorios: updated });
    saveData('custos_acessorios', updated);
  },

  removeCustoAcessorio: (id) => {
    const updated = get().custosAcessorios.filter((c) => c.id !== id);
    set({ custosAcessorios: updated });
    saveData('custos_acessorios', updated);
  },

  // ── Hydrate ─────────────────────────────────────────────────────────────

  hydrate: () => {
    set({
      operacoes: loadData<Operacao[]>('operacoes', []),
      cenarios: loadData<SavedScenario[]>('cenarios', []),
      custosAcessorios: loadData<CustoAcessorio[]>('custos_acessorios', []),
    });
  },
}));

// Auto-hydrate on import
if (typeof window !== 'undefined') {
  // Defer hydration to next tick so localStorage is ready
  setTimeout(() => {
    useDataStore.getState().hydrate();
  }, 0);
}

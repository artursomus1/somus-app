import { create } from 'zustand';
import type { ModuleName, PageKey, Role } from '../types';

// ── Page definitions per module ────────────────────────────────────────────────

export interface PageDef {
  key: PageKey;
  label: string;
  icon: string;
}

export const MODULE_LABELS: Record<ModuleName, string> = {
  mesa_produtos: 'Mesa Produtos',
  corporate: 'Corporate',
  seguros: 'Seguros',
};

export const MODULE_PAGES: Record<ModuleName, PageDef[]> = {
  mesa_produtos: [
    { key: 'dashboard', label: 'Dashboard', icon: 'LayoutDashboard' },
    { key: 'fluxo_rf', label: 'Fluxo RF', icon: 'FileText' },
    { key: 'informativo', label: 'Informativo', icon: 'Newspaper' },
    { key: 'info_agio', label: 'Info Agio', icon: 'TrendingUp' },
    { key: 'envio_ordens', label: 'Envio Ordens', icon: 'Send' },
    { key: 'ctrl_receita', label: 'Ctrl Receita', icon: 'DollarSign' },
    { key: 'organizador', label: 'Organizador', icon: 'FolderOpen' },
    { key: 'consolidador', label: 'Consolidador', icon: 'Layers' },
    { key: 'envio_saldos', label: 'Envio Saldos', icon: 'Mail' },
    { key: 'envio_mesa', label: 'Envio Mesa', icon: 'MailPlus' },
    { key: 'envio_aniversarios', label: 'Envio Aniversarios', icon: 'Cake' },
    { key: 'tarefas', label: 'Tarefas', icon: 'CheckSquare' },
  ],
  corporate: [
    { key: 'dashboard', label: 'Dashboard', icon: 'LayoutDashboard' },
    { key: 'simulador', label: 'Simulador', icon: 'Calculator' },
    { key: 'comparativo-vpl', label: 'Comparativo VPL', icon: 'TrendingUp' },
    { key: 'consorcio-vs-financ', label: 'Consorcio vs Financ', icon: 'Layers' },
    { key: 'gerador-propostas', label: 'Gerador Propostas', icon: 'FileText' },
    { key: 'fluxo-receitas', label: 'Fluxo Receitas', icon: 'DollarSign' },
    { key: 'cenarios', label: 'Cenarios', icon: 'FolderOpen' },
  ],
  seguros: [
    { key: 'renovacoes', label: 'Renovacoes Anuais', icon: 'ShieldCheck' },
  ],
};

// ── Roles and which modules they can access ────────────────────────────────────

export const ROLE_MODULES: Record<Role, ModuleName[]> = {
  admin: ['mesa_produtos', 'corporate', 'seguros'],
  mesa_produtos: ['mesa_produtos'],
  corporate: ['corporate'],
  seguros: ['seguros'],
};

// ── Store ──────────────────────────────────────────────────────────────────────

interface AppStore {
  currentModule: ModuleName;
  currentPage: PageKey;
  role: Role;
  userName: string;
  sidebarCollapsed: boolean;

  setModule: (module: ModuleName) => void;
  setPage: (page: PageKey) => void;
  setRole: (role: Role) => void;
  setUserName: (name: string) => void;
  toggleSidebar: () => void;
}

export const useAppStore = create<AppStore>((set) => ({
  currentModule: 'mesa_produtos',
  currentPage: 'dashboard',
  role: 'admin',
  userName: 'Artur Brito',
  sidebarCollapsed: false,

  setModule: (module) =>
    set(() => {
      const pages = MODULE_PAGES[module];
      return {
        currentModule: module,
        currentPage: pages[0]?.key ?? 'dashboard',
      };
    }),

  setPage: (page) =>
    set({ currentPage: page }),

  setRole: (role) =>
    set((state) => {
      const allowedModules = ROLE_MODULES[role];
      const moduleOk = allowedModules.includes(state.currentModule);
      const newModule = moduleOk ? state.currentModule : allowedModules[0];
      const pages = MODULE_PAGES[newModule];
      return {
        role,
        currentModule: newModule,
        currentPage: pages[0]?.key ?? 'dashboard',
      };
    }),

  setUserName: (name) =>
    set({ userName: name }),

  toggleSidebar: () =>
    set((state) => ({ sidebarCollapsed: !state.sidebarCollapsed })),
}));

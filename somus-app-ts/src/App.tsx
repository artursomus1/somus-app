import React from 'react';
import {
  LayoutDashboard,
  FileText,
  Newspaper,
  TrendingUp,
  Send,
  DollarSign,
  FolderOpen,
  Layers,
  Mail,
  MailPlus,
  Cake,
  CheckSquare,
  Calculator,
  ShieldCheck,
  ChevronLeft,
  ChevronRight,
  Minus,
  Square,
  X,
  Building2,
  Briefcase,
  Shield,
} from 'lucide-react';
import {
  useAppStore,
  MODULE_LABELS,
  MODULE_PAGES,
  ROLE_MODULES,
  type PageDef,
} from './stores/appStore';
import type { ModuleName } from './types';
import { cn } from './utils/cn';

// ── Page imports ─────────────────────────────────────────────────────────────

// Mesa Produtos
import MPDashboard from './pages/mesa-produtos/Dashboard';
import FluxoRF from './pages/mesa-produtos/FluxoRF';
import Informativo from './pages/mesa-produtos/Informativo';
import InfoAgio from './pages/mesa-produtos/InfoAgio';
import EnvioOrdens from './pages/mesa-produtos/EnvioOrdens';
import CtrlReceita from './pages/mesa-produtos/CtrlReceita';
import Organizador from './pages/mesa-produtos/Organizador';
import Consolidador from './pages/mesa-produtos/Consolidador';
import EnvioSaldos from './pages/mesa-produtos/EnvioSaldos';
import EnvioMesa from './pages/mesa-produtos/EnvioMesa';
import EnvioAniversarios from './pages/mesa-produtos/EnvioAniversarios';
import Tarefas from './pages/mesa-produtos/Tarefas';

// Corporate
import CorpDashboard from './pages/corporate/Dashboard';
import Simulador from './pages/corporate/Simulador';
import ComparativoVPL from './pages/corporate/ComparativoVPL';
import ConsorcioVsFinanc from './pages/corporate/ConsorcioVsFinanc';
import GeradorPropostas from './pages/corporate/GeradorPropostas';
import FluxoReceitas from './pages/corporate/FluxoReceitas';
import Cenarios from './pages/corporate/Cenarios';

// Seguros
import RenovacoesAnuais from './pages/seguros/RenovacoesAnuais';

// ── Page registry ────────────────────────────────────────────────────────────

const PAGE_COMPONENTS: Record<string, Record<string, React.FC>> = {
  mesa_produtos: {
    dashboard: MPDashboard,
    fluxo_rf: FluxoRF,
    informativo: Informativo,
    info_agio: InfoAgio,
    envio_ordens: EnvioOrdens,
    ctrl_receita: CtrlReceita,
    organizador: Organizador,
    consolidador: Consolidador,
    envio_saldos: EnvioSaldos,
    envio_mesa: EnvioMesa,
    envio_aniversarios: EnvioAniversarios,
    tarefas: Tarefas,
  },
  corporate: {
    dashboard: CorpDashboard,
    simulador: Simulador,
    'comparativo-vpl': ComparativoVPL,
    'consorcio-vs-financ': ConsorcioVsFinanc,
    'gerador-propostas': GeradorPropostas,
    'fluxo-receitas': FluxoReceitas,
    cenarios: Cenarios,
  },
  seguros: {
    renovacoes: RenovacoesAnuais,
  },
};

// ── Icon map ─────────────────────────────────────────────────────────────────

// eslint-disable-next-line @typescript-eslint/no-explicit-any
const ICON_MAP: Record<string, any> = {
  LayoutDashboard,
  FileText,
  Newspaper,
  TrendingUp,
  Send,
  DollarSign,
  FolderOpen,
  Layers,
  Mail,
  MailPlus,
  Cake,
  CheckSquare,
  Calculator,
  ShieldCheck,
};

// eslint-disable-next-line @typescript-eslint/no-explicit-any
const MODULE_ICONS: Record<ModuleName, any> = {
  mesa_produtos: Briefcase,
  corporate: Building2,
  seguros: Shield,
};

// ── Titlebar ─────────────────────────────────────────────────────────────────

function Titlebar() {
  const handleMinimize = () => window.electron?.app.minimize();
  const handleMaximize = () => window.electron?.app.maximize();
  const handleClose = () => window.electron?.app.close();

  return (
    <div className="titlebar-drag flex items-center justify-between h-9 bg-neutral-950 border-b border-neutral-800 px-4 shrink-0">
      <div className="flex items-center gap-2">
        <span className="somus-gradient-text text-sm font-bold tracking-wide">SOMUS CAPITAL</span>
      </div>
      <div className="titlebar-no-drag flex items-center">
        <button onClick={handleMinimize} className="flex items-center justify-center w-11 h-9 hover:bg-neutral-800 transition-colors" aria-label="Minimizar">
          <Minus size={14} className="text-neutral-400" />
        </button>
        <button onClick={handleMaximize} className="flex items-center justify-center w-11 h-9 hover:bg-neutral-800 transition-colors" aria-label="Maximizar">
          <Square size={12} className="text-neutral-400" />
        </button>
        <button onClick={handleClose} className="flex items-center justify-center w-11 h-9 hover:bg-red-600 transition-colors" aria-label="Fechar">
          <X size={14} className="text-neutral-400" />
        </button>
      </div>
    </div>
  );
}

// ── Module Selector ──────────────────────────────────────────────────────────

function ModuleSelector() {
  const { currentModule, setModule, role } = useAppStore();
  const allowedModules = ROLE_MODULES[role];

  return (
    <div className="flex gap-1 p-2">
      {allowedModules.map((mod) => {
        const Icon = MODULE_ICONS[mod];
        const isActive = currentModule === mod;
        return (
          <button
            key={mod}
            onClick={() => setModule(mod)}
            className={cn(
              'flex-1 flex items-center justify-center gap-1.5 px-2 py-2 rounded-lg text-xs font-medium transition-all duration-150',
              isActive
                ? 'bg-emerald-600/15 text-emerald-400 border border-emerald-500/30'
                : 'text-neutral-500 hover:text-neutral-300 hover:bg-neutral-800/60 border border-transparent',
            )}
            title={MODULE_LABELS[mod]}
          >
            <Icon size={14} />
            <span className="hidden xl:inline">{MODULE_LABELS[mod]}</span>
          </button>
        );
      })}
    </div>
  );
}

// ── Sidebar ──────────────────────────────────────────────────────────────────

function AppSidebar() {
  const { currentModule, currentPage, setPage, sidebarCollapsed, toggleSidebar, userName } =
    useAppStore();
  const pages = MODULE_PAGES[currentModule];

  return (
    <aside
      className={cn(
        'flex flex-col bg-neutral-950 border-r border-neutral-800 transition-all duration-200 shrink-0',
        sidebarCollapsed ? 'w-16' : 'w-60',
      )}
    >
      {!sidebarCollapsed && <ModuleSelector />}
      <div className="h-px bg-neutral-800 mx-3" />
      <nav className="flex-1 overflow-y-auto py-2 px-2 space-y-0.5">
        {pages.map((page: PageDef) => {
          const Icon = ICON_MAP[page.icon];
          const isActive = currentPage === page.key;
          return (
            <button
              key={page.key}
              onClick={() => setPage(page.key)}
              className={cn(
                'nav-item w-full',
                isActive && 'active',
                sidebarCollapsed && 'justify-center px-0',
              )}
              title={page.label}
            >
              {Icon && <Icon size={18} />}
              {!sidebarCollapsed && <span>{page.label}</span>}
            </button>
          );
        })}
      </nav>
      <div className="h-px bg-neutral-800 mx-3" />
      <div className="p-3 flex items-center gap-2">
        {!sidebarCollapsed && (
          <div className="flex-1 min-w-0">
            <p className="text-xs font-medium text-neutral-300 truncate">{userName}</p>
            <p className="text-[10px] text-neutral-600">Somus Capital</p>
          </div>
        )}
        <button
          onClick={toggleSidebar}
          className="p-1.5 rounded-md hover:bg-neutral-800 text-neutral-500 hover:text-neutral-300 transition-colors"
          aria-label={sidebarCollapsed ? 'Expandir sidebar' : 'Recolher sidebar'}
        >
          {sidebarCollapsed ? <ChevronRight size={16} /> : <ChevronLeft size={16} />}
        </button>
      </div>
    </aside>
  );
}

// ── Page Content (renders actual page components) ────────────────────────────

function PageContent() {
  const { currentModule, currentPage } = useAppStore();

  const modulePages = PAGE_COMPONENTS[currentModule];
  const PageComponent = modulePages?.[currentPage];

  if (PageComponent) {
    return (
      <div className="flex-1 overflow-auto bg-somus-gray-50 animate-fade-in">
        <PageComponent />
      </div>
    );
  }

  // Fallback for unmapped pages
  const pages = MODULE_PAGES[currentModule];
  const pageDef = pages.find((p) => p.key === currentPage);
  const Icon = pageDef ? ICON_MAP[pageDef.icon] : null;

  return (
    <div className="flex-1 overflow-auto">
      <header className="sticky top-0 z-20 bg-neutral-950/80 backdrop-blur-md border-b border-neutral-800 px-6 py-4">
        <div className="flex items-center gap-3">
          {Icon && <Icon size={22} className="text-emerald-500" />}
          <div>
            <h1 className="text-lg font-semibold text-neutral-100">{pageDef?.label ?? currentPage}</h1>
            <p className="text-xs text-neutral-500">{MODULE_LABELS[currentModule]}</p>
          </div>
        </div>
      </header>
      <main className="p-6 animate-fade-in">
        <div className="glass p-8 flex flex-col items-center justify-center min-h-[400px] gap-4">
          {Icon && <Icon size={48} className="text-neutral-700" />}
          <h2 className="text-xl font-semibold text-neutral-400">{pageDef?.label ?? currentPage}</h2>
          <p className="text-sm text-neutral-600 max-w-md text-center">Pagina em desenvolvimento.</p>
        </div>
      </main>
    </div>
  );
}

// ── App ──────────────────────────────────────────────────────────────────────

export default function App() {
  return (
    <div className="flex flex-col h-screen w-screen overflow-hidden bg-neutral-950">
      <Titlebar />
      <div className="flex flex-1 overflow-hidden">
        <AppSidebar />
        <PageContent />
      </div>
    </div>
  );
}

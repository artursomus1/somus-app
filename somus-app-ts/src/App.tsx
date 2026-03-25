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
  Table,
  List,
  LineChart,
  Presentation,
  BadgeDollarSign,
  Landmark,
  Receipt,
  Combine,
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
import FluxoFinanceiro from './pages/corporate/FluxoFinanceiro';
import Parcelas from './pages/corporate/Parcelas';
import GraficoView from './pages/corporate/GraficoView';
import ResumoCliente from './pages/corporate/ResumoCliente';
import GeradorPropostas from './pages/corporate/GeradorPropostas';
import FluxoReceitas from './pages/corporate/FluxoReceitas';
import Cenarios from './pages/corporate/Cenarios';
import VendaOperacao from './pages/corporate/VendaOperacao';
import CreditoLance from './pages/corporate/CreditoLance';
import ConsolidacaoCotas from './pages/corporate/ConsolidacaoCotas';
import CustosAcessorios from './pages/corporate/CustosAcessorios';
import CustoCombinado from './pages/corporate/CustoCombinado';

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
    'fluxo-financeiro': FluxoFinanceiro,
    parcelas: Parcelas,
    'comparativo-vpl': ComparativoVPL,
    'consorcio-vs-financ': ConsorcioVsFinanc,
    grafico: GraficoView,
    'resumo-cliente': ResumoCliente,
    cenarios: Cenarios,
    'gerador-propostas': GeradorPropostas,
    'fluxo-receitas': FluxoReceitas,
    'venda-operacao': VendaOperacao,
    'credito-lance': CreditoLance,
    'consolidacao': ConsolidacaoCotas,
    'custos-acessorios': CustosAcessorios,
    'custo-combinado': CustoCombinado,
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
  Table,
  List,
  LineChart,
  Presentation,
  BadgeDollarSign,
  Landmark,
  Receipt,
  Combine,
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
    <div className="titlebar-drag relative flex items-center justify-between h-9 bg-somus-bg-primary border-b border-somus-border px-4 shrink-0">
      <div className="flex items-center gap-2">
        <span className="somus-gradient-text text-sm font-bold tracking-wide">SOMUS CAPITAL</span>
      </div>
      <div className="titlebar-no-drag flex items-center">
        <button onClick={handleMinimize} className="flex items-center justify-center w-11 h-9 hover:bg-somus-bg-hover transition-colors" aria-label="Minimizar">
          <Minus size={14} className="text-somus-text-tertiary" />
        </button>
        <button onClick={handleMaximize} className="flex items-center justify-center w-11 h-9 hover:bg-somus-bg-hover transition-colors" aria-label="Maximizar">
          <Square size={12} className="text-somus-text-tertiary" />
        </button>
        <button onClick={handleClose} className="flex items-center justify-center w-11 h-9 hover:bg-red-600/80 transition-colors" aria-label="Fechar">
          <X size={14} className="text-somus-text-tertiary" />
        </button>
      </div>
      {/* Subtle green gradient line at bottom */}
      <div className="absolute bottom-0 left-0 right-0 h-px" style={{ background: 'linear-gradient(90deg, transparent, #1A7A3E, transparent)' }} />
    </div>
  );
}

// ── Module Selector ──────────────────────────────────────────────────────────

function ModuleSelector() {
  const { currentModule, setModule, role } = useAppStore();
  const allowedModules = ROLE_MODULES[role];
  const activeIndex = allowedModules.indexOf(currentModule);

  return (
    <div className="p-2">
      <div className="relative flex bg-somus-bg-input rounded-lg p-0.5">
        {/* Sliding indicator */}
        <div
          className="absolute top-0.5 bottom-0.5 rounded-md bg-gradient-to-r from-somus-green-500 to-somus-green-600 transition-all duration-300 ease-out"
          style={{
            width: `${100 / allowedModules.length}%`,
            left: `${(activeIndex * 100) / allowedModules.length}%`,
          }}
        />
        {/* Module buttons */}
        {allowedModules.map((mod) => {
          const Icon = MODULE_ICONS[mod];
          const isActive = currentModule === mod;
          return (
            <button
              key={mod}
              onClick={() => setModule(mod)}
              className={cn(
                'relative z-10 flex-1 flex items-center justify-center gap-1.5 px-2 py-2 rounded-md text-xs font-medium transition-colors duration-200',
                isActive ? 'text-white' : 'text-somus-text-tertiary hover:text-somus-text-secondary',
              )}
            >
              <Icon size={14} />
              <span>{MODULE_LABELS[mod]}</span>
            </button>
          );
        })}
      </div>
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
        'flex flex-col bg-somus-bg-primary border-r border-somus-border transition-all duration-200 shrink-0',
        sidebarCollapsed ? 'w-16' : 'w-60',
      )}
    >
      {!sidebarCollapsed && <ModuleSelector />}
      <div className="h-px bg-somus-border mx-3" />
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
      <div className="h-px bg-somus-border mx-3" />
      <div className="p-3 flex items-center gap-2">
        {!sidebarCollapsed && (
          <>
            <div className="h-7 w-7 rounded-full bg-somus-bg-tertiary border border-somus-border flex items-center justify-center text-xs font-bold text-somus-text-accent shrink-0">
              {userName?.charAt(0)?.toUpperCase() ?? 'U'}
            </div>
            <div className="flex-1 min-w-0">
              <p className="text-xs font-medium text-somus-text-primary truncate">{userName}</p>
              <p className="text-[10px] text-somus-text-tertiary">Somus Capital</p>
            </div>
          </>
        )}
        <button
          onClick={toggleSidebar}
          className="p-1.5 rounded-md hover:bg-somus-bg-hover text-somus-text-tertiary hover:text-somus-text-secondary transition-colors"
          aria-label={sidebarCollapsed ? 'Expandir sidebar' : 'Recolher sidebar'}
        >
          {sidebarCollapsed ? <ChevronRight size={16} /> : <ChevronLeft size={16} />}
        </button>
      </div>
      {/* Version badge */}
      {!sidebarCollapsed && (
        <div className="px-3 pb-2">
          <span className="text-[10px] font-mono text-somus-text-tertiary">v2.0.0</span>
        </div>
      )}
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
      <div className="flex-1 overflow-auto bg-somus-bg-primary animate-fade-in">
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
      <header className="sticky top-0 z-20 bg-somus-bg-primary/80 backdrop-blur-md border-b border-somus-border px-6 py-4">
        <div className="flex items-center gap-3">
          {Icon && <Icon size={22} className="text-somus-green-500" />}
          <div>
            <h1 className="text-lg font-semibold text-somus-text-primary">{pageDef?.label ?? currentPage}</h1>
            <p className="text-xs text-somus-text-tertiary">{MODULE_LABELS[currentModule]}</p>
          </div>
        </div>
      </header>
      <main className="p-6 animate-fade-in">
        <div className="glass p-8 flex flex-col items-center justify-center min-h-[400px] gap-4">
          {Icon && <Icon size={48} className="text-somus-text-tertiary" />}
          <h2 className="text-xl font-semibold text-somus-text-secondary">{pageDef?.label ?? currentPage}</h2>
          <p className="text-sm text-somus-text-tertiary max-w-md text-center">Pagina em desenvolvimento.</p>
        </div>
      </main>
    </div>
  );
}

// ── App ──────────────────────────────────────────────────────────────────────

export default function App() {
  return (
    <div className="flex flex-col h-screen w-screen overflow-hidden bg-somus-bg-primary">
      <Titlebar />
      <div className="flex flex-1 overflow-hidden">
        <AppSidebar />
        <PageContent />
      </div>
    </div>
  );
}

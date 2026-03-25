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
  ChevronDown,
  PanelLeftClose,
  PanelLeftOpen,
} from 'lucide-react';
import {
  useAppStore,
  MODULE_PAGES,
  MODULE_LABELS,
  ROLE_MODULES,
  type PageDef,
} from '@/stores/appStore';
import type { ModuleName } from '@/types';
import { cn } from '@/utils/cn';
import logoSomus from '@/assets/logo_somus.png';

// ─── Icon map (string name -> Lucide component) ─────────────────────

const ICON_SIZE = 'h-[18px] w-[18px]';

const iconMap: Record<string, React.ReactNode> = {
  LayoutDashboard: <LayoutDashboard className={ICON_SIZE} />,
  FileText: <FileText className={ICON_SIZE} />,
  Newspaper: <Newspaper className={ICON_SIZE} />,
  TrendingUp: <TrendingUp className={ICON_SIZE} />,
  Send: <Send className={ICON_SIZE} />,
  DollarSign: <DollarSign className={ICON_SIZE} />,
  FolderOpen: <FolderOpen className={ICON_SIZE} />,
  Layers: <Layers className={ICON_SIZE} />,
  Mail: <Mail className={ICON_SIZE} />,
  MailPlus: <MailPlus className={ICON_SIZE} />,
  Cake: <Cake className={ICON_SIZE} />,
  CheckSquare: <CheckSquare className={ICON_SIZE} />,
  Calculator: <Calculator className={ICON_SIZE} />,
  ShieldCheck: <ShieldCheck className={ICON_SIZE} />,
};

// ─── Module options for the dropdown ─────────────────────────────────

const allModules: ModuleName[] = ['mesa_produtos', 'corporate', 'seguros'];

// ─── Component ───────────────────────────────────────────────────────

export function Sidebar() {
  const {
    currentModule,
    currentPage,
    sidebarCollapsed,
    role,
    setModule,
    setPage,
    toggleSidebar,
  } = useAppStore();

  // Role-based module filtering
  const allowedModules = ROLE_MODULES[role];
  const availableModules = allModules.filter((m) => allowedModules.includes(m));

  const navItems: PageDef[] = MODULE_PAGES[currentModule];

  return (
    <aside
      className={cn(
        'flex flex-col h-full bg-somus-bg-primary border-r border-somus-border transition-all duration-300 ease-in-out shrink-0',
        sidebarCollapsed ? 'w-16' : 'w-60'
      )}
    >
      {/* ── Logo ── */}
      <div className="flex items-center justify-center h-14 px-3 border-b border-somus-border">
        {sidebarCollapsed ? (
          <img src={logoSomus} alt="Somus" className="h-7 w-7 object-contain" />
        ) : (
          <img src={logoSomus} alt="Somus Capital" className="h-8 object-contain" />
        )}
      </div>

      {/* ── Module Selector ── */}
      {!sidebarCollapsed && availableModules.length > 1 && (
        <div className="px-3 pt-4 pb-2">
          <div className="relative">
            <select
              value={currentModule}
              onChange={(e) => setModule(e.target.value as ModuleName)}
              className="w-full appearance-none bg-somus-bg-input border border-somus-border text-somus-text-primary text-sm font-medium rounded-lg px-3 py-2 pr-8 focus:outline-none focus:ring-2 focus:ring-somus-green-500/30 focus:border-somus-green-500/50 cursor-pointer transition-colors"
            >
              {availableModules.map((m) => (
                <option
                  key={m}
                  value={m}
                  className="bg-somus-bg-secondary text-somus-text-primary"
                >
                  {MODULE_LABELS[m]}
                </option>
              ))}
            </select>
            <ChevronDown className="pointer-events-none absolute right-2.5 top-1/2 -translate-y-1/2 h-4 w-4 text-somus-text-tertiary" />
          </div>
        </div>
      )}

      {/* ── Navigation ── */}
      <nav className="flex-1 overflow-y-auto py-2 px-2 space-y-0.5">
        {navItems.map((item) => {
          const isActive = currentPage === item.key;
          const icon = iconMap[item.icon] ?? <LayoutDashboard className={ICON_SIZE} />;

          return (
            <button
              key={item.key}
              onClick={() => setPage(item.key)}
              title={sidebarCollapsed ? item.label : undefined}
              className={cn(
                'flex items-center w-full rounded-lg text-sm font-medium transition-all duration-150 relative',
                sidebarCollapsed
                  ? 'justify-center px-0 py-2.5'
                  : 'gap-3 px-3 py-2',
                isActive
                  ? 'text-somus-text-accent'
                  : 'text-somus-text-secondary hover:bg-somus-bg-hover hover:text-somus-text-primary'
              )}
              style={isActive ? {
                background: 'linear-gradient(90deg, rgba(26,122,62,0.15) 0%, rgba(26,122,62,0.03) 100%)',
                borderLeft: '2px solid #1A7A3E',
              } : undefined}
            >
              <span className="shrink-0">{icon}</span>
              {!sidebarCollapsed && <span className="truncate">{item.label}</span>}
            </button>
          );
        })}
      </nav>

      {/* ── User Section ── */}
      <div className="border-t border-somus-border px-3 py-3">
        {!sidebarCollapsed && (
          <div className="flex items-center gap-2 mb-2">
            <div className="h-7 w-7 rounded-full bg-somus-bg-tertiary border border-somus-border flex items-center justify-center text-xs font-bold text-somus-text-accent shrink-0">
              U
            </div>
            <div className="flex-1 min-w-0">
              <p className="text-xs font-medium text-somus-text-primary truncate">Somus Capital</p>
              <p className="text-[10px] text-somus-text-tertiary">Mesa de Produtos</p>
            </div>
          </div>
        )}

        {/* Version badge */}
        {!sidebarCollapsed && (
          <div className="mb-2">
            <span className="inline-flex items-center px-2 py-0.5 rounded text-[10px] font-mono text-somus-text-tertiary bg-somus-bg-tertiary border border-somus-border">
              v2.0.0
            </span>
          </div>
        )}
      </div>

      {/* ── Collapse Toggle ── */}
      <div className="border-t border-somus-border p-2">
        <button
          onClick={toggleSidebar}
          className="flex items-center justify-center w-full py-2 rounded-lg text-somus-text-tertiary hover:text-somus-text-primary hover:bg-somus-bg-hover transition-colors"
          title={sidebarCollapsed ? 'Expandir menu' : 'Recolher menu'}
        >
          {sidebarCollapsed ? (
            <PanelLeftOpen className="h-5 w-5" />
          ) : (
            <PanelLeftClose className="h-5 w-5" />
          )}
        </button>
      </div>
    </aside>
  );
}

export default Sidebar;

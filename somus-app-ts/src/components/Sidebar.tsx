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
        'flex flex-col h-full bg-somus-green text-white transition-all duration-300 ease-in-out shrink-0',
        sidebarCollapsed ? 'w-16' : 'w-60'
      )}
    >
      {/* ── Logo ── */}
      <div className="flex items-center justify-center h-14 px-3 border-b border-white/10">
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
              className="w-full appearance-none bg-white/10 border border-white/20 text-white text-sm font-medium rounded-lg px-3 py-2 pr-8 focus:outline-none focus:ring-2 focus:ring-white/30 cursor-pointer"
            >
              {availableModules.map((m) => (
                <option
                  key={m}
                  value={m}
                  className="bg-somus-green text-white"
                >
                  {MODULE_LABELS[m]}
                </option>
              ))}
            </select>
            <ChevronDown className="pointer-events-none absolute right-2.5 top-1/2 -translate-y-1/2 h-4 w-4 text-white/60" />
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
                'flex items-center w-full rounded-lg text-sm font-medium transition-colors duration-150',
                sidebarCollapsed
                  ? 'justify-center px-0 py-2.5'
                  : 'gap-3 px-3 py-2',
                isActive
                  ? 'bg-white/20 text-white'
                  : 'text-white/70 hover:bg-white/10 hover:text-white'
              )}
            >
              <span className="shrink-0">{icon}</span>
              {!sidebarCollapsed && <span className="truncate">{item.label}</span>}
            </button>
          );
        })}
      </nav>

      {/* ── Collapse Toggle ── */}
      <div className="border-t border-white/10 p-2">
        <button
          onClick={toggleSidebar}
          className="flex items-center justify-center w-full py-2 rounded-lg text-white/60 hover:text-white hover:bg-white/10 transition-colors"
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

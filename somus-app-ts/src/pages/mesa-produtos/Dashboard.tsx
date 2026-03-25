import React, { useState, useMemo } from 'react';
import { Card } from '@components/Card';
import { Button } from '@components/Button';
import {
  DollarSign,
  TrendingUp,
  BarChart3,
  Users,
  Filter,
  Download,
} from 'lucide-react';
import { cn, formatCurrency } from '@/lib/utils';
import {
  BarChart,
  Bar,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  ResponsiveContainer,
  PieChart,
  Pie,
  Cell,
  LineChart,
  Line,
  Legend,
} from 'recharts';

// ── Mock Data ────────────────────────────────────────────────────────────────

const EQUIPES = ['SP', 'LEBLON', 'PRODUTOS', 'CORPORATE', 'BACKOFFICE'] as const;

const MESES = [
  'Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun',
  'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez',
];

const PRODUTOS = [
  'Renda Fixa', 'Renda Variavel', 'Fundos', 'COE',
  'Previdencia', 'Seguros', 'Cambio', 'Consorcio',
];

const COLORS_CHART = [
  '#004D33', '#005C3D', '#8CCFB4', '#0D9488',
  '#F59E0B', '#EF4444', '#8B5CF6', '#EC4899',
];

function generateMockReceitaByEquipe() {
  return EQUIPES.map((equipe) => ({
    equipe,
    receita: Math.round(Math.random() * 2000000 + 500000),
  }));
}

function generateMockTopAssessors() {
  const nomes = [
    'Carlos Silva', 'Ana Santos', 'Pedro Costa', 'Maria Oliveira',
    'Joao Souza', 'Fernanda Lima', 'Lucas Almeida', 'Juliana Rocha',
    'Bruno Ferreira', 'Camila Pereira',
  ];
  return nomes.map((nome, i) => ({
    rank: i + 1,
    nome,
    equipe: EQUIPES[i % EQUIPES.length],
    receita: Math.round((10 - i) * 150000 + Math.random() * 80000),
    operacoes: Math.round(Math.random() * 50 + 10),
  }));
}

function generateMockReceitaPorProduto() {
  return PRODUTOS.map((produto, i) => ({
    name: produto,
    value: Math.round(Math.random() * 800000 + 100000),
    color: COLORS_CHART[i],
  }));
}

function generateMockEvolucaoMensal() {
  return MESES.map((mes) => ({
    mes,
    receita: Math.round(Math.random() * 3000000 + 1000000),
    meta: 2500000,
  }));
}

// ── Component ────────────────────────────────────────────────────────────────

export default function Dashboard() {
  const [mesFilter, setMesFilter] = useState(new Date().getMonth());
  const [anoFilter, setAnoFilter] = useState(new Date().getFullYear());
  const [equipeFilter, setEquipeFilter] = useState<string>('TODAS');
  const [produtoFilter, setProdutoFilter] = useState<string>('TODOS');

  const receitaByEquipe = useMemo(generateMockReceitaByEquipe, [mesFilter, anoFilter]);
  const topAssessors = useMemo(generateMockTopAssessors, [mesFilter, anoFilter, equipeFilter]);
  const receitaPorProduto = useMemo(generateMockReceitaPorProduto, [mesFilter, anoFilter]);
  const evolucaoMensal = useMemo(generateMockEvolucaoMensal, [anoFilter]);

  const receitaTotal = receitaByEquipe.reduce((s, e) => s + e.receita, 0);
  const receitaMes = Math.round(receitaTotal * 0.12);
  const qtdOperacoes = topAssessors.reduce((s, a) => s + a.operacoes, 0);
  const ticketMedio = qtdOperacoes > 0 ? Math.round(receitaMes / qtdOperacoes) : 0;

  const anos = Array.from({ length: 5 }, (_, i) => new Date().getFullYear() - i);

  return (
    <div className="space-y-6">
      {/* Header */}
      <div className="flex items-center justify-between">
        <div>
          <h1 className="text-2xl font-bold text-somus-text-primary">
            Dashboard de Receita
          </h1>
          <p className="text-sm text-somus-text-secondary mt-1">
            Visao geral da receita Mesa de Produtos
          </p>
        </div>
        <Button variant="secondary" icon={<Download className="h-4 w-4" />}>
          Exportar
        </Button>
      </div>

      {/* Filters */}
      <Card padding="sm">
        <div className="flex flex-wrap items-center gap-4 p-2">
          <div className="flex items-center gap-2">
            <Filter className="h-4 w-4 text-somus-text-tertiary" />
            <span className="text-sm font-medium text-somus-text-secondary">
              Filtros:
            </span>
          </div>

          <select
            value={mesFilter}
            onChange={(e) => setMesFilter(Number(e.target.value))}
            className="text-sm border border-somus-border rounded-lg px-3 py-1.5 bg-somus-bg-input text-somus-text-primary focus:ring-2 focus:ring-somus-green/40 focus:outline-none"
          >
            {MESES.map((m, i) => (
              <option key={i} value={i}>{m}</option>
            ))}
          </select>

          <select
            value={anoFilter}
            onChange={(e) => setAnoFilter(Number(e.target.value))}
            className="text-sm border border-somus-border rounded-lg px-3 py-1.5 bg-somus-bg-input text-somus-text-primary focus:ring-2 focus:ring-somus-green/40 focus:outline-none"
          >
            {anos.map((a) => (
              <option key={a} value={a}>{a}</option>
            ))}
          </select>

          <select
            value={equipeFilter}
            onChange={(e) => setEquipeFilter(e.target.value)}
            className="text-sm border border-somus-border rounded-lg px-3 py-1.5 bg-somus-bg-input text-somus-text-primary focus:ring-2 focus:ring-somus-green/40 focus:outline-none"
          >
            <option value="TODAS">Todas as equipes</option>
            {EQUIPES.map((eq) => (
              <option key={eq} value={eq}>{eq}</option>
            ))}
          </select>

          <select
            value={produtoFilter}
            onChange={(e) => setProdutoFilter(e.target.value)}
            className="text-sm border border-somus-border rounded-lg px-3 py-1.5 bg-somus-bg-input text-somus-text-primary focus:ring-2 focus:ring-somus-green/40 focus:outline-none"
          >
            <option value="TODOS">Todos os produtos</option>
            {PRODUTOS.map((p) => (
              <option key={p} value={p}>{p}</option>
            ))}
          </select>
        </div>
      </Card>

      {/* KPI Cards */}
      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
        <KPICard
          title="Receita Total"
          value={formatCurrency(receitaTotal)}
          icon={<DollarSign className="h-5 w-5" />}
          trend="+12,5%"
          trendUp
        />
        <KPICard
          title="Receita Mes"
          value={formatCurrency(receitaMes)}
          icon={<TrendingUp className="h-5 w-5" />}
          trend="+8,3%"
          trendUp
        />
        <KPICard
          title="Quantidade Operacoes"
          value={qtdOperacoes.toString()}
          icon={<BarChart3 className="h-5 w-5" />}
          trend="+15"
          trendUp
        />
        <KPICard
          title="Ticket Medio"
          value={formatCurrency(ticketMedio)}
          icon={<Users className="h-5 w-5" />}
          trend="-2,1%"
          trendUp={false}
        />
      </div>

      {/* Charts Row 1 */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        {/* Receita por Equipe */}
        <Card title="Receita por Equipe">
          <div className="h-72 mt-4">
            <ResponsiveContainer width="100%" height="100%">
              <BarChart data={receitaByEquipe}>
                <CartesianGrid strokeDasharray="3 3" stroke="#1E2A3A" />
                <XAxis dataKey="equipe" fontSize={12} tick={{ fill: '#8B95A5' }} />
                <YAxis
                  fontSize={12}
                  tick={{ fill: '#8B95A5' }}
                  tickFormatter={(v) => `${(v / 1000000).toFixed(1)}M`}
                />
                <Tooltip
                  formatter={(value: number) => [formatCurrency(value), 'Receita']}
                  contentStyle={{
                    borderRadius: '8px',
                    backgroundColor: '#0F1419',
                    border: '1px solid #1E2A3A',
                    color: '#E8ECF0',
                    fontSize: '12px',
                  }}
                />
                <Bar dataKey="receita" fill="#004D33" radius={[4, 4, 0, 0]} />
              </BarChart>
            </ResponsiveContainer>
          </div>
        </Card>

        {/* Receita por Produto */}
        <Card title="Receita por Produto">
          <div className="h-72 mt-4">
            <ResponsiveContainer width="100%" height="100%">
              <PieChart>
                <Pie
                  data={receitaPorProduto}
                  cx="50%"
                  cy="50%"
                  innerRadius={60}
                  outerRadius={100}
                  dataKey="value"
                  nameKey="name"
                  paddingAngle={2}
                >
                  {receitaPorProduto.map((entry, index) => (
                    <Cell key={index} fill={entry.color} />
                  ))}
                </Pie>
                <Tooltip
                  formatter={(value: number) => [formatCurrency(value), 'Receita']}
                  contentStyle={{
                    borderRadius: '8px',
                    backgroundColor: '#0F1419',
                    border: '1px solid #1E2A3A',
                    color: '#E8ECF0',
                    fontSize: '12px',
                  }}
                />
                <Legend
                  iconType="circle"
                  iconSize={8}
                  wrapperStyle={{ fontSize: '11px', color: '#8B95A5' }}
                />
              </PieChart>
            </ResponsiveContainer>
          </div>
        </Card>
      </div>

      {/* Evolucao Mensal */}
      <Card title="Evolucao Mensal" subtitle={`Receita ${anoFilter}`}>
        <div className="h-72 mt-4">
          <ResponsiveContainer width="100%" height="100%">
            <LineChart data={evolucaoMensal}>
              <CartesianGrid strokeDasharray="3 3" stroke="#1E2A3A" />
              <XAxis dataKey="mes" fontSize={12} tick={{ fill: '#8B95A5' }} />
              <YAxis
                fontSize={12}
                tick={{ fill: '#8B95A5' }}
                tickFormatter={(v) => `${(v / 1000000).toFixed(1)}M`}
              />
              <Tooltip
                formatter={(value: number) => [formatCurrency(value)]}
                contentStyle={{
                  borderRadius: '8px',
                  backgroundColor: '#0F1419',
                  border: '1px solid #1E2A3A',
                  color: '#E8ECF0',
                  fontSize: '12px',
                }}
              />
              <Legend iconType="line" wrapperStyle={{ fontSize: '12px', color: '#8B95A5' }} />
              <Line
                type="monotone"
                dataKey="receita"
                stroke="#004D33"
                strokeWidth={2}
                dot={{ fill: '#004D33', r: 4 }}
                name="Receita"
              />
              <Line
                type="monotone"
                dataKey="meta"
                stroke="#F59E0B"
                strokeWidth={2}
                strokeDasharray="5 5"
                dot={false}
                name="Meta"
              />
            </LineChart>
          </ResponsiveContainer>
        </div>
      </Card>

      {/* Top 10 Assessors */}
      <Card title="Top 10 Assessores" subtitle="Ranking por receita no periodo">
        <div className="mt-4 overflow-x-auto">
          <table className="w-full text-sm">
            <thead>
              <tr className="border-b border-somus-border">
                <th className="text-left py-3 px-4 font-semibold text-somus-text-secondary">#</th>
                <th className="text-left py-3 px-4 font-semibold text-somus-text-secondary">Assessor</th>
                <th className="text-left py-3 px-4 font-semibold text-somus-text-secondary">Equipe</th>
                <th className="text-right py-3 px-4 font-semibold text-somus-text-secondary">Receita</th>
                <th className="text-right py-3 px-4 font-semibold text-somus-text-secondary">Operacoes</th>
              </tr>
            </thead>
            <tbody>
              {topAssessors.map((a) => (
                <tr
                  key={a.rank}
                  className="border-b border-somus-border/30 hover:bg-somus-bg-hover transition-colors"
                >
                  <td className="py-3 px-4">
                    <span
                      className={cn(
                        'inline-flex items-center justify-center w-6 h-6 rounded-full text-xs font-bold',
                        a.rank <= 3
                          ? 'bg-somus-green text-white'
                          : 'bg-somus-bg-tertiary text-somus-text-secondary'
                      )}
                    >
                      {a.rank}
                    </span>
                  </td>
                  <td className="py-3 px-4 font-medium text-somus-text-primary">
                    {a.nome}
                  </td>
                  <td className="py-3 px-4">
                    <span className="inline-block px-2 py-0.5 text-xs rounded-full bg-somus-green/10 text-somus-green font-medium">
                      {a.equipe}
                    </span>
                  </td>
                  <td className="py-3 px-4 text-right font-medium text-somus-text-primary">
                    {formatCurrency(a.receita)}
                  </td>
                  <td className="py-3 px-4 text-right text-somus-text-secondary">
                    {a.operacoes}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </Card>
    </div>
  );
}

// ── KPI Card ─────────────────────────────────────────────────────────────────

function KPICard({
  title,
  value,
  icon,
  trend,
  trendUp,
}: {
  title: string;
  value: string;
  icon: React.ReactNode;
  trend: string;
  trendUp: boolean;
}) {
  return (
    <div className="bg-somus-bg-secondary rounded-lg border border-somus-border p-5">
      <div className="flex items-center justify-between mb-3">
        <span className="text-sm text-somus-text-secondary font-medium">{title}</span>
        <div className="p-2 rounded-lg bg-somus-green/10 text-somus-green">
          {icon}
        </div>
      </div>
      <div className="text-2xl font-bold text-somus-text-primary">{value}</div>
      <div className="mt-1 flex items-center gap-1">
        <span
          className={cn(
            'text-xs font-medium',
            trendUp ? 'text-somus-green-400' : 'text-red-400'
          )}
        >
          {trendUp ? '\u2191' : '\u2193'} {trend}
        </span>
        <span className="text-xs text-somus-text-tertiary">vs mes anterior</span>
      </div>
    </div>
  );
}

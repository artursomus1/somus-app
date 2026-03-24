import React from 'react';
import {
  Calculator,
  TrendingUp,
  Layers,
  FileText,
  DollarSign,
  FolderOpen,
} from 'lucide-react';
import { PageLayout } from '@components/PageLayout';
import { Card } from '@components/Card';
import { Button } from '@components/Button';
import { useAppStore } from '@/stores/appStore';

interface ToolCard {
  key: string;
  title: string;
  description: string;
  icon: React.ReactNode;
  color: string;
  bgColor: string;
}

const tools: ToolCard[] = [
  {
    key: 'simulador',
    title: 'Simulador de Consorcio',
    description: 'Simule operacoes de consorcio',
    icon: <Calculator className="h-6 w-6" />,
    color: 'text-emerald-600',
    bgColor: 'bg-emerald-50',
  },
  {
    key: 'comparativo-vpl',
    title: 'Comparativo de VPL',
    description: 'Analise VPL Goal-Based (NASA HD)',
    icon: <TrendingUp className="h-6 w-6" />,
    color: 'text-blue-600',
    bgColor: 'bg-blue-50',
  },
  {
    key: 'consorcio-vs-financ',
    title: 'Consorcio vs Financiamento',
    description: 'Compare lado a lado',
    icon: <Layers className="h-6 w-6" />,
    color: 'text-violet-600',
    bgColor: 'bg-violet-50',
  },
  {
    key: 'gerador-propostas',
    title: 'Gerador de Propostas',
    description: 'Gere apresentacoes PPTX',
    icon: <FileText className="h-6 w-6" />,
    color: 'text-orange-600',
    bgColor: 'bg-orange-50',
  },
  {
    key: 'fluxo-receitas',
    title: 'Fluxo de Receitas',
    description: 'Acompanhe pagamentos e receitas',
    icon: <DollarSign className="h-6 w-6" />,
    color: 'text-teal-600',
    bgColor: 'bg-teal-50',
  },
  {
    key: 'cenarios',
    title: 'Cenarios',
    description: 'Gerencie ate 10 cenarios',
    icon: <FolderOpen className="h-6 w-6" />,
    color: 'text-indigo-600',
    bgColor: 'bg-indigo-50',
  },
];

export default function CorporateDashboard() {
  const setPage = useAppStore((s) => s.setPage);

  return (
    <PageLayout title="Corporate" subtitle="Ferramentas de analise e simulacao para consorcios">
      <div className="max-w-6xl mx-auto">
        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-5">
          {tools.map((tool) => (
            <Card key={tool.key} padding="lg" className="hover:shadow-md transition-shadow">
              <div className="flex flex-col h-full">
                <div
                  className={`inline-flex items-center justify-center h-12 w-12 rounded-lg ${tool.bgColor} ${tool.color} mb-4`}
                >
                  {tool.icon}
                </div>
                <h3 className="text-base font-semibold text-somus-gray-900 mb-1">
                  {tool.title}
                </h3>
                <p className="text-sm text-somus-gray-500 mb-5 flex-1">
                  {tool.description}
                </p>
                <Button
                  variant="primary"
                  size="sm"
                  fullWidth
                  onClick={() => setPage(tool.key)}
                >
                  Acessar
                </Button>
              </div>
            </Card>
          ))}
        </div>
      </div>
    </PageLayout>
  );
}

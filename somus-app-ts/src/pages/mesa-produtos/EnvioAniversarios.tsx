import React, { useState, useMemo, useCallback } from 'react';
import { Card } from '@components/Card';
import { Button } from '@components/Button';
import {
  Cake,
  Mail,
  Send,
  Calendar,
  Search,
  Check,
  Edit3,
  ChevronLeft,
  ChevronRight,
} from 'lucide-react';
import { cn } from '@/lib/utils';
import { createDrafts } from '@/services/outlook';

// ── Types ────────────────────────────────────────────────────────────────────

interface Aniversariante {
  id: string;
  cliente: string;
  dataNascimento: string; // YYYY-MM-DD
  assessor: string;
  email: string;
  telefone: string;
}

// ── Mock Data ────────────────────────────────────────────────────────────────

const today = new Date();
const currentMonth = today.getMonth();
const currentYear = today.getFullYear();

function mockDate(day: number, monthOffset: number = 0) {
  const m = currentMonth + monthOffset;
  const y = currentYear - Math.floor(Math.random() * 30 + 25);
  return `${y}-${String(m + 1).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
}

const MOCK_ANIVERSARIANTES: Aniversariante[] = [
  { id: '1', cliente: 'Joao Mendes', dataNascimento: mockDate(5), assessor: 'Carlos Silva', email: 'joao@email.com', telefone: '(11) 99999-1111' },
  { id: '2', cliente: 'Maria Costa', dataNascimento: mockDate(12), assessor: 'Carlos Silva', email: 'maria@email.com', telefone: '(11) 99999-2222' },
  { id: '3', cliente: 'Pedro Lima', dataNascimento: mockDate(18), assessor: 'Ana Santos', email: 'pedro@email.com', telefone: '(21) 99999-3333' },
  { id: '4', cliente: 'Ana Souza', dataNascimento: mockDate(24), assessor: 'Pedro Costa', email: 'ana@email.com', telefone: '(11) 99999-4444' },
  { id: '5', cliente: 'Lucas Ferreira', dataNascimento: mockDate(28), assessor: 'Maria Oliveira', email: 'lucas@email.com', telefone: '(21) 99999-5555' },
  { id: '6', cliente: 'Fernanda Dias', dataNascimento: mockDate(3, 1), assessor: 'Carlos Silva', email: 'fernanda@email.com', telefone: '(11) 99999-6666' },
  { id: '7', cliente: 'Bruno Neves', dataNascimento: mockDate(10, 1), assessor: 'Ana Santos', email: 'bruno@email.com', telefone: '(21) 99999-7777' },
  { id: '8', cliente: 'Camila Rocha', dataNascimento: mockDate(22, 1), assessor: 'Pedro Costa', email: 'camila@email.com', telefone: '(11) 99999-8888' },
  { id: '9', cliente: 'Ricardo Gomes', dataNascimento: mockDate(15), assessor: 'Joao Souza', email: 'ricardo@email.com', telefone: '(21) 99999-9999' },
  { id: '10', cliente: 'Tatiana Mello', dataNascimento: mockDate(7, -1), assessor: 'Fernanda Lima', email: 'tatiana@email.com', telefone: '(11) 99999-0000' },
];

const MESES = [
  'Janeiro', 'Fevereiro', 'Marco', 'Abril', 'Maio', 'Junho',
  'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro',
];

const DEFAULT_TEMPLATE = `Prezado(a) {nome},

Gostaríamos de desejar a voce um Feliz Aniversario!

Que este novo ciclo seja repleto de conquistas, saude e prosperidade.

A equipe Somus Capital esta sempre à disposicao para auxiliar em seus investimentos e planejamento financeiro.

Um grande abraco,
Equipe Somus Capital`;

// ── Component ────────────────────────────────────────────────────────────────

export default function EnvioAniversarios() {
  const [aniversariantes] = useState<Aniversariante[]>(MOCK_ANIVERSARIANTES);
  const [calendarMonth, setCalendarMonth] = useState(currentMonth);
  const [calendarYear, setCalendarYear] = useState(currentYear);
  const [searchTerm, setSearchTerm] = useState('');
  const [messageTemplate, setMessageTemplate] = useState(DEFAULT_TEMPLATE);
  const [editingTemplate, setEditingTemplate] = useState(false);
  const [generating, setGenerating] = useState(false);
  const [generated, setGenerated] = useState(false);

  // Filter aniversariantes for the selected month
  const monthAniversariantes = useMemo(() => {
    return aniversariantes
      .filter((a) => {
        const birthMonth = new Date(a.dataNascimento + 'T00:00:00').getMonth();
        return birthMonth === calendarMonth;
      })
      .sort((a, b) => {
        const dayA = new Date(a.dataNascimento + 'T00:00:00').getDate();
        const dayB = new Date(b.dataNascimento + 'T00:00:00').getDate();
        return dayA - dayB;
      });
  }, [aniversariantes, calendarMonth]);

  const filteredAniversariantes = useMemo(() => {
    if (!searchTerm) return monthAniversariantes;
    const term = searchTerm.toLowerCase();
    return monthAniversariantes.filter(
      (a) =>
        a.cliente.toLowerCase().includes(term) ||
        a.assessor.toLowerCase().includes(term)
    );
  }, [monthAniversariantes, searchTerm]);

  // Calendar data
  const calendarDays = useMemo(() => {
    const firstDay = new Date(calendarYear, calendarMonth, 1).getDay();
    const daysInMonth = new Date(calendarYear, calendarMonth + 1, 0).getDate();
    const days: Array<{ day: number; hasAniversario: boolean }> = [];

    // Empty cells before first day
    for (let i = 0; i < firstDay; i++) {
      days.push({ day: 0, hasAniversario: false });
    }

    for (let d = 1; d <= daysInMonth; d++) {
      const hasAniversario = aniversariantes.some((a) => {
        const birth = new Date(a.dataNascimento + 'T00:00:00');
        return birth.getMonth() === calendarMonth && birth.getDate() === d;
      });
      days.push({ day: d, hasAniversario });
    }

    return days;
  }, [aniversariantes, calendarMonth, calendarYear]);

  const prevMonth = () => {
    if (calendarMonth === 0) {
      setCalendarMonth(11);
      setCalendarYear((y) => y - 1);
    } else {
      setCalendarMonth((m) => m - 1);
    }
  };

  const nextMonth = () => {
    if (calendarMonth === 11) {
      setCalendarMonth(0);
      setCalendarYear((y) => y + 1);
    } else {
      setCalendarMonth((m) => m + 1);
    }
  };

  const getAge = (dataNascimento: string) => {
    const birth = new Date(dataNascimento + 'T00:00:00');
    const thisYearBirthday = new Date(calendarYear, birth.getMonth(), birth.getDate());
    let age = calendarYear - birth.getFullYear();
    if (thisYearBirthday > today) age--;
    return age + 1; // age they will turn
  };

  const handleGerarRascunhos = useCallback(async () => {
    setGenerating(true);
    try {
      const emails = filteredAniversariantes.map((a) => ({
        to: a.email,
        subject: `Feliz Aniversario, ${a.cliente.split(' ')[0]}! - Somus Capital`,
        body: `<div style="font-family: DM Sans, sans-serif; max-width: 600px;">
          <div style="background: #004D33; color: white; padding: 20px; border-radius: 8px 8px 0 0; text-align: center;">
            <h2 style="margin: 0; font-size: 24px;">Feliz Aniversario!</h2>
          </div>
          <div style="padding: 25px; border: 1px solid #E5E7EB; border-top: none; border-radius: 0 0 8px 8px;">
            <p style="font-size: 14px; color: #111; line-height: 1.6; white-space: pre-line;">${messageTemplate.replace('{nome}', a.cliente.split(' ')[0])}</p>
          </div>
        </div>`,
      }));

      await createDrafts(emails);
      setGenerated(true);
      setTimeout(() => setGenerated(false), 3000);
    } catch (err) {
      console.error('Erro ao gerar rascunhos:', err);
    } finally {
      setGenerating(false);
    }
  }, [filteredAniversariantes, messageTemplate]);

  return (
    <div className="space-y-6">
      {/* Header */}
      <div className="flex items-center justify-between">
        <div>
          <h1 className="text-2xl font-bold text-somus-gray-900">Envio de Aniversarios</h1>
          <p className="text-sm text-somus-gray-500 mt-1">
            Envie mensagens de aniversario automaticamente
          </p>
        </div>
        <Button
          variant="primary"
          icon={<Mail className="h-4 w-4" />}
          loading={generating}
          onClick={handleGerarRascunhos}
          disabled={filteredAniversariantes.length === 0}
        >
          {generated ? 'Rascunhos Criados!' : `Gerar Rascunhos (${filteredAniversariantes.length})`}
        </Button>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
        {/* Calendar */}
        <Card title="Calendario">
          <div className="mt-4">
            {/* Month navigation */}
            <div className="flex items-center justify-between mb-4">
              <button onClick={prevMonth} className="p-1 hover:bg-somus-gray-100 rounded transition-colors">
                <ChevronLeft className="h-5 w-5 text-somus-gray-600" />
              </button>
              <span className="text-sm font-semibold text-somus-gray-900">
                {MESES[calendarMonth]} {calendarYear}
              </span>
              <button onClick={nextMonth} className="p-1 hover:bg-somus-gray-100 rounded transition-colors">
                <ChevronRight className="h-5 w-5 text-somus-gray-600" />
              </button>
            </div>

            {/* Weekday headers */}
            <div className="grid grid-cols-7 gap-1 mb-1">
              {['Dom', 'Seg', 'Ter', 'Qua', 'Qui', 'Sex', 'Sab'].map((d) => (
                <div key={d} className="text-center text-xs font-medium text-somus-gray-400 py-1">
                  {d}
                </div>
              ))}
            </div>

            {/* Days */}
            <div className="grid grid-cols-7 gap-1">
              {calendarDays.map((d, i) => (
                <div
                  key={i}
                  className={cn(
                    'aspect-square flex items-center justify-center text-sm rounded-lg',
                    d.day === 0 && 'invisible',
                    d.day === today.getDate() &&
                      calendarMonth === currentMonth &&
                      calendarYear === currentYear &&
                      'bg-somus-green text-white font-bold',
                    d.hasAniversario && !(d.day === today.getDate() && calendarMonth === currentMonth) &&
                      'bg-amber-100 text-amber-800 font-semibold',
                    !d.hasAniversario &&
                      !(d.day === today.getDate() && calendarMonth === currentMonth && calendarYear === currentYear) &&
                      'text-somus-gray-600 hover:bg-somus-gray-50'
                  )}
                >
                  {d.day > 0 && (
                    <span className="relative">
                      {d.day}
                      {d.hasAniversario && (
                        <Cake className="absolute -top-2 -right-3 h-3 w-3 text-amber-600" />
                      )}
                    </span>
                  )}
                </div>
              ))}
            </div>

            <div className="mt-4 flex items-center gap-4 text-xs text-somus-gray-500">
              <div className="flex items-center gap-1">
                <div className="w-3 h-3 rounded bg-amber-100 border border-amber-200" />
                Aniversario
              </div>
              <div className="flex items-center gap-1">
                <div className="w-3 h-3 rounded bg-somus-green" />
                Hoje
              </div>
            </div>
          </div>
        </Card>

        {/* Aniversariantes List & Template */}
        <div className="lg:col-span-2 space-y-6">
          {/* Search */}
          <div className="relative">
            <Search className="absolute left-3 top-1/2 -translate-y-1/2 h-4 w-4 text-somus-gray-400" />
            <input
              type="text"
              placeholder="Buscar cliente ou assessor..."
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
              className="w-full text-sm border border-somus-gray-300 rounded-lg pl-9 pr-3 py-2 focus:ring-2 focus:ring-somus-green/40 focus:outline-none"
            />
          </div>

          {/* Aniversariantes Table */}
          <Card
            title={`Aniversariantes - ${MESES[calendarMonth]}`}
            subtitle={`${filteredAniversariantes.length} clientes`}
          >
            <div className="mt-4 overflow-x-auto">
              <table className="w-full text-sm">
                <thead>
                  <tr className="border-b border-somus-gray-200">
                    <th className="text-left py-3 px-4 font-semibold text-somus-gray-600">Cliente</th>
                    <th className="text-left py-3 px-4 font-semibold text-somus-gray-600">Data</th>
                    <th className="text-left py-3 px-4 font-semibold text-somus-gray-600">Idade</th>
                    <th className="text-left py-3 px-4 font-semibold text-somus-gray-600">Assessor</th>
                    <th className="text-left py-3 px-4 font-semibold text-somus-gray-600">Contato</th>
                  </tr>
                </thead>
                <tbody>
                  {filteredAniversariantes.map((a) => {
                    const birthDate = new Date(a.dataNascimento + 'T00:00:00');
                    const isPast = birthDate.getDate() < today.getDate() && calendarMonth === currentMonth;
                    const isToday = birthDate.getDate() === today.getDate() && calendarMonth === currentMonth;
                    return (
                      <tr
                        key={a.id}
                        className={cn(
                          'border-b border-somus-gray-100 transition-colors',
                          isToday ? 'bg-amber-50' : 'hover:bg-somus-gray-50'
                        )}
                      >
                        <td className="py-3 px-4">
                          <div className="flex items-center gap-2">
                            {isToday && <Cake className="h-4 w-4 text-amber-500" />}
                            <span className="font-medium text-somus-gray-900">{a.cliente}</span>
                          </div>
                        </td>
                        <td className="py-3 px-4 text-somus-gray-600">
                          {birthDate.toLocaleDateString('pt-BR')}
                        </td>
                        <td className="py-3 px-4 text-somus-gray-600">
                          {getAge(a.dataNascimento)} anos
                        </td>
                        <td className="py-3 px-4 text-somus-gray-600">{a.assessor}</td>
                        <td className="py-3 px-4 text-somus-gray-500 text-xs">{a.telefone}</td>
                      </tr>
                    );
                  })}
                  {filteredAniversariantes.length === 0 && (
                    <tr>
                      <td colSpan={5} className="py-12 text-center text-somus-gray-400">
                        Nenhum aniversariante neste mes
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          </Card>

          {/* Message Template */}
          <Card
            title="Modelo de Mensagem"
            headerRight={
              <Button
                variant="ghost"
                size="sm"
                icon={<Edit3 className="h-4 w-4" />}
                onClick={() => setEditingTemplate(!editingTemplate)}
              >
                {editingTemplate ? 'Salvar' : 'Editar'}
              </Button>
            }
          >
            <div className="mt-4">
              {editingTemplate ? (
                <textarea
                  value={messageTemplate}
                  onChange={(e) => setMessageTemplate(e.target.value)}
                  rows={10}
                  className="w-full text-sm border border-somus-gray-300 rounded-lg px-3 py-2 focus:ring-2 focus:ring-somus-green/40 focus:outline-none resize-none"
                />
              ) : (
                <div className="bg-somus-gray-50 rounded-lg p-4 text-sm text-somus-gray-700 whitespace-pre-line">
                  {messageTemplate}
                </div>
              )}
              <p className="text-xs text-somus-gray-400 mt-2">
                Use {'{nome}'} para inserir o nome do cliente automaticamente
              </p>
            </div>
          </Card>
        </div>
      </div>
    </div>
  );
}

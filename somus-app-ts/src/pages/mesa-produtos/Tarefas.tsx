import React, { useState, useCallback, useRef } from 'react';
import { Card } from '@components/Card';
import { Button } from '@components/Button';
import {
  Plus,
  X,
  Clock,
  AlertCircle,
  CheckCircle2,
  GripVertical,
  Edit3,
  Trash2,
  Calendar,
  User,
  Flag,
} from 'lucide-react';
import { cn } from '@/lib/utils';
import { saveData, loadData } from '@/services/storage';

// ── Types ────────────────────────────────────────────────────────────────────

type Priority = 'baixa' | 'media' | 'alta' | 'urgente';
type Status = 'todo' | 'inprogress' | 'done';

interface Task {
  id: string;
  title: string;
  description: string;
  assignee: string;
  priority: Priority;
  dueDate: string;
  status: Status;
  createdAt: string;
}

const STORAGE_KEY = 'mesa-produtos-tarefas';

// ── Mock Data ────────────────────────────────────────────────────────────────

const ASSIGNEES = [
  'Carlos Silva', 'Ana Santos', 'Pedro Costa',
  'Maria Oliveira', 'Joao Souza', 'Fernanda Lima',
];

const DEFAULT_TASKS: Task[] = [
  { id: '1', title: 'Atualizar planilha de receita', description: 'Consolidar dados de receita do mes', assignee: 'Carlos Silva', priority: 'alta', dueDate: '2026-03-28', status: 'todo', createdAt: '2026-03-20' },
  { id: '2', title: 'Enviar informativo semanal', description: 'Preparar e enviar o informativo de mercado', assignee: 'Ana Santos', priority: 'media', dueDate: '2026-03-25', status: 'todo', createdAt: '2026-03-19' },
  { id: '3', title: 'Revisar ordens pendentes', description: 'Verificar ordens pendentes de execucao', assignee: 'Pedro Costa', priority: 'urgente', dueDate: '2026-03-24', status: 'inprogress', createdAt: '2026-03-18' },
  { id: '4', title: 'Gerar relatorio de agio', description: 'Extrair relatorio mensal de agio/desagio', assignee: 'Maria Oliveira', priority: 'media', dueDate: '2026-03-30', status: 'inprogress', createdAt: '2026-03-17' },
  { id: '5', title: 'Enviar saldos Q1', description: 'Envio trimestral de saldos completo', assignee: 'Joao Souza', priority: 'baixa', dueDate: '2026-03-31', status: 'done', createdAt: '2026-03-15' },
  { id: '6', title: 'Contatar clientes aniversariantes', description: 'Ligar para clientes com aniversario esta semana', assignee: 'Fernanda Lima', priority: 'baixa', dueDate: '2026-03-26', status: 'done', createdAt: '2026-03-16' },
];

const COLUMNS: { key: Status; label: string; icon: React.ReactNode; color: string }[] = [
  { key: 'todo', label: 'A Fazer', icon: <Clock className="h-4 w-4" />, color: 'text-somus-gray-500' },
  { key: 'inprogress', label: 'Em Progresso', icon: <AlertCircle className="h-4 w-4" />, color: 'text-blue-500' },
  { key: 'done', label: 'Concluido', icon: <CheckCircle2 className="h-4 w-4" />, color: 'text-emerald-500' },
];

const PRIORITY_STYLES: Record<Priority, { bg: string; text: string; label: string }> = {
  baixa: { bg: 'bg-somus-gray-100', text: 'text-somus-gray-600', label: 'Baixa' },
  media: { bg: 'bg-blue-50', text: 'text-blue-700', label: 'Media' },
  alta: { bg: 'bg-amber-50', text: 'text-amber-700', label: 'Alta' },
  urgente: { bg: 'bg-red-50', text: 'text-red-700', label: 'Urgente' },
};

// ── Component ────────────────────────────────────────────────────────────────

export default function Tarefas() {
  const [tasks, setTasks] = useState<Task[]>(() =>
    loadData<Task[]>(STORAGE_KEY, DEFAULT_TASKS)
  );
  const [showModal, setShowModal] = useState(false);
  const [editingTask, setEditingTask] = useState<Task | null>(null);
  const [filterAssignee, setFilterAssignee] = useState('TODOS');
  const [filterPriority, setFilterPriority] = useState<string>('TODAS');
  const [draggedTaskId, setDraggedTaskId] = useState<string | null>(null);
  const [dragOverColumn, setDragOverColumn] = useState<Status | null>(null);

  // Form state
  const [formTitle, setFormTitle] = useState('');
  const [formDescription, setFormDescription] = useState('');
  const [formAssignee, setFormAssignee] = useState(ASSIGNEES[0]);
  const [formPriority, setFormPriority] = useState<Priority>('media');
  const [formDueDate, setFormDueDate] = useState('');

  const persistTasks = useCallback((newTasks: Task[]) => {
    setTasks(newTasks);
    saveData(STORAGE_KEY, newTasks);
  }, []);

  const filteredTasks = tasks.filter((t) => {
    if (filterAssignee !== 'TODOS' && t.assignee !== filterAssignee) return false;
    if (filterPriority !== 'TODAS' && t.priority !== filterPriority) return false;
    return true;
  });

  const getColumnTasks = (status: Status) =>
    filteredTasks.filter((t) => t.status === status);

  // Modal handlers
  const openAddModal = () => {
    setEditingTask(null);
    setFormTitle('');
    setFormDescription('');
    setFormAssignee(ASSIGNEES[0]);
    setFormPriority('media');
    setFormDueDate('');
    setShowModal(true);
  };

  const openEditModal = (task: Task) => {
    setEditingTask(task);
    setFormTitle(task.title);
    setFormDescription(task.description);
    setFormAssignee(task.assignee);
    setFormPriority(task.priority);
    setFormDueDate(task.dueDate);
    setShowModal(true);
  };

  const handleSave = () => {
    if (!formTitle) return;

    if (editingTask) {
      const updated = tasks.map((t) =>
        t.id === editingTask.id
          ? {
              ...t,
              title: formTitle,
              description: formDescription,
              assignee: formAssignee,
              priority: formPriority,
              dueDate: formDueDate,
            }
          : t
      );
      persistTasks(updated);
    } else {
      const newTask: Task = {
        id: String(Date.now()),
        title: formTitle,
        description: formDescription,
        assignee: formAssignee,
        priority: formPriority,
        dueDate: formDueDate,
        status: 'todo',
        createdAt: new Date().toISOString().split('T')[0],
      };
      persistTasks([...tasks, newTask]);
    }
    setShowModal(false);
  };

  const deleteTask = (id: string) => {
    persistTasks(tasks.filter((t) => t.id !== id));
  };

  // Drag and drop
  const handleDragStart = (taskId: string) => {
    setDraggedTaskId(taskId);
  };

  const handleDragOver = (e: React.DragEvent, column: Status) => {
    e.preventDefault();
    setDragOverColumn(column);
  };

  const handleDragLeave = () => {
    setDragOverColumn(null);
  };

  const handleDrop = (e: React.DragEvent, column: Status) => {
    e.preventDefault();
    setDragOverColumn(null);
    if (!draggedTaskId) return;

    const updated = tasks.map((t) =>
      t.id === draggedTaskId ? { ...t, status: column } : t
    );
    persistTasks(updated);
    setDraggedTaskId(null);
  };

  const isOverdue = (dueDate: string) => {
    if (!dueDate) return false;
    return new Date(dueDate + 'T23:59:59') < new Date();
  };

  return (
    <div className="space-y-6">
      {/* Header */}
      <div className="flex items-center justify-between">
        <div>
          <h1 className="text-2xl font-bold text-somus-gray-900">Tarefas</h1>
          <p className="text-sm text-somus-gray-500 mt-1">
            Organizador de tarefas da equipe
          </p>
        </div>
        <Button variant="primary" icon={<Plus className="h-4 w-4" />} onClick={openAddModal}>
          Nova Tarefa
        </Button>
      </div>

      {/* Filters */}
      <Card padding="sm">
        <div className="flex flex-wrap items-center gap-4 p-2">
          <select
            value={filterAssignee}
            onChange={(e) => setFilterAssignee(e.target.value)}
            className="text-sm border border-somus-gray-300 rounded-lg px-3 py-1.5 focus:ring-2 focus:ring-somus-green/40 focus:outline-none"
          >
            <option value="TODOS">Todos responsaveis</option>
            {ASSIGNEES.map((a) => (
              <option key={a} value={a}>{a}</option>
            ))}
          </select>
          <select
            value={filterPriority}
            onChange={(e) => setFilterPriority(e.target.value)}
            className="text-sm border border-somus-gray-300 rounded-lg px-3 py-1.5 focus:ring-2 focus:ring-somus-green/40 focus:outline-none"
          >
            <option value="TODAS">Todas prioridades</option>
            <option value="urgente">Urgente</option>
            <option value="alta">Alta</option>
            <option value="media">Media</option>
            <option value="baixa">Baixa</option>
          </select>
          <span className="text-xs text-somus-gray-400">
            {filteredTasks.length} tarefas
          </span>
        </div>
      </Card>

      {/* Kanban Board */}
      <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
        {COLUMNS.map((col) => {
          const colTasks = getColumnTasks(col.key);
          return (
            <div
              key={col.key}
              onDragOver={(e) => handleDragOver(e, col.key)}
              onDragLeave={handleDragLeave}
              onDrop={(e) => handleDrop(e, col.key)}
              className={cn(
                'bg-somus-gray-50 rounded-xl p-4 min-h-[300px] transition-colors',
                dragOverColumn === col.key && 'bg-somus-green/5 ring-2 ring-somus-green/20'
              )}
            >
              {/* Column Header */}
              <div className="flex items-center justify-between mb-4">
                <div className="flex items-center gap-2">
                  <span className={col.color}>{col.icon}</span>
                  <span className="text-sm font-semibold text-somus-gray-700">{col.label}</span>
                  <span className="text-xs bg-somus-gray-200 text-somus-gray-600 rounded-full px-2 py-0.5">
                    {colTasks.length}
                  </span>
                </div>
              </div>

              {/* Tasks */}
              <div className="space-y-3">
                {colTasks.map((task) => (
                  <div
                    key={task.id}
                    draggable
                    onDragStart={() => handleDragStart(task.id)}
                    className={cn(
                      'bg-white rounded-lg border border-somus-gray-200 p-4 shadow-sm cursor-grab active:cursor-grabbing hover:shadow-md transition-shadow',
                      draggedTaskId === task.id && 'opacity-50'
                    )}
                  >
                    {/* Task header */}
                    <div className="flex items-start justify-between mb-2">
                      <h4 className="text-sm font-semibold text-somus-gray-900 flex-1">
                        {task.title}
                      </h4>
                      <div className="flex items-center gap-1 ml-2">
                        <button
                          onClick={() => openEditModal(task)}
                          className="text-somus-gray-400 hover:text-somus-green transition-colors p-0.5"
                        >
                          <Edit3 className="h-3.5 w-3.5" />
                        </button>
                        <button
                          onClick={() => deleteTask(task.id)}
                          className="text-somus-gray-400 hover:text-red-500 transition-colors p-0.5"
                        >
                          <Trash2 className="h-3.5 w-3.5" />
                        </button>
                      </div>
                    </div>

                    {task.description && (
                      <p className="text-xs text-somus-gray-500 mb-3 line-clamp-2">
                        {task.description}
                      </p>
                    )}

                    {/* Task footer */}
                    <div className="flex items-center justify-between">
                      <span
                        className={cn(
                          'inline-flex items-center gap-1 px-2 py-0.5 text-xs rounded-full font-medium',
                          PRIORITY_STYLES[task.priority].bg,
                          PRIORITY_STYLES[task.priority].text
                        )}
                      >
                        <Flag className="h-3 w-3" />
                        {PRIORITY_STYLES[task.priority].label}
                      </span>
                      <div className="flex items-center gap-2">
                        {task.dueDate && (
                          <span
                            className={cn(
                              'text-xs flex items-center gap-1',
                              isOverdue(task.dueDate) && task.status !== 'done'
                                ? 'text-red-500 font-medium'
                                : 'text-somus-gray-400'
                            )}
                          >
                            <Calendar className="h-3 w-3" />
                            {new Date(task.dueDate + 'T00:00:00').toLocaleDateString('pt-BR')}
                          </span>
                        )}
                      </div>
                    </div>

                    {/* Assignee */}
                    <div className="mt-2 flex items-center gap-1.5">
                      <div className="w-5 h-5 rounded-full bg-somus-green/20 flex items-center justify-center">
                        <User className="h-3 w-3 text-somus-green" />
                      </div>
                      <span className="text-xs text-somus-gray-500">{task.assignee}</span>
                    </div>
                  </div>
                ))}

                {colTasks.length === 0 && (
                  <div className="text-center py-8 text-xs text-somus-gray-400">
                    Arraste tarefas aqui
                  </div>
                )}
              </div>
            </div>
          );
        })}
      </div>

      {/* Modal */}
      {showModal && (
        <div className="fixed inset-0 bg-black/40 flex items-center justify-center z-50">
          <div className="bg-white rounded-xl shadow-xl w-full max-w-md mx-4">
            {/* Modal Header */}
            <div className="flex items-center justify-between px-6 py-4 border-b border-somus-gray-200">
              <h3 className="text-lg font-semibold text-somus-gray-900">
                {editingTask ? 'Editar Tarefa' : 'Nova Tarefa'}
              </h3>
              <button
                onClick={() => setShowModal(false)}
                className="text-somus-gray-400 hover:text-somus-gray-600 transition-colors"
              >
                <X className="h-5 w-5" />
              </button>
            </div>

            {/* Modal Body */}
            <div className="px-6 py-4 space-y-4">
              <div>
                <label className="block text-sm font-medium text-somus-gray-700 mb-1">Titulo *</label>
                <input
                  type="text"
                  value={formTitle}
                  onChange={(e) => setFormTitle(e.target.value)}
                  placeholder="Titulo da tarefa"
                  className="w-full text-sm border border-somus-gray-300 rounded-lg px-3 py-2 focus:ring-2 focus:ring-somus-green/40 focus:outline-none"
                />
              </div>

              <div>
                <label className="block text-sm font-medium text-somus-gray-700 mb-1">Descricao</label>
                <textarea
                  value={formDescription}
                  onChange={(e) => setFormDescription(e.target.value)}
                  placeholder="Descricao da tarefa..."
                  rows={3}
                  className="w-full text-sm border border-somus-gray-300 rounded-lg px-3 py-2 focus:ring-2 focus:ring-somus-green/40 focus:outline-none resize-none"
                />
              </div>

              <div className="grid grid-cols-2 gap-4">
                <div>
                  <label className="block text-sm font-medium text-somus-gray-700 mb-1">Responsavel</label>
                  <select
                    value={formAssignee}
                    onChange={(e) => setFormAssignee(e.target.value)}
                    className="w-full text-sm border border-somus-gray-300 rounded-lg px-3 py-2 focus:ring-2 focus:ring-somus-green/40 focus:outline-none"
                  >
                    {ASSIGNEES.map((a) => (
                      <option key={a} value={a}>{a}</option>
                    ))}
                  </select>
                </div>
                <div>
                  <label className="block text-sm font-medium text-somus-gray-700 mb-1">Prioridade</label>
                  <select
                    value={formPriority}
                    onChange={(e) => setFormPriority(e.target.value as Priority)}
                    className="w-full text-sm border border-somus-gray-300 rounded-lg px-3 py-2 focus:ring-2 focus:ring-somus-green/40 focus:outline-none"
                  >
                    <option value="baixa">Baixa</option>
                    <option value="media">Media</option>
                    <option value="alta">Alta</option>
                    <option value="urgente">Urgente</option>
                  </select>
                </div>
              </div>

              <div>
                <label className="block text-sm font-medium text-somus-gray-700 mb-1">Data Limite</label>
                <input
                  type="date"
                  value={formDueDate}
                  onChange={(e) => setFormDueDate(e.target.value)}
                  className="w-full text-sm border border-somus-gray-300 rounded-lg px-3 py-2 focus:ring-2 focus:ring-somus-green/40 focus:outline-none"
                />
              </div>
            </div>

            {/* Modal Footer */}
            <div className="flex items-center justify-end gap-3 px-6 py-4 border-t border-somus-gray-200">
              <Button variant="secondary" onClick={() => setShowModal(false)}>
                Cancelar
              </Button>
              <Button variant="primary" onClick={handleSave} disabled={!formTitle}>
                {editingTask ? 'Salvar' : 'Criar'}
              </Button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
